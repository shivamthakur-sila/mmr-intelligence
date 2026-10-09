import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import * as XLSX from "xlsx";
import { posix } from "path";
import { readWorkbookRows } from "./xlsx";

// Text stays text, exactly as written. By default the parser turns a run that
// is only a number into a number and trims a run that is only a space, and
// collectText keeps strings: every numeric table cell on a slide - a Sr. No.
// column, a kWh reading, a headcount - reached the model empty, and a chart
// title written in runs came out as "ClientComplaints".
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  trimValues: false,
});

/**
 * Recursively walks a parsed slide-XML tree collecting every <a:t> text run,
 * regardless of nesting depth. This is what makes grouped shapes (org
 * charts, etc.) work correctly without any special-case code - a group's
 * nested shapes are just deeper in the same object tree, and a generic
 * walk finds them the same as top-level ones. Confirmed against a real
 * grouped org chart during testing.
 */
function collectText(node: unknown, out: string[]): void {
  if (node == null) return;
  if (Array.isArray(node)) {
    for (const item of node) collectText(item, out);
    return;
  }
  if (typeof node === "object") {
    const obj = node as Record<string, unknown>;
    if (typeof obj["a:t"] === "string") {
      out.push(obj["a:t"]);
    } else if (Array.isArray(obj["a:t"])) {
      for (const t of obj["a:t"]) if (typeof t === "string") out.push(t);
    }
    for (const key of Object.keys(obj)) {
      if (key === "a:t") continue;
      collectText(obj[key], out);
    }
  }
}

/** Extracts native PPTX tables (a:tbl elements) as row/cell structures. */
function collectTables(node: unknown, out: string[][][]): void {
  if (node == null) return;
  if (Array.isArray(node)) {
    for (const item of node) collectTables(item, out);
    return;
  }
  if (typeof node === "object") {
    const obj = node as Record<string, unknown>;
    if (obj["a:tbl"]) {
      const rows = asArray(getPath(obj["a:tbl"], ["a:tr"]));
      const table: string[][] = rows.map((row) => {
        const cells = asArray(getPath(row, ["a:tc"]));
        return cells.map((cell) => {
          const texts: string[] = [];
          collectText(cell, texts);
          return texts.join(" ").replace(/\s+/g, " ").trim();
        });
      });
      if (table.length) out.push(table);
    }
    for (const key of Object.keys(obj)) {
      if (key === "a:tbl") continue;
      collectTables(obj[key], out);
    }
  }
}

function asArray<T>(v: T | T[] | undefined): T[] {
  if (v === undefined) return [];
  return Array.isArray(v) ? v : [v];
}

function getPath(obj: unknown, path: string[]): unknown {
  let cur = obj;
  for (const p of path) {
    if (cur == null || typeof cur !== "object") return undefined;
    cur = (cur as Record<string, unknown>)[p];
  }
  return cur;
}

/** Removes rows that are just leftover formula artifacts (errors, stray zeros) rather than real data. */
function isRealRow(row: unknown[]): boolean {
  const meaningful = row.filter((c) => c !== null && c !== undefined && c !== "" && c !== 0);
  if (meaningful.length === 0) return false;
  if (meaningful.every((c) => typeof c === "string" && c.startsWith("#"))) return false;
  return true;
}

/** The relationships a part declares, with targets resolved to package paths. */
async function relationships(zip: JSZip, part: string): Promise<{ type: string; target: string }[]> {
  const dir = posix.dirname(part);
  const rels = zip.file(`${dir}/_rels/${posix.basename(part)}.rels`);
  if (!rels) return [];
  const parsed = parser.parse(await rels.async("string"));
  return asArray(getPath(parsed, ["Relationships", "Relationship"])).flatMap((r) => {
    const rel = r as Record<string, unknown>;
    const type = rel["@_Type"];
    const target = rel["@_Target"];
    if (typeof type !== "string" || typeof target !== "string" || rel["@_TargetMode"] === "External") return [];
    // A target is relative to the part's folder, or rooted at the package.
    const resolved = target.startsWith("/") ? target.slice(1) : posix.join(dir, target);
    return [{ type, target: posix.normalize(resolved) }];
  });
}

export type EmbeddedWorkbook = {
  fileName: string;
  sheets: Record<string, unknown[][]>;
  /** The slide that shows this workbook, when one does. */
  slideNumber?: number;
  /** The title of the chart it feeds, when it feeds a titled chart. */
  chartTitle?: string;
};

export type ExtractedSlide = {
  slideNumber: number;
  text: string[];
  tables: string[][][];
  hasEmbeddedObject: boolean;
};

export async function extractPptx(buffer: Buffer): Promise<{
  slides: ExtractedSlide[];
  embeddedWorkbooks: EmbeddedWorkbook[];
  /** Slides plus every embedded workbook, undeduplicated. The n8n contract. */
  summary: string;
  /** Slides only, for callers that render the workbooks themselves. */
  slideSummary: string;
}> {
  const zip = await JSZip.loadAsync(buffer);

  const slideFiles = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)\.xml/)![1], 10);
      const nb = parseInt(b.match(/slide(\d+)\.xml/)![1], 10);
      return na - nb;
    });

  const slides: ExtractedSlide[] = [];

  for (const slideFile of slideFiles) {
    const slideNumber = parseInt(slideFile.match(/slide(\d+)\.xml/)![1], 10);
    const xml = await zip.file(slideFile)!.async("string");
    const parsed = parser.parse(xml);

    const runs: string[] = [];
    collectText(parsed, runs);
    // Runs are kept untrimmed for chart titles; slide text is read as words.
    const text = runs.map((t) => t.trim()).filter(Boolean);

    const tables: string[][][] = [];
    collectTables(parsed, tables);

    // A slide with an embedded OLE object references it via a relationship
    // of type "oleObject", pointing at a .bin file rather than an .xlsx. The
    // flag marks the slide; the workbook inside the .bin is read at the
    // package level below.
    let hasEmbeddedObject = false;
    const relsPath = `ppt/slides/_rels/${slideFile.split("/").pop()}.rels`;
    const relsFile = zip.file(relsPath);
    if (relsFile) {
      const relsXml = await relsFile.async("string");
      const relsParsed = parser.parse(relsXml);
      const relationships = asArray(getPath(relsParsed, ["Relationships", "Relationship"]));
      hasEmbeddedObject = relationships.some((rel) => {
        const type = (rel as Record<string, unknown>)["@_Type"] as string | undefined;
        return !!type && type.includes("oleObject");
      });
    }

    slides.push({ slideNumber, text, tables, hasEmbeddedObject });
  }

  // Which slide and chart each embedded workbook belongs to. Without this the
  // model saw several blocks with identical headers ("Total Complaints") and
  // nothing to say that one backs "Client Complaints - YTD" and another "Self
  // Generated Request - MTD"; which was which rested on zip order, which is
  // not slide order. A chart reaches its workbook through two relationships,
  // slide -> chart -> package; an OLE object is referenced by the slide itself.
  const origin = new Map<string, { slideNumber: number; chartTitle?: string }>();
  for (const s of slides) {
    const slidePath = `ppt/slides/slide${s.slideNumber}.xml`;
    for (const rel of await relationships(zip, slidePath)) {
      if (/\/oleObject$/.test(rel.type)) {
        origin.set(rel.target, { slideNumber: s.slideNumber });
      } else if (/\/chart$/.test(rel.type)) {
        const chartFile = zip.file(rel.target);
        const titleTexts: string[] = [];
        if (chartFile) collectText(getPath(parser.parse(await chartFile.async("string")), ["c:chartSpace", "c:chart", "c:title"]), titleTexts);
        const chartTitle = titleTexts.join("").trim() || undefined;
        for (const inner of await relationships(zip, rel.target)) {
          if (/\/package$/.test(inner.type)) origin.set(inner.target, { slideNumber: s.slideNumber, chartTitle });
        }
      }
    }
  }

  // Package-level pass: every embedded workbook, wherever it sits. Real testing
  // showed these are often multiple copies of the same master workbook (from
  // repeated Insert > Object actions against one source file), so all copies
  // are surfaced here and extract-all removes exact duplicates.
  //
  // An object inserted with Insert > Object is stored as oleObjectN.bin, an OLE
  // container whose "Package" stream is the complete .xlsx. There is not always
  // a readable .xlsx copy elsewhere in the package: in a real MMR, the asset
  // list and PPM schedule existed only inside .bin files and never reached the
  // model, so their 2025 dates could not be checked against the report month.
  const embeddedWorkbooks: EmbeddedWorkbook[] = [];
  const embeddingFiles = Object.keys(zip.files).filter((f) => /^ppt\/embeddings\/.*\.(xlsx|bin)$/i.test(f));
  for (const path of embeddingFiles) {
    let embedBuffer: Buffer = await zip.file(path)!.async("nodebuffer");
    try {
      if (/\.bin$/i.test(path)) {
        const pkg = XLSX.CFB.find(XLSX.CFB.read(embedBuffer, { type: "buffer" }), "Package");
        // Only a zip (an .xlsx) is a workbook; a .bin can equally hold a
        // Word document or a PDF, which are not read here.
        if (!pkg?.content || pkg.content.length < 2 || pkg.content[0] !== 0x50 || pkg.content[1] !== 0x4b) continue;
        embedBuffer = Buffer.from(pkg.content as Uint8Array);
      }
      // Same reader as standalone workbooks, so date cells arrive as
      // "Apr 2025" rather than Excel's serial 45748 here too.
      const sheets: Record<string, unknown[][]> = {};
      for (const [sheetName, raw] of Object.entries(readWorkbookRows(embedBuffer))) {
        sheets[sheetName] = raw.filter(isRealRow);
      }
      embeddedWorkbooks.push({ fileName: path.split("/").pop()!, sheets, ...origin.get(path) });
    } catch {
      // Not every embedded file is a valid workbook; skip anything that
      // fails to parse rather than crash the whole extraction.
    }
  }

  const summary = slides
    .map((s) => {
      const parts = [`--- Slide ${s.slideNumber} ---`];
      if (s.text.length) parts.push(s.text.join(" "));
      for (const t of s.tables) {
        parts.push(`[TABLE ${t.length}x${t[0]?.length ?? 0}]`);
        for (const row of t) parts.push(row.join(" | "));
      }
      if (s.hasEmbeddedObject) parts.push("[SLIDE REFERENCES AN EMBEDDED OBJECT - see package-level embedded workbooks]");
      if (parts.length === 1) parts.push("(no content on this slide)");
      return parts.join("\n");
    })
    .join("\n\n");

  const embeddedSummary = embeddedWorkbooks
    .map((ew) => {
      const parts = [`[EMBEDDED WORKBOOK: ${ew.fileName}]`];
      for (const [sheetName, rows] of Object.entries(ew.sheets)) {
        parts.push(`  sheet "${sheetName}": ${rows.length} real data rows`);
        for (const row of rows.slice(0, 8)) parts.push("  " + JSON.stringify(row));
      }
      return parts.join("\n");
    })
    .join("\n\n");

  // `summary` keeps both halves because app/api/parse-submission returns it
  // straight to n8n, which expects the workbook content in it — changing that
  // would break a workflow in another repo with no compile error to warn us.
  // `slideSummary` is the slides alone, for callers that render the embedded
  // workbooks themselves and would otherwise emit them twice.
  return {
    slides,
    embeddedWorkbooks,
    summary: summary + "\n\n" + embeddedSummary,
    slideSummary: summary,
  };
}
