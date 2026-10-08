import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { readWorkbookRows } from "./xlsx";

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });

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
          return texts.join(" ").trim();
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

export type ExtractedSlide = {
  slideNumber: number;
  text: string[];
  tables: string[][][];
  hasEmbeddedObject: boolean;
};

export async function extractPptx(buffer: Buffer): Promise<{
  slides: ExtractedSlide[];
  embeddedWorkbooks: { fileName: string; sheets: Record<string, unknown[][]> }[];
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

    const text: string[] = [];
    collectText(parsed, text);

    const tables: string[][][] = [];
    collectTables(parsed, tables);

    // A slide with an embedded OLE object references it via a relationship
    // of type "oleObject", pointing at a .bin file - NOT directly at a
    // .xlsx, even though a readable .xlsx copy exists elsewhere in the
    // package. Precisely pairing a specific .bin to its exact .xlsx sibling
    // requires parsing the compound-binary OLE format itself; the pragmatic
    // and still-correct approach (confirmed against a real file) is to flag
    // which slides reference *an* embedded object, and separately pull all
    // embedded .xlsx files at the package level below.
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

  // Package-level pass: every embedded .xlsx, wherever it sits. Real testing
  // showed these are often multiple copies of the same master workbook
  // (from repeated Insert > Object actions against one source file), so
  // de-duplicating by content isn't attempted here - all copies are surfaced,
  // and Agent 1's prompt is expected to reconcile overlapping data.
  const embeddedWorkbooks: { fileName: string; sheets: Record<string, unknown[][]> }[] = [];
  const embeddingFiles = Object.keys(zip.files).filter((f) => /^ppt\/embeddings\/.*\.xlsx$/i.test(f));
  for (const path of embeddingFiles) {
    const embedBuffer = await zip.file(path)!.async("nodebuffer");
    try {
      // Same reader as standalone workbooks, so date cells arrive as
      // "Apr 2025" rather than Excel's serial 45748 here too.
      const sheets: Record<string, unknown[][]> = {};
      for (const [sheetName, raw] of Object.entries(readWorkbookRows(embedBuffer))) {
        sheets[sheetName] = raw.filter(isRealRow);
      }
      embeddedWorkbooks.push({ fileName: path.split("/").pop()!, sheets });
    } catch {
      // Not all files matching *.xlsx in this folder are necessarily valid
      // standalone workbooks - skip anything that fails to parse rather
      // than crash the whole extraction.
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
