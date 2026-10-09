import { createHash } from "node:crypto";
import { extractPptx } from "./parsing/pptx";
import { extractXlsx, sheetText, workbookText } from "./parsing/xlsx";
import { extractImages, imageManifest, type ExtractedImage } from "./parsing/images";

function sheetFingerprint(sheets: Record<string, unknown[][]>): string {
  // Dedup exists because one MMR had several embedded copies of a single
  // master workbook, and without it every sheet of every copy became its own
  // near-duplicate slide. It hashes the content, not the shape: an earlier
  // fingerprint of sheet names and row counts collapsed different workbooks
  // that merely had the same layout - a real MMR had four "Sheet1", 14-row
  // workbooks holding four different datasets, and three never reached the
  // model. Sheet names are sorted so only a true copy matches.
  const ordered = Object.keys(sheets)
    .sort()
    .map((name) => [name, sheets[name]]);
  return createHash("sha1").update(JSON.stringify(ordered)).digest("hex");
}

export async function extractAllFiles(files: { buffer: Buffer; name: string }[]) {
  const textParts: string[] = [];
  const allImages: ExtractedImage[] = [];
  const seenFingerprints = new Set<string>();
  let imageIdOffset = 0;

  for (const file of files) {
    const lower = file.name.toLowerCase();

    if (lower.endsWith(".pptx")) {
      // Slides only. `summary` would carry every embedded workbook as well,
      // undeduplicated, and the loop below adds them again deduplicated — one
      // real submission reached Claude with the same workbook five times over,
      // each copy truncated at a different row, for ~35k characters of pure
      // repetition in every prompt.
      const { slides, embeddedWorkbooks, slideSummary } = await extractPptx(file.buffer);
      textParts.push(`=== Source file: ${file.name} ===\n${slideSummary}`);

      const titles = new Map(slides.map((s) => [s.slideNumber, s.text[0] || ""]));
      const images = await extractImages(file.buffer, titles);
      // Keep ids unique across multiple uploaded files
      for (const img of images) {
        allImages.push({ ...img, id: `img${++imageIdOffset}` });
      }

      // Copies of one workbook are emitted once, named by every chart and
      // slide that uses them: five charts in one real MMR were backed by the
      // same master workbook, and labelling it with whichever copy came first
      // tied it to one chart and hid that the other four drew on it too.
      const groups = new Map<string, typeof embeddedWorkbooks>();
      for (const wb of embeddedWorkbooks) {
        const fp = sheetFingerprint(wb.sheets);
        if (seenFingerprints.has(fp)) continue; // already emitted from an earlier file
        groups.set(fp, [...(groups.get(fp) ?? []), wb]);
      }
      for (const [fp, copies] of groups) {
        seenFingerprints.add(fp);
        // A sheet the model is not shown in full says so, so that it never
        // counts or totals part of a list as if it were the whole.
        const sheetsText = Object.entries(copies[0].sheets)
          .filter(([, rows]) => rows.length > 0)
          .map(([name, rows]) => sheetText(`  sheet "${name}":`, rows, "    "))
          .join("\n");
        if (!sheetsText) continue;
        const uses = [
          ...new Set(
            copies
              .filter((c) => c.slideNumber !== undefined)
              .sort((a, b) => a.slideNumber! - b.slideNumber!)
              .map((c) => (c.chartTitle ? `chart "${c.chartTitle}" on slide ${c.slideNumber}` : `slide ${c.slideNumber}`))
          ),
        ];
        textParts.push(
          `=== Embedded workbook in ${file.name} (${copies[0].fileName})` +
            `${uses.length > 0 ? ` - used by ${uses.join("; ")}` : ""} ===\n${sheetsText}`
        );
      }
    } else if (lower.endsWith(".xlsx")) {
      const { sheets } = extractXlsx(file.buffer);
      textParts.push(`=== Source file: ${file.name} ===\n${workbookText(sheets)}`);
    }
  }

  return {
    combinedText: textParts.join("\n\n"),
    images: allImages,
    imageManifest: imageManifest(allImages),
  };
}
