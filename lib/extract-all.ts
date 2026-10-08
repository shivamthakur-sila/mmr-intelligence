import { createHash } from "node:crypto";
import { extractPptx } from "./parsing/pptx";
import { extractXlsx } from "./parsing/xlsx";
import { extractImages, imageManifest, type ExtractedImage } from "./parsing/images";

/** Rows of an embedded sheet shown to the model before the rest are summarised as a count. */
const EMBEDDED_SHEET_ROW_LIMIT = 25;

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

      for (const wb of embeddedWorkbooks) {
        const fp = sheetFingerprint(wb.sheets);
        if (seenFingerprints.has(fp)) continue;
        seenFingerprints.add(fp);
        // Each sheet carries its real row count and says when it was cut, as
        // standalone workbooks already do; otherwise the model cannot tell a
        // truncated list from a complete one and may count or sum a part of it.
        const sheetText = Object.entries(wb.sheets)
          .filter(([, rows]) => rows.length > 0)
          .map(([name, rows]) => {
            const lines = [`  sheet "${name}" (${rows.length} real data rows):`];
            for (const r of rows.slice(0, EMBEDDED_SHEET_ROW_LIMIT)) lines.push("    " + JSON.stringify(r));
            if (rows.length > EMBEDDED_SHEET_ROW_LIMIT) {
              lines.push(`    ... ${rows.length - EMBEDDED_SHEET_ROW_LIMIT} more rows not shown`);
            }
            return lines.join("\n");
          })
          .join("\n");
        if (sheetText) {
          textParts.push(`=== Embedded workbook in ${file.name} (${wb.fileName}) ===\n${sheetText}`);
        }
      }
    } else if (lower.endsWith(".xlsx")) {
      const { summary } = extractXlsx(file.buffer);
      textParts.push(`=== Source file: ${file.name} ===\n${summary}`);
    }
  }

  return {
    combinedText: textParts.join("\n\n"),
    images: allImages,
    imageManifest: imageManifest(allImages),
  };
}
