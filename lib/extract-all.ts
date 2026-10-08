import { extractPptx } from "./parsing/pptx";
import { extractXlsx } from "./parsing/xlsx";
import { extractImages, imageManifest, type ExtractedImage } from "./parsing/images";

function sheetFingerprint(sheets: Record<string, unknown[][]>): string {
  // Two embedded workbooks count as the same for our purposes if they
  // share sheet names and row counts. Confirmed real case: one MMR had
  // several embedded copies of a single master workbook, and without
  // this, every sheet of every copy became its own near-duplicate slide.
  return Object.entries(sheets)
    .map(([name, rows]) => `${name}:${rows.length}`)
    .sort()
    .join("|");
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
        const sheetText = Object.entries(wb.sheets)
          .filter(([, rows]) => rows.length > 1)
          .map(([name, rows]) => `  sheet "${name}":\n` + rows.slice(0, 25).map((r) => "    " + JSON.stringify(r)).join("\n"))
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
