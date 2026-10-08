import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });

// Verified against a real 29-slide MMR: content photography ran 49-184KB,
// while every icon, bullet graphic and embedded-OLE placeholder icon came
// in under 15KB. 30KB sits comfortably between the two clusters.
const MIN_CONTENT_IMAGE_BYTES = 30 * 1024;

// .wmf is vector clipart from old Office builds - not usable as a photo
// and not reliably embeddable, so it's excluded regardless of size.
const USABLE_EXT = /\.(jpe?g|png)$/i;

export type ExtractedImage = {
  id: string;
  slideNumber: number;
  slideTitle: string;
  ext: string;
  bytes: number;
  buffer: Buffer;
};

function asArray<T>(v: T | T[] | undefined): T[] {
  if (v === undefined) return [];
  return Array.isArray(v) ? v : [v];
}

export async function extractImages(
  buffer: Buffer,
  slideTitles: Map<number, string>
): Promise<ExtractedImage[]> {
  const zip = await JSZip.loadAsync(buffer);
  const out: ExtractedImage[] = [];
  let counter = 0;

  const relsFiles = Object.keys(zip.files).filter((n) =>
    /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(n)
  );

  for (const relsName of relsFiles) {
    const slideNumber = parseInt(relsName.match(/slide(\d+)\.xml\.rels/)![1], 10);
    const relsXml = await zip.file(relsName)!.async("string");
    const parsed = parser.parse(relsXml);
    const rels = asArray(parsed?.Relationships?.Relationship);

    for (const rel of rels) {
      const target = (rel as Record<string, unknown>)["@_Target"] as string | undefined;
      if (!target || !target.includes("/media/")) continue;
      if (!USABLE_EXT.test(target)) continue;

      const mediaPath = "ppt/" + target.replace(/^\.\.\//, "");
      const file = zip.file(mediaPath);
      if (!file) continue;

      const imgBuffer = await file.async("nodebuffer");
      if (imgBuffer.length < MIN_CONTENT_IMAGE_BYTES) continue;

      const ext = (mediaPath.match(USABLE_EXT)?.[0] ?? ".jpg").toLowerCase();
      out.push({
        id: `img${++counter}`,
        slideNumber,
        slideTitle: slideTitles.get(slideNumber) ?? "",
        ext: ext === ".jpeg" ? ".jpg" : ext,
        bytes: imgBuffer.length,
        buffer: imgBuffer,
      });
    }
  }

  return out.sort((a, b) => a.slideNumber - b.slideNumber);
}

/**
 * A compact text manifest of the images, for Claude. Claude doesn't need
 * the pixels to place these sensibly - the slide title it came from is
 * what identifies the subject ("8 photos from a Pest Control Activity
 * slide"). Sending 40+ real photos into the prompt would be slow and
 * expensive for no real gain in placement quality.
 */
export function imageManifest(images: ExtractedImage[]): string {
  if (images.length === 0) return "(no content photography found in the source files)";
  const bySlide = new Map<number, ExtractedImage[]>();
  for (const img of images) {
    if (!bySlide.has(img.slideNumber)) bySlide.set(img.slideNumber, []);
    bySlide.get(img.slideNumber)!.push(img);
  }
  const lines: string[] = [];
  for (const [slideNumber, imgs] of [...bySlide.entries()].sort((a, b) => a[0] - b[0])) {
    lines.push(
      `- Source slide ${slideNumber} ("${imgs[0].slideTitle}"): ${imgs.length} photo(s) — ids: ${imgs.map((i) => i.id).join(", ")}`
    );
  }
  return lines.join("\n");
}
