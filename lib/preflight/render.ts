import { execFile } from "child_process";
import { promisify } from "util";
import fs from "fs";
import path from "path";

const run = promisify(execFile);

/**
 * Turns a built .pptx into one PNG per slide.
 *
 * This is the step Vercel cannot do — it needs LibreOffice — and it is the
 * reason deck generation has to move to a container. Everything the generator
 * has got wrong so far (stretched photographs, a dropped table, a vanished
 * photo grid, randomly coloured charts, a table running off the slide) was
 * invisible to the compiler and obvious in a picture.
 *
 * Binaries are taken from the environment so the same code works in a
 * container, where soffice and pdftoppm are installed, and on a workstation
 * that only has Docker:
 *
 *   SOFFICE_BIN            path to soffice            (default: "soffice")
 *   PDFTOPPM_BIN           path to pdftoppm           (default: "pdftoppm")
 *   GS_BIN                 ghostscript, used only if pdftoppm is absent
 *   PREFLIGHT_SOFFICE_IMAGE  run soffice in this Docker image instead
 */

export type RenderedSlide = { page: number; file: string };

const SOFFICE = process.env.SOFFICE_BIN ?? "soffice";
const PDFTOPPM = process.env.PDFTOPPM_BIN ?? "pdftoppm";
const GS = process.env.GS_BIN ?? "gs";
const DOCKER_IMAGE = process.env.PREFLIGHT_SOFFICE_IMAGE;

async function toPdf(workDir: string, pptxName: string): Promise<string> {
  if (DOCKER_IMAGE) {
    // Workstation path. The container path below is what production uses.
    await run("docker", [
      "run", "--rm",
      "-e", "HOME=/tmp",
      "-v", `${workDir}:/data`,
      "--entrypoint", "/usr/bin/soffice",
      DOCKER_IMAGE,
      "--headless", "--norestore", "--convert-to", "pdf",
      "--outdir", "/data", `/data/${pptxName}`,
    ], { maxBuffer: 1 << 26 });
  } else {
    await run(SOFFICE, [
      "--headless", "--norestore", "--convert-to", "pdf",
      "--outdir", workDir, path.join(workDir, pptxName),
    ], { maxBuffer: 1 << 26 });
  }

  const pdf = path.join(workDir, pptxName.replace(/\.pptx$/i, ".pdf"));
  if (!fs.existsSync(pdf)) {
    throw new Error(
      `LibreOffice produced no PDF for ${pptxName}. Checked ${pdf}. ` +
        `Is soffice installed (SOFFICE_BIN) or PREFLIGHT_SOFFICE_IMAGE set?`
    );
  }
  return pdf;
}

async function toPngs(pdf: string, workDir: string, dpi: number): Promise<void> {
  const stem = path.join(workDir, "slide");
  try {
    await run(PDFTOPPM, ["-png", "-r", String(dpi), pdf, stem], { maxBuffer: 1 << 26 });
    return;
  } catch {
    // Ghostscript is the fallback; it is what a Windows box tends to have.
    await run(GS, [
      "-sDEVICE=png16m", `-r${dpi}`, "-dNOPAUSE", "-dBATCH", "-q",
      `-sOutputFile=${stem}-%02d.png`, pdf,
    ], { maxBuffer: 1 << 26 });
  }
}

/**
 * Renders every slide and returns the images in slide order.
 *
 * `dpi` 100 gives roughly 1333x750 for a 16:9 deck — enough for a model to
 * read axis labels and table text without paying for a full-resolution image.
 */
export async function renderPptx(
  pptx: Buffer,
  workDir: string,
  dpi = 100
): Promise<RenderedSlide[]> {
  fs.mkdirSync(workDir, { recursive: true });
  const pptxName = "deck.pptx";
  fs.writeFileSync(path.join(workDir, pptxName), pptx);

  const pdf = await toPdf(workDir, pptxName);
  await toPngs(pdf, workDir, dpi);

  const slides = fs
    .readdirSync(workDir)
    .filter((f) => /^slide-?\d+\.png$/i.test(f))
    .map((f) => ({ page: parseInt(f.match(/(\d+)\.png$/i)![1], 10), file: path.join(workDir, f) }))
    .sort((a, b) => a.page - b.page);

  if (slides.length === 0) throw new Error(`No slide images were produced in ${workDir}.`);
  return slides;
}
