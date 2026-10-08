import { spawn, type ChildProcess } from "child_process";
import { randomBytes } from "crypto";
import { pathToFileURL } from "url";
import fs from "fs";
import path from "path";

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
 *   PREFLIGHT_TIMEOUT_MS   limit on each external step (default: 120000)
 */

export type RenderedSlide = { page: number; file: string };

const SOFFICE = process.env.SOFFICE_BIN ?? "soffice";
const PDFTOPPM = process.env.PDFTOPPM_BIN ?? "pdftoppm";
const GS = process.env.GS_BIN ?? "gs";
const DOCKER_IMAGE = process.env.PREFLIGHT_SOFFICE_IMAGE;

// A hung LibreOffice never exits on its own, so without a limit one bad deck
// would hold its request open forever. Two minutes is several times what a
// large real submission takes to convert.
const TIMEOUT_MS = Number(process.env.PREFLIGHT_TIMEOUT_MS) > 0
  ? Number(process.env.PREFLIGHT_TIMEOUT_MS)
  : 120_000;

const SLIDE_PNG = /^slide-?\d+\.png$/i;

/** A failed external step. `notInstalled` is set only when the binary itself
 *  could not be found, which is the one failure that justifies a fallback. */
class StepError extends Error {
  constructor(message: string, readonly notInstalled = false) {
    super(message);
  }
}

/** The end of a tool's output, which is where it says what went wrong. */
function tail(output: string): string {
  const text = output.trim();
  if (!text) return "(no output)";
  return text.length > 2000 ? `...${text.slice(-2000)}` : text;
}

/**
 * Kills a step and everything it started. soffice is a launcher that starts
 * soffice.bin as its own child, and that child inherits the output pipes, so
 * killing only the launcher would leave the conversion running and the pipes
 * open. On POSIX each step runs in its own process group so the whole group
 * can be killed; on Windows taskkill /T does the same for the process tree.
 */
function killTree(child: ChildProcess): void {
  if (child.pid === undefined) return;
  if (process.platform === "win32") {
    spawn("taskkill", ["/pid", String(child.pid), "/T", "/F"], { windowsHide: true, stdio: "ignore" })
      .on("error", () => child.kill("SIGKILL"));
    return;
  }
  try {
    process.kill(-child.pid, "SIGKILL");
  } catch {
    child.kill("SIGKILL");
  }
}

/**
 * Runs one external step to completion, or fails with the step's name and
 * what the tool printed. It settles as soon as the time limit passes rather
 * than waiting for the pipes to close, because a surviving grandchild can
 * hold them open indefinitely.
 */
function run(step: string, file: string, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(file, args, {
      detached: process.platform !== "win32",
      windowsHide: true,
      stdio: ["ignore", "pipe", "pipe"],
    });

    // LibreOffice reports some failures on stdout, so both streams are kept,
    // capped so a chatty tool cannot exhaust memory.
    let output = "";
    const collect = (chunk: Buffer) => {
      if (output.length < 1 << 16) output += chunk.toString();
    };
    child.stdout!.on("data", collect);
    child.stderr!.on("data", collect);

    let settled = false;
    const finish = (err?: StepError) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (err) reject(err);
      else resolve();
    };

    const timer = setTimeout(() => {
      killTree(child);
      child.stdout!.destroy();
      child.stderr!.destroy();
      finish(new StepError(
        `${step} timed out after ${TIMEOUT_MS}ms and was killed (PREFLIGHT_TIMEOUT_MS). ` +
          `Output: ${tail(output)}`
      ));
    }, TIMEOUT_MS);

    child.on("error", (e: NodeJS.ErrnoException) => {
      finish(new StepError(`${step} could not be started (${file}): ${e.message}`, e.code === "ENOENT"));
    });
    child.on("close", (code, signal) => {
      if (code === 0) finish();
      else finish(new StepError(
        `${step} failed (${code !== null ? `exit ${code}` : `signal ${signal}`}): ${tail(output)}`
      ));
    });
  });
}

/**
 * Removes everything an earlier render left in `workDir`. Without this, a
 * reused directory returns the previous deck's slides beyond this deck's
 * length, and a stale PDF would satisfy the check that soffice produced one.
 * The LibreOffice profile goes too, because a profile left by a killed
 * conversion can still carry its lock.
 */
function clearOutput(workDir: string): void {
  for (const f of fs.readdirSync(workDir)) {
    if (SLIDE_PNG.test(f) || f === "deck.pdf") fs.rmSync(path.join(workDir, f), { force: true });
  }
  fs.rmSync(path.join(workDir, "lo-profile"), { recursive: true, force: true });
}

function clearSlides(workDir: string): void {
  for (const f of fs.readdirSync(workDir)) {
    if (SLIDE_PNG.test(f)) fs.rmSync(path.join(workDir, f), { force: true });
  }
}

async function toPdf(workDir: string, pptxName: string): Promise<string> {
  if (DOCKER_IMAGE) {
    // Workstation path. The container path below is what production uses.
    // Killing the docker client does not stop the container it started, so
    // the container is named and removed explicitly if the step fails.
    const name = `preflight-${randomBytes(6).toString("hex")}`;
    try {
      await run("soffice (docker)", "docker", [
        "run", "--rm", "--name", name,
        "-e", "HOME=/tmp",
        "-v", `${workDir}:/data`,
        "--entrypoint", "/usr/bin/soffice",
        DOCKER_IMAGE,
        "--headless", "--norestore", "--convert-to", "pdf",
        "--outdir", "/data", `/data/${pptxName}`,
      ]);
    } catch (e) {
      await run("docker rm", "docker", ["rm", "-f", name]).catch(() => {});
      throw e;
    }
  } else {
    // LibreOffice allows one running instance per user profile, so two
    // conversions sharing the default profile under $HOME collide and one of
    // them fails. Each call gets its own profile inside its own workDir. The
    // value must be a file URL, which pathToFileURL builds correctly for
    // Windows paths as well.
    const profile = pathToFileURL(path.join(workDir, "lo-profile")).href;
    await run("soffice", SOFFICE, [
      `-env:UserInstallation=${profile}`,
      "--headless", "--norestore", "--convert-to", "pdf",
      "--outdir", workDir, path.join(workDir, pptxName),
    ]);
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
    await run("pdftoppm", PDFTOPPM, ["-png", "-r", String(dpi), pdf, stem]);
    return;
  } catch (e) {
    // Only a missing pdftoppm justifies the fallback. If pdftoppm ran and
    // failed, its own error is the one that explains the problem, and
    // Ghostscript failing on the same PDF would only hide it.
    if (!(e instanceof StepError && e.notInstalled)) throw e;
  }

  // Ghostscript is the fallback; it is what a Windows box tends to have.
  clearSlides(workDir);
  try {
    await run("Ghostscript", GS, [
      "-sDEVICE=png16m", `-r${dpi}`, "-dNOPAUSE", "-dBATCH", "-q",
      `-sOutputFile=${stem}-%02d.png`, pdf,
    ]);
  } catch (e) {
    throw new Error(
      `pdftoppm is not installed (PDFTOPPM_BIN=${PDFTOPPM}), and the Ghostscript fallback ` +
        `also failed: ${(e as Error).message}`
    );
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
  clearOutput(workDir);
  const pptxName = "deck.pptx";
  fs.writeFileSync(path.join(workDir, pptxName), pptx);

  const pdf = await toPdf(workDir, pptxName);
  await toPngs(pdf, workDir, dpi);

  const slides = fs
    .readdirSync(workDir)
    .filter((f) => SLIDE_PNG.test(f))
    .map((f) => ({ page: parseInt(f.match(/(\d+)\.png$/i)![1], 10), file: path.join(workDir, f) }))
    .sort((a, b) => a.page - b.page);

  if (slides.length === 0) throw new Error(`No slide images were produced in ${workDir}.`);
  return slides;
}
