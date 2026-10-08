import os from "os";
import path from "path";
import fs from "fs";
import { renderPptx, type RenderedSlide } from "./render";
import { critiqueSlides, type Finding } from "./critique";

export { type Finding } from "./critique";
export { type RenderedSlide } from "./render";

export type PreflightReport = {
  slides: RenderedSlide[];
  findings: Finding[];
  blockers: number;
  majors: number;
  minors: number;
  /** True when nothing would stop this deck going to a client. */
  passed: boolean;
};

/**
 * Render a built deck and inspect it.
 *
 * This is the quality gate the project has been missing: the deck is only
 * ever as good as the last time somebody looked at it, and nobody looks at
 * 1,000 decks a month. `workDir` defaults to a temp directory; pass one to
 * keep the images for inspection.
 */
export async function preflight(
  pptx: Buffer,
  opts: { workDir?: string; model?: string; dpi?: number } = {}
): Promise<PreflightReport> {
  const workDir = opts.workDir ?? fs.mkdtempSync(path.join(os.tmpdir(), "preflight-"));
  const slides = await renderPptx(pptx, workDir, opts.dpi ?? 100);
  const findings = await critiqueSlides(slides, opts.model);

  const count = (s: Finding["severity"]) => findings.filter((f) => f.severity === s).length;
  const blockers = count("blocker");

  return {
    slides,
    findings,
    blockers,
    majors: count("major"),
    minors: count("minor"),
    passed: blockers === 0,
  };
}

/** The report as a short plain-text summary, for logs and CI output. */
export function formatReport(report: PreflightReport): string {
  const lines = [
    `preflight: ${report.slides.length} slides · ` +
      `${report.blockers} blocker, ${report.majors} major, ${report.minors} minor`,
  ];
  for (const f of report.findings) {
    lines.push(`  [${f.severity}] slide ${f.slide} (${f.rule}): ${f.issue}`);
    lines.push(`      fix: ${f.fix}`);
  }
  if (report.findings.length === 0) lines.push("  nothing found");
  return lines.join("\n");
}
