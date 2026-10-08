/**
 * Chart colour, axis and number-format decisions, shared by the .pptx renderer
 * and the browser preview.
 *
 * These used to be written twice, once in each, and drifted: the preview drew
 * a third series in a different colour from the deck, and grouped numbers the
 * Indian way while the deck grouped them the Western way. The submitter
 * approves the preview, so anything the two disagree on is something the
 * submitter never actually saw. One copy, imported by both.
 *
 * Pure arithmetic and strings - safe to import from a client component.
 */

export const FM_BLUE = "264170";
export const SUNSHINE = "F7A328";

/** Mixes a hex colour (no #) toward white. 0 returns it unchanged, 1 returns white. */
export function tint(hex: string, amount: number): string {
  const n = parseInt(hex, 16);
  const mix = (c: number) => Math.round(c + (255 - c) * amount);
  return [mix((n >> 16) & 255), mix((n >> 8) & 255), mix(n & 255)]
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

/**
 * Exactly `n` distinct brand colours (hex, no #), alternating the two brand
 * hues and stepping through tints, so adjacent entries never collide.
 *
 * The LENGTH is the point, not just the hues. pptxgenjs colours pie slices -
 * and the bars of a single-series bar chart - per data point, and once the
 * index passes the end of the palette it fills the rest with
 * `chartColors[Math.floor(Math.random() * chartColors.length)]`
 * (dist/pptxgen.cjs.js, the `<c:dPt>` branches). Verified by generating the
 * same five-slice pie six times against a two-colour palette and getting six
 * different colour assignments. Always pass one colour per point.
 *
 * Six tint steps cover the 12-category ceiling the curation sanitiser enforces. With
 * four, the ninth colour repeated the first, so a 9-12 slice pie had two
 * slices - and two legend entries - in the same blue. The steps are spaced
 * widely because small ones do not separate: at 0.21 a third series came out
 * a slate blue that preflight flagged as reading like the first.
 */
const STEPS = [0, 0.55, 0.3, 0.75, 0.15, 0.88];

export function chartPalette(n: number): string[] {
  const bases = [FM_BLUE, SUNSHINE];
  return Array.from({ length: Math.max(1, n) }, (_, i) =>
    tint(bases[i % bases.length], STEPS[Math.floor(i / bases.length) % STEPS.length])
  );
}

/**
 * Decimal places needed to show every value exactly, capped at 3.
 *
 * The format used to be hard-coded "#,##0", which rounds: 12.5 was charted and
 * labelled as 13, a figure that is not in the source - the one thing this
 * system must never do. Values are normalised through toPrecision first, so
 * spreadsheet float noise (528.0000000000007) counts as 528, not as 13 places.
 */
export function decimalsNeeded(values: number[]): number {
  let dp = 0;
  for (const v of values) {
    if (!Number.isFinite(v)) continue;
    const clean = String(Number(v.toPrecision(12)));
    const frac = clean.includes("e") ? "" : clean.split(".")[1] ?? "";
    dp = Math.max(dp, frac.length);
  }
  return Math.min(3, dp);
}

/** Excel/PowerPoint number format with thousands separators and `dp` decimals. */
export function numberFormatCode(dp: number): string {
  return dp > 0 ? `#,##0.${"0".repeat(dp)}` : "#,##0";
}

/**
 * The same number as numberFormatCode would show it, for the preview.
 *
 * Western grouping, to match what PowerPoint draws from "#,##0". The preview
 * used Indian grouping (14,48,000) while the deck showed 1,448,000; the two
 * have to read identically.
 */
export function formatNumber(v: number, dp: number): string {
  return v.toLocaleString("en-US", { minimumFractionDigits: dp, maximumFractionDigits: dp });
}

/** A round step from the 1-2-5 ladder, aiming for about `target` intervals. */
export function niceStep(span: number, target = 5): number {
  if (!(span > 0) || !Number.isFinite(span)) return 1;
  const raw = span / target;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const nice = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10;
  return nice * mag;
}

export type Axis = { min: number; max: number; step: number };

/**
 * Value-axis bounds and gridline step for a set of values.
 *
 * - The step comes from the 1-2-5 ladder and is applied as the major unit, so
 *   gridlines land on round numbers. Before, the step was a power of ten
 *   halved, which put gridlines on half-units of the data, and the engine
 *   chose its own unit on top of that.
 * - Non-negative data never gets a negative axis.
 * - `includeZero` is for bars, which encode value as length and so must start
 *   at zero - and, when there are negative values, must reach below zero, or a
 *   negative bar has nowhere to be drawn and vanishes.
 * - Without it (lines), the real range is padded, because a line encodes value
 *   by position: left at zero, consumption moving 27,240 to 28,920 drew as a
 *   dead flat rule.
 */
export function niceAxis(values: number[], opts: { includeZero?: boolean; pad?: number } = {}): Axis {
  const finite = values.filter((v) => Number.isFinite(v));
  let lo = finite.length ? Math.min(...finite) : 0;
  let hi = finite.length ? Math.max(...finite) : 1;
  if (opts.includeZero) {
    lo = Math.min(0, lo);
    hi = Math.max(0, hi);
  }
  const span = hi - lo || Math.abs(hi) || 1;
  const pad = span * (opts.pad ?? 0.15);

  let rawMin = opts.includeZero && lo === 0 ? 0 : lo - pad;
  let rawMax = opts.includeZero && hi === 0 ? 0 : hi + pad;
  if (lo >= 0 && rawMin < 0) rawMin = 0;
  if (hi <= 0 && rawMax > 0) rawMax = 0;
  if (rawMax <= rawMin) rawMax = rawMin + 1;

  const step = niceStep(rawMax - rawMin);
  return {
    min: Math.floor(rawMin / step + 1e-9) * step,
    max: Math.ceil(rawMax / step - 1e-9) * step,
    step,
  };
}
