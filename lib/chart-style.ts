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

/** Mixes a hex colour (no #) toward black. 0 returns it unchanged. */
export function shade(hex: string, amount: number): string {
  const n = parseInt(hex, 16);
  const mix = (c: number) => Math.round(c * (1 - amount));
  return [mix((n >> 16) & 255), mix((n >> 8) & 255), mix(n & 255)]
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

const CHARCOAL = "3C3C3B";

/**
 * Twelve colours for the 12-category ceiling the curation sanitiser enforces,
 * led by the two brand hues and built only from them and a neutral grey.
 *
 * Chosen by measured perceptual distance, not by being different hex codes.
 * Tints of two hues alone cannot make twelve colours a reader can tell apart:
 * the previous steps gave a 10-slice pie two oranges 4 apart in CIEDE2000
 * (Plumbing and Fire Safety read as one colour on the rendered legend) and two
 * near-white slices that vanished against the slide. Darker shades and a grey
 * add the separation. Every pair below is at least 10.8 apart and every
 * colour at least 15 from white; the first four are at least 12 apart, so a
 * small chart gets the most distinct colours. Searched greedily from a fixed
 * brand-led start, then checked by rendering a 12-slice pie.
 */
const PALETTE = [
  FM_BLUE,
  SUNSHINE,
  tint(FM_BLUE, 0.55), // 9DAABF
  tint(CHARCOAL, 0.45), // 949493
  shade(SUNSHINE, 0.45), // 885A16
  tint(SUNSHINE, 0.7), // FDE3BF
  tint(CHARCOAL, 0.15), // 595958
  tint(FM_BLUE, 0.25), // 5C7194
  shade(SUNSHINE, 0.25), // B97A1E
  tint(CHARCOAL, 0.65), // BBBBBA
  shade(FM_BLUE, 0.45), // 15243E
  tint(SUNSHINE, 0.4), // FAC87E
];

/**
 * Exactly `n` brand colours (hex, no #), one per data point.
 *
 * The LENGTH is the point, not just the hues. pptxgenjs colours pie slices -
 * and the bars of a single-series bar chart - per data point, and once the
 * index passes the end of the palette it fills the rest with
 * `chartColors[Math.floor(Math.random() * chartColors.length)]`
 * (dist/pptxgen.cjs.js, the `<c:dPt>` branches). Verified by generating the
 * same five-slice pie six times against a two-colour palette and getting six
 * different colour assignments. Always pass one colour per point. Past twelve
 * - which the sanitiser never lets through - the list repeats.
 */
export function chartPalette(n: number): string[] {
  return Array.from({ length: Math.max(1, n) }, (_, i) => PALETTE[i % PALETTE.length]);
}

/**
 * Decimal places needed to show every value exactly, up to 6.
 *
 * The format used to be hard-coded "#,##0", which rounds: 12.5 was charted and
 * labelled as 13, a figure that is not in the source - the one thing this
 * system must never do. A cap of 3 did the same to a power factor of 0.9995,
 * drawn as unity, "1.000". Values are normalised through toPrecision first,
 * so spreadsheet float noise (528.0000000000007) counts as 528, not as 13
 * places; the cap only stops a value like 46890.41095890411 filling a label.
 */
export function decimalsNeeded(values: number[]): number {
  let dp = 0;
  for (const v of values) {
    if (!Number.isFinite(v)) continue;
    const clean = String(Number(v.toPrecision(12)));
    const frac = clean.includes("e") ? "" : clean.split(".")[1] ?? "";
    dp = Math.max(dp, frac.length);
  }
  return Math.min(6, dp);
}

/**
 * Decimal places for a value axis's labels: enough to tell every gridline
 * from its neighbours. Taken from the step, not the data. With the data's
 * places, gridlines 0.0005 apart were all labelled to 3 places, so an axis
 * read 0.995, 0.995, 0.996, 0.996, 0.997 - duplicate labels, and a top label
 * that is not a figure anybody submitted.
 */
export function axisDecimals(axis: Axis): number {
  return decimalsNeeded([axis.min, axis.step, axis.max]);
}

/**
 * Excel/PowerPoint number format with Indian digit grouping (10,03,500) and
 * `dp` decimals, the way SILA's clients read figures. Excel has no grouping
 * code for lakhs and crores, so magnitude conditions pick the comma pattern;
 * the last section covers everything under a lakh, where both systems agree.
 */
export function numberFormatCode(dp: number): string {
  const d = dp > 0 ? `.${"0".repeat(dp)}` : "";
  return `[>=10000000]##\\,##\\,##\\,##0${d};[>=100000]##\\,##\\,##0${d};#,##0${d}`;
}

/**
 * The same number as numberFormatCode would show it, for the preview. Both
 * must read identically: the submitter approves the preview.
 */
export function formatNumber(v: number, dp: number): string {
  return v.toLocaleString("en-IN", { minimumFractionDigits: dp, maximumFractionDigits: dp });
}

/** A run of digits as Indian grouping: 1003500 -> 10,03,500. Decimals kept. */
function groupIndian(digits: string, decimals = ""): string {
  const last3 = digits.slice(-3);
  const rest = digits.slice(0, -3).replace(/\B(?=(\d{2})+(?!\d))/g, ",");
  return (rest ? `${rest},${last3}` : last3) + decimals;
}

/**
 * Regroups the large figures in a piece of display text the Indian way:
 * "1003500" and "1,003,500" both become "10,03,500". Only figures of five
 * digits or more change - a year or a four-digit count reads the same either
 * way and is left as typed. Display only: the figures themselves, and what
 * the grounding check compares, are untouched.
 */
export function indianFigures(text: string): string {
  return text.replace(
    /(?<![\d.,])(\d{1,3}(?:,\d{3})+|\d{5,})(\.\d+)?(?![\d,])/g,
    (_, whole: string, dec: string | undefined) => {
      const digits = whole.replace(/,/g, "");
      if (digits.length < 5 || /^0/.test(digits)) return whole + (dec ?? "");
      return groupIndian(digits, dec ?? "");
    }
  );
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

  // Counts are whole numbers, so their gridlines are too: incidents of
  // [0, 1, 0] read oddly against an axis of 0.5 and 1.5.
  const whole = finite.length > 0 && finite.every((v) => Number.isInteger(v));
  const step = whole ? Math.max(1, niceStep(rawMax - rawMin)) : niceStep(rawMax - rawMin);
  return {
    min: Math.floor(rawMin / step + 1e-9) * step,
    max: Math.ceil(rawMax / step - 1e-9) * step,
    step,
  };
}
