import type { Block } from "./curate";

/**
 * How much room each block needs, how blocks pair up across a slide, where a
 * section has to break onto a second slide, and which slides the deck has.
 *
 * This lives apart from the renderer because the browser preview has to reach
 * the same answer. When only the .pptx knew how to split a section, the
 * preview showed on one slide what the deck spread over two, and a chart the
 * submitter approved as a legible panel was filed as a sliver. Both measure
 * with this.
 *
 * Every text measurement here was taken off a render made with the fonts the
 * deck is drawn in (render/Dockerfile): an earlier calibration was made
 * against substitute fonts 10-15% wider, and sections spilled onto
 * near-empty continuation slides because of it.
 *
 * Everything here is pure arithmetic in inches, safe to import from a client
 * component.
 */

type Table = Extract<Block, { type: "table" }>;

// Geometry of the content area, shared with the renderer. It starts under
// the title's orange rule and stops above the page-number square.
export const CONTENT_W = 12.13;
export const CONTENT_TOP = 1.95;
export const CONTENT_BOTTOM = 6.7;
export const CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP;

// A chart below about an inch and a half has colliding axis labels; above
// three and a half it swamps whatever else is on the slide.
export const CHART_MIN_H = 1.5;
export const CHART_MAX_H = 3.4;
export const BLOCK_SPACING = 0.22;

/**
 * How far a page's estimated height may run past the content area before the
 * planner moves a row to the next slide. The estimates err slightly tall, so
 * a small overshoot still fits; this keeps clear of the page square.
 */
export const FIT_TOLERANCE = 0.06;

/** Gutter between the two halves of a split row. */
export const SPLIT_GUTTER = 0.3;
export const SPLIT_W = (CONTENT_W - SPLIT_GUTTER) / 2;

/** Width of the KPI column when KPIs stand beside a table or chart. */
export const KPI_COL_W = 2.6;
/** A row of KPI cards across the slide. */
export const KPI_ROW_H = 1.15;
/** One card in a KPI column, and the gap between cards. */
export const KPI_STACK_CARD_H = 0.98;
export const KPI_STACK_GAP = 0.14;
/** A KPI card is never wider than this: a single 12in card put its one figure at the far left of a white bar. */
export const KPI_CARD_MAX_W = 3.6;

export const isFlexible = (b: Block) => b.type === "photos" || b.type === "chart";

// ---------------------------------------------------------------------------
// Text measurement
// ---------------------------------------------------------------------------

// Trebuchet MS, measured at the 10th percentile of real spans (so wide text
// still fits): body 10pt 13.6 characters to the inch, bold 10.5pt headers
// 11.6; lines 0.161in at 10pt. Set a little under, so estimates err tall -
// too tall moves a row to the next slide, too short draws one block over
// another.
const BODY_CPI_AT_10 = 13;
const HEADER_CPI_AT_10_5 = 11.2;
const cpi = (size: number, bold: boolean) => (bold ? (HEADER_CPI_AT_10_5 * 10.5) / size : (BODY_CPI_AT_10 * 10) / size);
const lineH = (size: number) => 0.0168 * size;

/**
 * Lines a piece of text wraps to at `perLine` characters. A hard line break
 * starts a new line wherever it falls - an Excel cell typed with Alt+Enter
 * keeps its breaks all the way to the slide - so each paragraph is wrapped on
 * its own and the lines summed. Counting a break as a space costed a
 * three-line cell as one line, and the table ran off the slide.
 */
function wrappedLines(text: string, perLine: number): number {
  return String(text ?? "")
    .split(/\r\n|\r|\n|\v/)
    .reduce((n, para) => n + paragraphLines(para, perLine), 0);
}

/** Lines one paragraph wraps to at `perLine` characters, breaking on spaces. */
function paragraphLines(text: string, perLine: number): number {
  const words = text.split(/\s+/).filter(Boolean);
  if (words.length === 0) return 1;
  let lines = 1;
  let used = 0;
  for (const word of words) {
    // A word longer than a whole line is broken across lines on its own.
    const len = word.length;
    if (len > perLine) {
      if (used > 0) lines++;
      lines += Math.ceil(len / perLine) - 1;
      used = len % perLine || perLine;
      continue;
    }
    const need = used === 0 ? len : used + 1 + len;
    if (need > perLine) {
      lines++;
      used = len;
    } else {
      used = need;
    }
  }
  return lines;
}

const longestToken = (text: unknown) =>
  Math.max(0, ...String(text ?? "").split(/\s+/).map((w) => w.length));

// Narrative at 12.5pt, bullets at 12pt, notes at 10pt and footnotes at 9pt,
// as generate-deck draws them.
const NARRATIVE_SIZE = 12.5;
const BULLET_SIZE = 12;
const NOTE_SIZE = 10;
export const FOOTNOTE_SIZE = 9;

function narrativeHeight(text: string, width: number): number {
  const lines = wrappedLines(text, Math.floor(width * cpi(NARRATIVE_SIZE, false)));
  return Math.max(0.5, 0.1 + lines * lineH(NARRATIVE_SIZE) * 1.1);
}

function bulletHeight(item: string, width: number): number {
  // The bullet glyph and its indent take about 0.4in of the line.
  const lines = wrappedLines(item, Math.floor((width - 0.4) * cpi(BULLET_SIZE, false)));
  return 0.06 + lines * lineH(BULLET_SIZE) * 1.1;
}

/** Height of a footnote drawn under a table or chart, or 0 when there is none. */
export function footnoteHeight(text: string | undefined, width: number): number {
  if (!text) return 0;
  return 0.08 + wrappedLines(text, Math.floor(width * cpi(FOOTNOTE_SIZE, false))) * lineH(FOOTNOTE_SIZE);
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

/** pptxgenjs's default cell margin is 0.1in on each side. */
const CELL_MARGIN_X = 0.2;
/** Above this many rows a table is drawn compact, as the house decks do. */
export const DENSE_ROWS = 12;

/** Font sizes a table is drawn at. Long tables are compact, never below 9pt; short ones with room to spare are drawn large. */
export function tableType(b: Table): { body: number; header: number; pad: number } {
  if (b.dense || b.rows.length > DENSE_ROWS) return { body: 9, header: 9.5, pad: 0.1 };
  if (b.roomy) return { body: 12, header: 12, pad: 0.24 };
  return { body: 10, header: 10.5, pad: 0.105 };
}

const ID_HEADER = /\b(sr|s\.no|no|number|id|code|year|phone|mobile|contact|tel|pin|account|invoice|bill)\b/i;
const NUMERIC = /^[-+]?(?:\d{1,3}(?:,\d{2,3})+|\d+)(?:\.\d+)?\s*%?$/;

/**
 * Columns whose cells are figures: at least seven in ten non-empty cells are
 * numbers. They are right-aligned, as figures are in the house tables, so a
 * column of them reads down by place value.
 */
export function numericColumns(b: Table): boolean[] {
  return b.headers.map((_, c) => {
    const cells = b.rows.map((r) => String(r[c] ?? "").trim()).filter(Boolean);
    if (cells.length === 0) return false;
    return cells.filter((v) => NUMERIC.test(v)).length / cells.length >= 0.7;
  });
}

/** Columns whose figures are identifiers (a serial number, a phone number, a year): never regrouped. */
export function identifierColumns(b: Table): boolean[] {
  return b.headers.map((h) => ID_HEADER.test(h));
}

/**
 * Width of each column, in inches, filling `width`.
 *
 * Each column gets at least what its longest unbroken word needs, so a figure
 * like 1003500 is never split across two lines, and the rest of the width is
 * shared by how much text each column carries. Equal columns gave a 3-letter
 * unit code 4in while a description beside it wrapped to five lines.
 */
export function tableColumnWidths(b: Table, width: number): number[] {
  const n = Math.max(1, b.headers.length);
  const { body, header } = tableType(b);
  const need = (chars: number, bold: boolean) => chars / cpi(bold ? header : body, bold) + CELL_MARGIN_X;
  const mins: number[] = [];
  const wants: number[] = [];
  for (let c = 0; c < n; c++) {
    const cells = b.rows.map((r) => String(r[c] ?? ""));
    const min = Math.max(0.55, need(longestToken(b.headers[c]), true), ...cells.map((v) => need(longestToken(v), false)));
    // The 75th-percentile cell length, so one long remark does not claim a
    // column the other rows leave empty.
    const lens = cells.map((v) => v.replace(/\s+/g, " ").trim().length).sort((x, y) => x - y);
    const p75 = lens.length ? lens[Math.min(lens.length - 1, Math.floor(lens.length * 0.75))] : 0;
    mins.push(min);
    wants.push(Math.max(min, Math.min(need(p75, false), 4.2), need(b.headers[c]?.length ?? 0, true) * 0.6));
  }
  const sum = (xs: number[]) => xs.reduce((a, x) => a + x, 0);
  if (sum(mins) >= width) return mins.map((m) => (m * width) / sum(mins));
  if (sum(wants) <= width) return wants.map((w) => (w * width) / sum(wants));
  // Every column its minimum, then the rest shared by how much more each wants.
  const spare = width - sum(mins);
  const extra = wants.map((w, i) => w - mins[i]);
  return mins.map((m, i) => m + (spare * extra[i]) / sum(extra));
}

function rowHeight(cells: string[], widths: number[], size: number, bold: boolean, pad = 0.105): number {
  const lines = Math.max(
    1,
    ...cells.map((c, i) =>
      wrappedLines(String(c ?? ""), Math.max(3, Math.floor(Math.max(0.2, (widths[i] ?? 1) - CELL_MARGIN_X) * cpi(size, bold))))
    )
  );
  return pad + lines * lineH(size);
}

/** Heights of a table's header, each row and each band, as drawn at `width`. */
export function tableGeometry(b: Table, width: number) {
  const widths = tableColumnWidths(b, width);
  const { body, header, pad } = tableType(b);
  const headerH = rowHeight(b.headers, widths, header, true, pad);
  const rowHs = b.rows.map((r) => rowHeight(r, widths, body, false, pad));
  // A band spans every column, in bold, and wraps like any cell.
  const bandHs = (b.groups ?? []).map((g) => rowHeight([g.label], [width], body, true, pad));
  return { widths, headerH, rowHs, bandHs, total: headerH + rowHs.reduce((a, h) => a + h, 0) + bandHs.reduce((a, h) => a + h, 0) };
}

/**
 * Whether every column's longest word fits on one line at `width`. A table
 * that would split a figure across lines at that width is drawn wider instead.
 */
export function tokensFit(b: Table, width: number): boolean {
  const { body, header } = tableType(b);
  const need = (chars: number, bold: boolean) => chars / cpi(bold ? header : body, bold) + CELL_MARGIN_X;
  const mins = b.headers.map((h, c) =>
    Math.max(need(longestToken(h), true), ...b.rows.map((r) => need(longestToken(r[c]), false)))
  );
  return mins.reduce((a, m) => a + m, 0) <= width;
}

// ---------------------------------------------------------------------------
// Photo grids
// ---------------------------------------------------------------------------

export type PhotoGeometry = { cols: number; rows: number; w: number; h: number; used: number };

const PHOTO_GAP = 0.18;
/** Below this a tile stops being evidence anyone can read. */
const MIN_TILE_W = 1.35;
/** A lone photo is not blown up past this; it starts to look like a poster. */
const MAX_TILE_W = 5.2;

/**
 * Tile shape for a grid, from the shape of its photos: portrait phone shots
 * get portrait tiles, near-square photos square ones, landscape photos
 * landscape ones. A fixed landscape tile kept 28% of a portrait photo's
 * height - ladder legs and torsos.
 */
export function photoTileRatio(aspect?: number): number {
  if (aspect === undefined) return 4 / 3;
  if (aspect < 0.9) return 3 / 4;
  if (aspect <= 1.2) return 1;
  return aspect >= 1.55 ? 1.6 : 4 / 3;
}

/**
 * Geometry of a photo grid in `available` vertical inches and `width`, or
 * null when no arrangement gives tiles of a readable size. Every column count
 * is tried and the one with the largest tiles wins: a fixed three columns
 * drew six near-square before/after photos as inch-wide thumbnails in a
 * slide with room for twice that. The renderer draws with this and the
 * planner measures with it.
 */
export function photoGeometry(count: number, available: number, aspect?: number, width = CONTENT_W): PhotoGeometry | null {
  const n = Math.min(count, 6);
  if (n <= 0) return null;
  const ratio = photoTileRatio(aspect);
  let best: PhotoGeometry | null = null;
  for (let cols = 1; cols <= n; cols++) {
    const rows = Math.ceil(n / cols);
    const byWidth = (width - PHOTO_GAP * (cols - 1)) / cols;
    const byHeight = ((available - PHOTO_GAP * (rows - 1)) / rows) * ratio;
    const w = Math.min(byWidth, byHeight, MAX_TILE_W);
    if (w < MIN_TILE_W) continue;
    const h = w / ratio;
    // Strictly larger tiles win, so a tie keeps the arrangement with fewer
    // columns found first only when it is genuinely as large.
    if (!best || w * h > best.w * best.h + 1e-6) {
      best = { cols, rows, w, h, used: rows * h + (rows - 1) * PHOTO_GAP };
    }
  }
  return best;
}

/**
 * The least height a grid can be drawn in - exactly the point below which
 * photoGeometry finds no readable arrangement. A flat 1.1in estimate once
 * left a six-photo grid too little room, and the photographs were silently
 * not drawn.
 */
export function photosMinHeight(count: number, aspect?: number, width = CONTENT_W): number {
  const n = Math.min(Math.max(count, 1), 6);
  const ratio = photoTileRatio(aspect);
  let least = Infinity;
  for (let cols = 1; cols <= n; cols++) {
    if (cols * MIN_TILE_W + (cols - 1) * PHOTO_GAP > width) break;
    const rows = Math.ceil(n / cols);
    least = Math.min(least, rows * (MIN_TILE_W / ratio) + (rows - 1) * PHOTO_GAP);
  }
  // A hair over the threshold, so floating-point rounding cannot land a grid
  // exactly on the boundary and have it rejected.
  return least + 0.01;
}

// ---------------------------------------------------------------------------
// Block and row heights
// ---------------------------------------------------------------------------

/** Height of a column of KPI cards. */
export function kpiStackHeight(count: number): number {
  return count * KPI_STACK_CARD_H + Math.max(0, count - 1) * KPI_STACK_GAP;
}

/** Height estimate so blocks stack without overlapping. */
export function blockHeight(b: Block, width = CONTENT_W): number {
  switch (b.type) {
    case "kpis":
      return width < CONTENT_W / 2 ? kpiStackHeight(b.items.length) : KPI_ROW_H;
    case "table":
      return tableGeometry(b, width).total + footnoteHeight(b.footnote, width);
    case "narrative":
      return narrativeHeight(b.text, width);
    case "bullets":
      return b.items.reduce((sum, item) => sum + bulletHeight(item, width), 0);
    case "photos":
    case "chart":
      // Photo grids and charts flex to whatever vertical space remains rather
      // than claiming a fixed height.
      return 0;
    case "note": {
      // Two lines fit the standard box; a longer caveat gets a taller one
      // rather than spilling out of its outline.
      const lines = wrappedLines(b.text, Math.floor((width - 0.32) * cpi(NOTE_SIZE, false)));
      return Math.max(0.44, 0.14 + lines * lineH(NOTE_SIZE));
    }
  }
}

/** The least vertical space a block can be drawn in and still be worth drawing. */
export function minBlockHeight(b: Block, width = CONTENT_W): number {
  if (b.type === "chart") return CHART_MIN_H + footnoteHeight(b.footnote, width);
  if (b.type === "photos") return photosMinHeight(b.imageIds.length, b.aspect, width);
  return blockHeight(b, width);
}

/**
 * One band across the slide: a single block spanning the full width, or two
 * blocks side by side with the left one `leftW` wide.
 */
export type Row =
  | { kind: "full"; block: Block }
  | { kind: "split"; left: Block; right: Block; leftW: number };

export const splitRightW = (row: { leftW: number }) => CONTENT_W - row.leftW - SPLIT_GUTTER;

const chartAndTable = (a: Block | undefined, b: Block | undefined) =>
  !!a && !!b && ((a.type === "chart" && b.type === "table") || (a.type === "table" && b.type === "chart"));

/**
 * Pairs blocks that belong side by side, the way the house decks lay them out:
 *
 * - a chart and the table of the same figures, chart on the left, each half
 *   the width. Stacked, a chart above a ten-row table was left an inch.
 * - KPIs beside the table or chart they summarise, in a narrow column on the
 *   left. Stacked above a table, a row of KPI cards was what spilled onto a
 *   slide of its own, a single figure on white.
 *
 * A pairing is skipped when the table would not survive the narrower width:
 * too tall to fit beside anything, or with a figure too long for its column.
 */
export function planRows(blocks: Block[]): Row[] {
  const rows: Row[] = [];
  for (let i = 0; i < blocks.length; i++) {
    const a = blocks[i];
    const b = blocks[i + 1];

    if (chartAndTable(a, b)) {
      const table = (a.type === "table" ? a : b) as Table;
      if (tokensFit(table, SPLIT_W) && blockHeight(table, SPLIT_W) <= CONTENT_H) {
        rows.push({ kind: "split", left: a.type === "chart" ? a : b!, right: a.type === "chart" ? b! : a, leftW: SPLIT_W });
        i++;
        continue;
      }
    }

    if (a.type === "kpis" && a.items.length <= 4 && b && (b.type === "table" || b.type === "chart") && !chartAndTable(b, blocks[i + 2])) {
      const rightW = CONTENT_W - KPI_COL_W - SPLIT_GUTTER;
      const fits =
        b.type === "chart" || (tokensFit(b, rightW) && blockHeight(b, rightW) <= CONTENT_H);
      if (fits && kpiStackHeight(a.items.length) <= CONTENT_H) {
        rows.push({ kind: "split", left: a, right: b, leftW: KPI_COL_W });
        i++;
        continue;
      }
    }

    rows.push({ kind: "full", block: a });
  }
  return rows;
}

/** Height a row needs at minimum. A split row is as tall as its taller half. */
export function rowMinHeight(row: Row): number {
  if (row.kind === "full") return minBlockHeight(row.block);
  return Math.max(minBlockHeight(row.left, row.leftW), minBlockHeight(row.right, splitRightW(row)));
}

/** True when the row grows into leftover space rather than claiming a fixed height. */
export function rowIsFlexible(row: Row): boolean {
  return row.kind === "full" ? isFlexible(row.block) : isFlexible(row.left) || isFlexible(row.right);
}

// ---------------------------------------------------------------------------
// Placing a slide's rows
// ---------------------------------------------------------------------------

export type PlacedRow = { row: Row; y: number; h: number; geo?: PhotoGeometry };

/**
 * Where each row of a slide goes: its top and height, and a photo grid's
 * geometry. The renderer draws from this and the preview positions from it,
 * so the two cannot place a block differently.
 *
 * Fixed rows take their measured height. Charts and split rows holding a
 * chart then share what is left: each first gets its minimum - the planner
 * has checked those fit - and the remainder is shared out, a chart on its own
 * capped so it does not balloon across the slide. Photo grids share what is
 * left after that, each first given its own minimum.
 *
 * Content starts under the title, as on every reference slide. It used to be
 * centred vertically, which floated a short table in mid-air with an inch and
 * a half of white above it.
 */
export function placeRows(rows: Row[]): PlacedRow[] {
  type Measured = { row: Row; h: number; geo?: PhotoGeometry };
  const measured: Measured[] = rows.map((row) => ({ row, h: rowIsFlexible(row) ? 0 : rowMinHeight(row) }));
  const gaps = BLOCK_SPACING * Math.max(0, rows.length - 1);
  let flexBudget = CONTENT_H - gaps - measured.reduce((sum, m) => sum + m.h, 0);

  const photoOf = (m: Measured) =>
    m.row.kind === "full" && m.row.block.type === "photos" ? m.row.block : null;
  const growable = measured.filter((m) => rowIsFlexible(m.row) && !photoOf(m));
  // Each grid's real minimum is held back, so a chart growing into the spare
  // space cannot leave a grid too little room to be drawn at all.
  const photoReserve = measured.reduce((sum, m) => {
    const p = photoOf(m);
    return p ? sum + photosMinHeight(p.imageIds.length, p.aspect) : sum;
  }, 0);

  for (const m of growable) {
    m.h = Math.max(CHART_MIN_H, rowMinHeight(m.row));
    flexBudget -= m.h;
  }
  let spare = Math.max(0, flexBudget - photoReserve);
  for (let i = 0; i < growable.length && spare > 0.01; i++) {
    const m = growable[i];
    const ceiling = m.row.kind === "split" ? CONTENT_H : CHART_MAX_H;
    const give = Math.min(Math.max(0, ceiling - m.h), spare / (growable.length - i));
    m.h += give;
    spare -= give;
    flexBudget -= give;
  }

  const photoRows = measured.filter(photoOf);
  let excess = flexBudget - photoReserve;
  let left = photoRows.length;
  for (const m of photoRows) {
    const p = photoOf(m)!;
    const min = photosMinHeight(p.imageIds.length, p.aspect);
    const share = min + Math.max(0, excess) / Math.max(1, left);
    left--;
    const geo = photoGeometry(Math.min(p.imageIds.length, 6), share, p.aspect);
    if (geo) {
      m.geo = geo;
      m.h = geo.used;
      excess -= Math.max(0, geo.used - min);
    }
  }

  const placed: PlacedRow[] = [];
  let y = CONTENT_TOP;
  for (const m of measured) {
    if (m.h === 0 && rowIsFlexible(m.row)) continue;
    placed.push({ row: m.row, y, h: m.h, geo: m.geo });
    y += m.h + BLOCK_SPACING;
  }
  return placed;
}

// ---------------------------------------------------------------------------
// Splitting sections across slides
// ---------------------------------------------------------------------------

/**
 * A short note that follows a table or chart becomes a footnote under it, in
 * small type, as the house decks caption a figure. As its own dashed box it
 * took 0.66in with its gap, and on a real deck that was the block that
 * spilled onto a slide of its own. A longer note keeps its box.
 */
function attachFootnotes(blocks: Block[]): Block[] {
  const out: Block[] = [];
  for (const b of blocks) {
    const prev = out[out.length - 1];
    if (
      b.type === "note" &&
      prev &&
      (prev.type === "table" || prev.type === "chart") &&
      !prev.footnote &&
      wrappedLines(b.text, Math.floor(SPLIT_W * cpi(FOOTNOTE_SIZE, false))) <= 2
    ) {
      out[out.length - 1] = { ...prev, footnote: b.text };
      continue;
    }
    out.push(b);
  }
  return out;
}

/**
 * Cuts a table taller than a slide into slide-sized tables of about equal
 * length, header repeated on each; category bands travel with the rows they
 * head. Filling each slide in turn left a six-row table followed by a stub of
 * two. A long table is drawn compact first, which is often enough to keep it
 * on one slide.
 */
function chunkTable(b: Table): Block[] {
  const t: Table = b.rows.length > DENSE_ROWS ? { ...b, dense: true } : b;
  if (blockHeight(t) <= CONTENT_H + FIT_TOLERANCE) return [t];

  const groups = t.groups ?? [];
  const pieceOf = (start: number, end: number, last: boolean): Table => ({
    type: "table",
    headers: t.headers,
    rows: t.rows.slice(start, end),
    dense: t.dense,
    // A band after the last row heads nothing, but renderTable draws it at
    // the end rather than losing it, so it travels with the last chunk.
    groups: groups
      .filter((g) => g.afterRow >= start && (g.afterRow < end || (last && g.afterRow >= end)))
      .map((g) => ({ label: g.label, afterRow: Math.min(g.afterRow, end) - start })),
    ...(last && t.footnote ? { footnote: t.footnote } : {}),
  });
  const fits = (start: number, end: number, last: boolean) =>
    blockHeight(pieceOf(start, end, last)) <= CONTENT_H + FIT_TOLERANCE;

  // The fewest chunks that fit, then rows dealt out evenly across them.
  for (let n = 2; n <= t.rows.length; n++) {
    // n chunks whose lengths differ by at most one row.
    const base = Math.floor(t.rows.length / n);
    const rem = t.rows.length % n;
    const bounds: [number, number][] = [];
    for (let k = 0, s = 0; k < n; k++) {
      const len = base + (k < rem ? 1 : 0);
      bounds.push([s, s + len]);
      s += len;
    }
    if (bounds.every(([s, e], k) => fits(s, e, k === bounds.length - 1))) {
      return bounds.map(([s, e], k) => pieceOf(s, e, k === bounds.length - 1));
    }
  }
  // One row per slide still too tall: each row on its own; nothing is lost.
  return t.rows.map((_, r) => pieceOf(r, r + 1, r === t.rows.length - 1));
}

/**
 * Greedily packs `parts` into runs whose height, by `height`, fits a slide.
 * A part too tall on its own still gets a run to itself; nothing is dropped.
 */
function pack<T>(parts: T[], height: (run: T[]) => number): T[][] {
  const runs: T[][] = [];
  let run: T[] = [];
  for (const p of parts) {
    if (run.length > 0 && height([...run, p]) > CONTENT_H) {
      runs.push(run);
      run = [];
    }
    run.push(p);
  }
  if (run.length > 0) runs.push(run);
  return runs;
}

/**
 * Cuts a bullet list or a narrative taller than a slide into slide-sized
 * blocks, as chunkTable does for tables. A narrative breaks between sentences,
 * and inside a sentence only if one sentence alone is too tall.
 */
function chunkText(b: Extract<Block, { type: "bullets" | "narrative" }>): Block[] {
  if (blockHeight(b) <= CONTENT_H + FIT_TOLERANCE) return [b];
  if (b.type === "bullets") {
    return pack(b.items, (items) => blockHeight({ type: "bullets", items })).map((items) => ({ type: "bullets", items }));
  }
  const fits = (text: string) => narrativeHeight(text, CONTENT_W) <= CONTENT_H;
  const pieces = b.text
    .split(/(?<=[.!?])\s+/)
    .flatMap((sentence) => (fits(sentence) ? [sentence] : sentence.split(/\s+/)));
  return pack(pieces, (run) => narrativeHeight(run.join(" "), CONTENT_W)).map((run) => ({
    type: "narrative",
    text: run.join(" "),
  }));
}

const isOrphanable = (row: Row) =>
  row.kind === "full" && (row.block.type === "kpis" || row.block.type === "note");

/**
 * Takes the rows that fit on one slide and leaves the rest for a continuation
 * slide. Spilling is the one option that loses nothing.
 *
 * KPI cards or a note left as the last thing on a slide, with what they
 * describe on the next, go over with it: a slide ending in a stranded figure
 * reads as broken.
 */
export function splitRows(rows: Row[]): { take: Row[]; rest: Row[] } {
  const take: Row[] = [];
  let used = 0;

  for (const row of rows) {
    const h = rowMinHeight(row);
    const gap = take.length > 0 ? BLOCK_SPACING : 0;
    if (take.length > 0 && used + gap + h > CONTENT_H + FIT_TOLERANCE) break;
    take.push(row);
    used += gap + h;
  }
  // Always take one. chunkTable and chunkText have already cut any block that
  // could not fit a slide on its own, so this only guards against a loop.
  if (take.length === 0) take.push(rows[0]);
  while (take.length > 1 && take.length < rows.length && isOrphanable(take[take.length - 1])) take.pop();
  return { take, rest: rows.slice(take.length) };
}

/**
 * A short table with the slide to itself is drawn large. At body size a
 * three-row table filled a sixth of the slide under the title and left the
 * rest white; the reference decks set a short table at reading size. Only
 * when nothing on the slide grows into the space anyway - a chart or photo
 * grid takes it - and only when the larger table still fits.
 */
function roomyTables(rows: Row[]): Row[] {
  if (rows.some(rowIsFlexible)) return rows;
  const used = rows.reduce((sum, r) => sum + rowMinHeight(r), 0) + BLOCK_SPACING * Math.max(0, rows.length - 1);
  let spare = CONTENT_H - used;
  return rows.map((row) => {
    if (row.kind !== "full" || row.block.type !== "table") return row;
    const t = row.block;
    if (t.dense || t.rows.length > 6 || spare < 1) return row;
    const roomy: Table = { ...t, roomy: true };
    const grow = blockHeight(roomy) - blockHeight(t);
    if (grow > spare) return row;
    spare -= grow;
    return { kind: "full", block: roomy };
  });
}

/** Every slide a section becomes, in order. */
export function sectionSlides<T extends { label: string; blocks: Block[] }>(
  section: T
): { label: string; rows: Row[] }[] {
  const blocks = attachFootnotes(section.blocks).flatMap((b) =>
    b.type === "table" ? chunkTable(b) : b.type === "bullets" || b.type === "narrative" ? chunkText(b) : [b]
  );
  const out: { label: string; rows: Row[] }[] = [];
  let remaining = planRows(blocks);
  while (remaining.length > 0) {
    const { take, rest } = splitRows(remaining);
    out.push({
      label: out.length === 0 ? section.label : `${section.label} (continued)`,
      rows: roomyTables(take),
    });
    remaining = rest;
  }
  return out;
}

// ---------------------------------------------------------------------------
// The deck as a whole
// ---------------------------------------------------------------------------

export type OutlineSlide =
  | { kind: "cover" }
  | { kind: "glance" }
  | { kind: "contents"; labels: string[] }
  | { kind: "section"; label: string; rows: Row[] }
  | { kind: "closing" };

/**
 * Every slide the deck contains, in order.
 *
 * The preview draws from this and the generate route counts it. When the
 * contents and closing slides were added to the renderer, the preview and the
 * reported slide count both went on describing the old deck - the submitter
 * approved a deck two slides shorter than the one filed. Keep this in step
 * with generateDeck.
 */
export function deckOutline<T extends { label: string; blocks: Block[] }>(sections: T[]): OutlineSlide[] {
  return [
    { kind: "cover" },
    { kind: "glance" },
    { kind: "contents", labels: sections.map((s) => s.label) },
    ...sections.flatMap((s) => sectionSlides(s).map((page) => ({ kind: "section" as const, ...page }))),
    { kind: "closing" },
  ];
}
