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
 * Everything here is pure arithmetic in inches, safe to import from a client
 * component.
 */

// Geometry of the content area, shared with the renderer.
export const CONTENT_W = 12.13;
export const CONTENT_TOP = 1.78;
export const CONTENT_BOTTOM = 6.85;
export const CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP;

// A chart below about an inch and a half has colliding axis labels; above
// three and a half it swamps whatever else is on the slide.
export const CHART_MIN_H = 1.5;
export const CHART_MAX_H = 3.4;
export const BLOCK_SPACING = 0.22;

/**
 * How far a page's estimated height may run past the content area before the
 * planner moves a row to the next slide.
 *
 * The estimates are deliberately a little generous, and on a real PPM section
 * an overshoot of 0.04in pushed a lone KPI card onto a slide of its own. 0.06in
 * stays clear of the page-number box, which starts at y=6.92.
 */
export const FIT_TOLERANCE = 0.06;

/** Gutter between the two halves of a split row. */
export const SPLIT_GUTTER = 0.3;
export const SPLIT_W = (CONTENT_W - SPLIT_GUTTER) / 2;

export const isFlexible = (b: Block) => b.type === "photos" || b.type === "chart";

// ---------------------------------------------------------------------------
// Photo grids
// ---------------------------------------------------------------------------

export type PhotoGeometry = { cols: number; rows: number; w: number; h: number; used: number };

const PHOTO_GAP = 0.18;
const PHOTO_MIN_SCALE = 0.35;

/**
 * Geometry of a photo grid in `available` vertical inches, or null when it
 * cannot be drawn at a useful size. The renderer draws with this and the
 * planner measures with it; they used to keep separate figures.
 */
export function photoGeometry(count: number, available: number): PhotoGeometry | null {
  if (count <= 0) return null;
  const cols = count <= 3 ? count : 3;
  const rows = Math.ceil(count / cols);
  const ratio = rows > 1 ? 1.6 : 4 / 3;

  let w = (CONTENT_W - PHOTO_GAP * (cols - 1)) / cols;
  let h = w / ratio;

  const needed = rows * h + (rows - 1) * PHOTO_GAP;
  if (needed > available) {
    const scale = (available - (rows - 1) * PHOTO_GAP) / (rows * h);
    if (scale < PHOTO_MIN_SCALE) return null;
    w *= scale;
    h *= scale;
  }
  return { cols, rows, w, h, used: rows * h + (rows - 1) * PHOTO_GAP };
}

/**
 * The least height a grid of `count` photos can be drawn in - exactly the
 * point below which photoGeometry gives up.
 *
 * The planner used to assume a flat 1.1in for any grid. A six-photo grid
 * actually needs about 1.9in, so a slide planned with a chart or a split row
 * left it 1.1-1.9in, photoGeometry returned null, and the photographs - the
 * strongest evidence on the slide - were silently not drawn.
 */
export function photosMinHeight(count: number): number {
  const n = Math.min(Math.max(count, 1), 6);
  const cols = n <= 3 ? n : 3;
  const rows = Math.ceil(n / cols);
  const ratio = rows > 1 ? 1.6 : 4 / 3;
  const h = (CONTENT_W - PHOTO_GAP * (cols - 1)) / cols / ratio;
  // A hair over the threshold, so floating-point rounding cannot land a grid
  // exactly on the boundary and have it rejected.
  return rows * h * PHOTO_MIN_SCALE + (rows - 1) * PHOTO_GAP + 0.01;
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

/** pptxgenjs's default cell margin is 0.1in on each side. */
const CELL_MARGIN_X = 0.2;
// Calibri at 10pt runs about 12 characters to the inch; bold 10.5pt headers
// about 11. Erring narrow makes the estimate tall, which is the safe side:
// too tall moves a row to the next slide, too short draws one block over
// another.
const BODY_CHARS_PER_IN = 12;
const HEADER_CHARS_PER_IN = 11;

/** Lines a cell's text wraps to at `perLine` characters, breaking on spaces. */
function wrappedLines(text: string, perLine: number): number {
  const words = String(text ?? "").split(/\s+/).filter(Boolean);
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

function tableRowHeight(cells: string[], width: number, cols: number, header: boolean): number {
  const usable = Math.max(0.3, width / cols - CELL_MARGIN_X);
  const perLine = Math.max(4, Math.floor(usable * (header ? HEADER_CHARS_PER_IN : BODY_CHARS_PER_IN)));
  const lines = Math.max(1, ...cells.map((c) => wrappedLines(String(c ?? ""), perLine)));
  return 0.06 + lines * 0.22;
}

/**
 * Whether every cell's longest unbroken token fits on one line at `width`.
 *
 * A figure like "1003500" has no space to break at, so in a column too narrow
 * for it PowerPoint splits the number across two lines - "10035 / 00". A table
 * that does this at half width must not be paired; it goes full width instead.
 */
export function tokensFit(b: Extract<Block, { type: "table" }>, width: number): boolean {
  const cols = Math.max(1, b.headers.length);
  const usable = Math.max(0.3, width / cols - CELL_MARGIN_X);
  const bodyCap = Math.floor(usable * BODY_CHARS_PER_IN);
  const headCap = Math.floor(usable * HEADER_CHARS_PER_IN);
  const longest = (cells: string[]) =>
    Math.max(0, ...cells.flatMap((c) => String(c ?? "").split(/\s+/).map((w) => w.length)));
  return longest(b.headers) <= headCap && b.rows.every((r) => longest(r) <= bodyCap);
}

// ---------------------------------------------------------------------------
// Block and row heights
// ---------------------------------------------------------------------------

/**
 * Height estimate so blocks stack without overlapping.
 *
 * Tables are costed row by row against the space inside each cell - the
 * column width less its margins - wrapping by words. Costing by raw column
 * width and a flat characters-per-line figure underestimated narrow columns,
 * and the block below was drawn over the table's last rows.
 */
export function blockHeight(b: Block, width = CONTENT_W): number {
  switch (b.type) {
    case "kpis":
      return 1.15;
    case "table": {
      const cols = Math.max(1, b.headers.length);
      const bands = (b.groups?.length ?? 0) * 0.28;
      return (
        b.rows.reduce((sum, r) => sum + tableRowHeight(r, width, cols, false), tableRowHeight(b.headers, width, cols, true)) +
        bands
      );
    }
    case "narrative":
      return Math.max(0.5, Math.ceil(b.text.length / 135) * 0.3);
    case "bullets":
      return b.items.length * 0.32;
    case "photos":
    case "chart":
      // Photo grids and charts flex to whatever vertical space remains rather
      // than claiming a fixed height. A fixed estimate meant a six-photo grid
      // could be judged "too tall to fit" and silently dropped.
      return 0;
    case "note":
      return 0.44;
  }
}

/** The least vertical space a block can be drawn in and still be worth drawing. */
export function minBlockHeight(b: Block, width = CONTENT_W): number {
  if (b.type === "chart") return CHART_MIN_H;
  if (b.type === "photos") return photosMinHeight(b.imageIds.length);
  return blockHeight(b, width);
}

/**
 * One band across the slide: either a single block spanning the full width, or
 * a chart and its table side by side.
 */
export type Row =
  | { kind: "full"; block: Block }
  | { kind: "split"; left: Block; right: Block };

/**
 * Pairs a chart with the table of the same figures so they sit side by side.
 *
 * Stacked full-width, a chart above a ten-row table was left about an inch of
 * height. The house decks put the table on one side and the chart on the
 * other. Pairing is skipped when the table would not survive half width -
 * too tall to fit beside anything, or with figures too long for their column
 * - and it stacks full width instead.
 */
export function planRows(blocks: Block[]): Row[] {
  const rows: Row[] = [];
  for (let i = 0; i < blocks.length; i++) {
    const a = blocks[i];
    const b = blocks[i + 1];
    const pairable =
      b !== undefined &&
      ((a.type === "chart" && b.type === "table") || (a.type === "table" && b.type === "chart"));

    if (pairable) {
      const table = (a.type === "table" ? a : b) as Extract<Block, { type: "table" }>;
      const fits =
        tokensFit(table, SPLIT_W) && Math.max(CHART_MIN_H, blockHeight(table, SPLIT_W)) <= CONTENT_H;
      if (fits) {
        // Chart on the left, table on the right, whichever order they arrived in.
        const left = a.type === "chart" ? a : b;
        const right = a.type === "chart" ? b : a;
        rows.push({ kind: "split", left, right });
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
  return Math.max(minBlockHeight(row.left, SPLIT_W), minBlockHeight(row.right, SPLIT_W));
}

/** True when the row grows into leftover space rather than claiming a fixed height. */
export function rowIsFlexible(row: Row): boolean {
  return row.kind === "full" ? isFlexible(row.block) : true;
}

// ---------------------------------------------------------------------------
// Splitting sections across slides
// ---------------------------------------------------------------------------

/**
 * Cuts a table taller than a whole slide into slide-sized tables, header
 * repeated on each. Category band rows travel with the rows they head.
 *
 * Without this the planner had to take such a table whole - it could not be
 * moved anywhere it fitted - and it was drawn on past the bottom edge with its
 * last rows lost.
 */
function chunkTable(b: Extract<Block, { type: "table" }>): Block[] {
  if (blockHeight(b) <= CONTENT_H + FIT_TOLERANCE) return [b];

  const cols = Math.max(1, b.headers.length);
  const headerH = tableRowHeight(b.headers, CONTENT_W, cols, true);
  const groups = b.groups ?? [];
  const chunks: Block[] = [];

  let start = 0;
  while (start < b.rows.length) {
    let used = headerH;
    let end = start;
    while (end < b.rows.length) {
      const bands = groups.filter((g) => g.afterRow === end).length * 0.28;
      const h = tableRowHeight(b.rows[end], CONTENT_W, cols, false) + bands;
      if (end > start && used + h > CONTENT_H) break;
      used += h;
      end++;
    }
    chunks.push({
      type: "table",
      headers: b.headers,
      rows: b.rows.slice(start, end),
      groups: groups
        .filter((g) => g.afterRow >= start && g.afterRow < end)
        .map((g) => ({ label: g.label, afterRow: g.afterRow - start })),
    });
    start = end;
  }
  return chunks;
}

/**
 * Takes the rows that fit on one slide and leaves the rest for a continuation
 * slide.
 *
 * Sections really do overflow: Claude is asked to put a chart with the table of
 * the same figures, and a ten-row table plus KPIs plus a note already fills the
 * content area on its own. Before this existed the overflow was dropped in
 * silence. Spilling is the one option that loses nothing.
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

  // Always take one. chunkTable has already cut any table that could not fit
  // a slide on its own, so this only guards against a loop.
  if (take.length === 0) take.push(rows[0]);
  return { take, rest: rows.slice(take.length) };
}

/** Every slide a section becomes, in order. */
export function sectionSlides<T extends { label: string; blocks: Block[] }>(
  section: T
): { label: string; rows: Row[] }[] {
  const blocks = section.blocks.flatMap((b) => (b.type === "table" ? chunkTable(b) : [b]));
  const out: { label: string; rows: Row[] }[] = [];
  let remaining = planRows(blocks);
  while (remaining.length > 0) {
    const { take, rest } = splitRows(remaining);
    out.push({
      label: out.length === 0 ? section.label : `${section.label} (continued)`,
      rows: take,
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
