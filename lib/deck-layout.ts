import type { Block } from "./curate";

/**
 * How much room each block needs, how blocks pair up across a slide, and
 * where a section has to break onto a second slide.
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
// photoGeometry gives up under this, so it is the grid's real minimum.
export const PHOTOS_MIN_H = 1.1;
export const BLOCK_SPACING = 0.22;

/** Gutter between the two halves of a split row. */
export const SPLIT_GUTTER = 0.3;
export const SPLIT_W = (CONTENT_W - SPLIT_GUTTER) / 2;

export const isFlexible = (b: Block) => b.type === "photos" || b.type === "chart";

/**
 * Height estimate so blocks stack without overlapping. The table figures were
 * corrected against real rendered output — PowerPoint lays rows out noticeably
 * taller than a naive guess, and underestimating meant the following block
 * drew on top of the table instead of below it.
 *
 * `width` matters for tables: the same rows wrap far more in half a slide than
 * across the full one.
 */
export function blockHeight(b: Block, width = CONTENT_W): number {
  switch (b.type) {
    case "kpis":
      return 1.15;
    case "table": {
      // ~0.28 per single-line row, measured against real rendered output (an
      // 11-row grouped table occupied ~3.3in). Long cells wrap, though, and
      // treating every row as one line put a three-row fit-out table — whose
      // every cell ran to three lines — straight off the bottom of the slide.
      // So rows are costed by how many lines their widest cell needs.
      const cols = Math.max(1, b.headers.length);
      // 10pt Calibri runs around 13 characters to the inch.
      const perLine = Math.max(6, Math.floor((width / cols) * 13));
      const rowHeight = (cells: string[]) => {
        const lines = Math.max(1, ...cells.map((c) => Math.ceil(String(c ?? "").length / perLine)));
        return 0.06 + lines * 0.22;
      };
      const bands = (b.groups?.length ?? 0) * 0.28;
      return b.rows.reduce((sum, r) => sum + rowHeight(r), rowHeight(b.headers)) + bands;
    }
    case "narrative":
      return Math.max(0.5, Math.ceil(b.text.length / 135) * 0.3);
    case "bullets":
      return b.items.length * 0.32;
    case "photos":
    case "chart":
      // Photo grids and charts flex to whatever vertical space remains rather
      // than claiming a fixed height. A fixed estimate meant a six-photo grid
      // could be judged "too tall to fit" and silently dropped — losing the
      // strongest content on the slide.
      return 0;
    case "note":
      return 0.44;
  }
}

/** The least vertical space a block can be drawn in and still be worth drawing. */
export function minBlockHeight(b: Block, width = CONTENT_W): number {
  if (b.type === "chart") return CHART_MIN_H;
  if (b.type === "photos") return PHOTOS_MIN_H;
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
 * height — bars reduced to slivers with half the category labels missing. The
 * house decks put the table on one side and the chart on the other, and that
 * is also the only arrangement where both are legible. Pairing is skipped when
 * the table is too tall to fit beside anything, so it falls back to stacking
 * and, if need be, a continuation slide.
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
      const table = a.type === "table" ? a : b;
      const fits = Math.max(CHART_MIN_H, blockHeight(table, SPLIT_W)) <= CONTENT_H;
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

/**
 * Takes the rows that fit on one slide and leaves the rest for a continuation
 * slide.
 *
 * Sections really do overflow: Claude is asked to put a chart with the table of
 * the same figures, and a ten-row table plus KPIs plus a note already fills the
 * content area on its own. Before this existed the overflow was dropped in
 * silence — a chart curation had chosen simply never appeared, with nothing in
 * the file or the logs to say so. Two of five charts were lost that way on a
 * real submission. Spilling is the one option that loses nothing.
 */
export function splitRows(rows: Row[]): { take: Row[]; rest: Row[] } {
  const take: Row[] = [];
  let used = 0;

  for (const row of rows) {
    const h = rowMinHeight(row);
    const gap = take.length > 0 ? BLOCK_SPACING : 0;
    if (take.length > 0 && used + gap + h > CONTENT_H) break;
    take.push(row);
    used += gap + h;
  }

  // Always take one, however tall: a row bigger than a whole slide would
  // otherwise spill forever.
  if (take.length === 0) take.push(rows[0]);
  return { take, rest: rows.slice(take.length) };
}

/** Every slide a section becomes, in order. */
export function sectionSlides<T extends { label: string; blocks: Block[] }>(
  section: T
): { label: string; rows: Row[] }[] {
  const out: { label: string; rows: Row[] }[] = [];
  let remaining = planRows(section.blocks);
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
