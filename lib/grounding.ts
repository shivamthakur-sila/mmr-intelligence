import type { Block, CuratedDeck } from "./curate";

/**
 * Checks every figure the curated deck states against the submitted source,
 * and removes the ones it cannot account for.
 *
 * This is the mechanical half of "never invent a number". The prompts ask for
 * it; nothing used to check it. A figure is accepted when it:
 *
 *   1. appears in the source text or in an answer the submitter typed, or
 *   2. is the sum of figures the same section shows - a chart series, a table
 *      column or a table row - or
 *   3. is the count of rows in a table, or of categories in a chart, that the
 *      same section shows.
 *
 * (2) and (3) are allowed because a total sitting beside its parts can be
 * checked by the reader on the slide itself. Anything else - a percentage the
 * source does not state, an average, a rounded or estimated figure - is not.
 *
 * What happens to a figure that fails depends on what it is part of:
 *   - a chart value: the chart is dropped. A chart with one invented bar is a
 *     wrong chart, and there is no honest way to draw it with a gap.
 *   - a KPI: that card is dropped.
 *   - a bullet or summary point: that line is dropped. If the headline fails,
 *     the first point that passes takes its place.
 *   - a narrative or note: the block is dropped.
 *   - a table cell: reported, not changed. Tables are copied from the source
 *     and a false alarm there would cut real data; the report is what lets a
 *     human look.
 *
 * Limits, stated so nobody relies on more than this gives: it matches numbers,
 * not their meaning. A real figure placed against the wrong label, month or
 * site passes, and small numbers (0, 1, 2, 100) appear in almost any source.
 * It catches the failure that matters most - a specific figure that is in no
 * submitted file - and it catches it every time.
 */

export type GroundingIssue = {
  section: string;
  where: string;
  value: string;
  action: "dropped" | "reported";
};

// A number with thousands grouping, Western (1,448,000) or Indian (21,32,395).
// Anchored on both sides so that a list like 45748,27240 - two numbers in a
// JSON row - is never read as one grouped number.
const GROUPED = /(?<![\d.])\d{1,3}(?:,\d{2,3})+(?:\.\d+)?(?!\d)/g;
// Any run of digits, which is how list items and ungrouped figures appear.
const PLAIN = /\d+(?:\.\d+)?/g;
// What a figure in the deck is read as: grouped if it is grouped, else plain.
const FIGURE = /(?<![\d.])(?:\d{1,3}(?:,\d{2,3})+|\d+)(?:\.\d+)?(?!\d)/g;

/**
 * Canonical form of a number. toPrecision(12) folds spreadsheet float noise
 * (528.0000000000007) onto the figure a person would write (528) without
 * rounding any real decimal a source actually states.
 */
function canon(raw: string): string | null {
  const n = Number(raw.replace(/,/g, ""));
  if (!Number.isFinite(n)) return null;
  return String(Number(n.toPrecision(12)));
}

/** Every figure in a piece of deck text, as the reader would read it. */
export function figuresIn(text: unknown): string[] {
  const out: string[] = [];
  for (const m of String(text ?? "").matchAll(FIGURE)) {
    const c = canon(m[0]);
    if (c !== null) out.push(c);
  }
  return out;
}

/**
 * Every number in the source. Generous on purpose: a grouped number is added
 * both whole and as its parts, because "12,345" in a source may be one figure
 * or two list items, and adding both only ever prevents a false alarm.
 */
export function sourcePool(source: string): Set<string> {
  const pool = new Set<string>();
  for (const re of [GROUPED, PLAIN]) {
    for (const m of source.matchAll(re)) {
      const c = canon(m[0]);
      if (c !== null) pool.add(c);
    }
  }
  return pool;
}

function numeric(cell: unknown): number | null {
  const s = String(cell ?? "").trim();
  if (!/^-?[\d,]+(\.\d+)?$/.test(s)) return null;
  const n = Number(s.replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
}

/** Totals and counts a section's own content lets a reader check. */
function derivable(blocks: Block[]): Set<string> {
  const out = new Set<string>();
  const add = (n: number) => {
    const c = canon(String(n));
    if (c !== null) out.add(c);
  };

  for (const b of blocks) {
    if (b.type === "chart") {
      add(b.categories.length);
      for (const s of b.series) add(s.values.reduce((a, v) => a + v, 0));
    }
    if (b.type === "table") {
      add(b.rows.length);
      const cols = b.headers.length;
      for (let c = 0; c < cols; c++) {
        const col = b.rows.map((r) => numeric(r[c]));
        const present = col.filter((v): v is number => v !== null);
        if (present.length >= 2) add(present.reduce((a, v) => a + v, 0));
      }
      for (const r of b.rows) {
        const nums = r.map(numeric).filter((v): v is number => v !== null);
        if (nums.length >= 2) add(nums.reduce((a, v) => a + v, 0));
        // A row's last figure is often its own total; a total of the others
        // must be checkable too.
        if (nums.length >= 3) add(nums.slice(0, -1).reduce((a, v) => a + v, 0));
      }
    }
  }
  return out;
}

export function groundDeck(
  deck: CuratedDeck,
  source: string
): { deck: CuratedDeck; issues: GroundingIssue[] } {
  const pool = sourcePool(source);
  const issues: GroundingIssue[] = [];
  const allDerived = new Set<string>();

  const sections = deck.sections.map((section) => {
    const derived = derivable(section.blocks);
    derived.forEach((d) => allDerived.add(d));
    const known = (f: string) => pool.has(f) || derived.has(f);
    const unknownIn = (text: unknown) => figuresIn(text).filter((f) => !known(f));
    const drop = (where: string, value: unknown) =>
      issues.push({ section: section.key, where, value: String(value), action: "dropped" });

    const blocks: Block[] = [];
    for (const b of section.blocks) {
      switch (b.type) {
        case "chart": {
          const bad = b.series.flatMap((s) => s.values.map(String)).filter((v) => unknownIn(v).length > 0);
          if (bad.length > 0) {
            drop("chart", bad.join(", "));
            continue;
          }
          blocks.push(b);
          break;
        }
        case "kpis": {
          const items = b.items.filter((k) => {
            if (unknownIn(k.value).length === 0) return true;
            drop(`kpi "${k.label}"`, k.value);
            return false;
          });
          if (items.length > 0) blocks.push({ ...b, items });
          break;
        }
        case "bullets": {
          const items = b.items.filter((t) => {
            if (unknownIn(t).length === 0) return true;
            drop("bullet", t);
            return false;
          });
          if (items.length > 0) blocks.push({ ...b, items });
          break;
        }
        case "narrative":
        case "note": {
          if (unknownIn(b.text).length > 0) {
            drop(b.type, b.text);
            continue;
          }
          blocks.push(b);
          break;
        }
        case "table": {
          for (const row of b.rows) {
            for (const cell of row) {
              if (unknownIn(cell).length > 0) {
                issues.push({ section: section.key, where: "table cell", value: String(cell), action: "reported" });
              }
            }
          }
          blocks.push(b);
          break;
        }
        default:
          blocks.push(b);
      }
    }
    return { ...section, blocks };
  });

  // The summary speaks for the whole deck, so a total from any section counts.
  const knownAnywhere = (f: string) => pool.has(f) || allDerived.has(f);
  const summaryOk = (t: string) => figuresIn(t).every(knownAnywhere);
  const points = deck.summary.points.filter((p) => {
    if (summaryOk(p)) return true;
    issues.push({ section: "summary", where: "point", value: p, action: "dropped" });
    return false;
  });
  let headline = deck.summary.headline;
  if (!summaryOk(headline)) {
    issues.push({ section: "summary", where: "headline", value: headline, action: "dropped" });
    headline = points.shift() ?? "";
  }

  return {
    deck: {
      ...deck,
      summary: { headline, points },
      sections: sections.filter((s) => s.blocks.length > 0),
    },
    issues,
  };
}
