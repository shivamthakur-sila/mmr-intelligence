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
 * checked by the reader on the slide itself. So they are computed only from
 * parts that passed (1) and stay on the slide: a total of a chart that was
 * dropped, or of a table cell that is in no source, would be a total of
 * nothing the reader can see. Anything else - a percentage the source does not
 * state, an average, a rounded or estimated figure - is not accepted.
 *
 * Every field that reaches a slide is checked, not only values and body
 * text: a chart title or a KPI label states figures to the client as plainly
 * as a bar does. What happens to a figure that fails depends on where it is:
 *   - a chart value, category, series name or unit: the chart is dropped. A
 *     chart with one invented bar is a wrong chart, and there is no honest way
 *     to draw it with a gap. A failing chart title is removed alone.
 *   - a KPI value or label: that card is dropped.
 *   - a bullet or summary point: that line is dropped. If the headline fails,
 *     the first point that passes takes its place.
 *   - a narrative or note: the block is dropped.
 *   - a table cell: its row is dropped. A table header: the table is dropped.
 *     A group band label: the band is dropped and its rows stay.
 *   - a section label: the parts of it that state the figure are removed.
 * Every removal is returned, and the submitter sees each one before filing.
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
  action: "dropped";
};

// A number with thousands grouping, Western (1,448,000) or Indian (21,32,395).
// Anchored on both sides so that a list like 45748,27240 - two numbers in a
// JSON row - is never read as one grouped number.
const GROUPED = /(?<![\d.])\d{1,3}(?:,\d{2,3})+(?:\.\d+)?(?!\d)/g;
// Any run of digits, which is how list items and ungrouped figures appear.
const PLAIN = /\d+(?:\.\d+)?/g;
// What a figure in the deck is read as: grouped if it is grouped, else plain.
const FIGURE = /(?<![\d.])(?:\d{1,3}(?:,\d{2,3})+|\d+)(?:\.\d+)?(?!\d)/g;

// Numbers the extractor writes around the content - slide numbers in its
// headings and the names of embedded files - are not figures anybody
// submitted, so they must not vouch for one. Chart titles in those headings
// are the submitter's own text and stay.
function withoutScaffolding(source: string): string {
  return source
    .replace(/^--- Slide \d+ ---$/gm, "")
    .replace(/^=== Embedded workbook .*$/gm, (line) =>
      line.replace(/\bslide \d+|\((?:Microsoft_Excel_Worksheet|oleObject)\d*\.(?:xlsx|bin)\)/g, "")
    );
}

/**
 * Canonical form of a number. toPrecision(12) folds spreadsheet float noise
 * (528.0000000000007) onto the figure a person would write (528) without
 * rounding any real decimal a source actually states.
 */
function canon(raw: string | number): string | null {
  const n = typeof raw === "number" ? raw : Number(raw.replace(/,/g, ""));
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
 *
 * A fraction is also added as a percentage. A spreadsheet cell formatted as
 * a percentage stores 0.985 and shows 98.5%, the deck rightly says 98.5%, and
 * without this a real PPM section was deleted outright.
 */
export function sourcePool(source: string): Set<string> {
  const text = withoutScaffolding(source);
  const pool = new Set<string>();
  for (const re of [GROUPED, PLAIN]) {
    for (const m of text.matchAll(re)) {
      const c = canon(m[0]);
      if (c === null) continue;
      pool.add(c);
      const n = Number(c);
      if (!Number.isInteger(n) && n > 0 && n <= 1) {
        const pct = canon(n * 100);
        if (pct !== null) pool.add(pct);
      }
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

/**
 * Totals and counts that a section's own content lets a reader check. Call it
 * only on blocks that have already passed: the parts of a total have to be on
 * the slide for the total to be checkable.
 */
function derivable(blocks: Block[]): Set<string> {
  const out = new Set<string>();
  const add = (n: number) => {
    const c = canon(n);
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

/** The figures in `text` that `known` cannot account for. */
function unknownIn(text: unknown, known: (f: string) => boolean): string[] {
  return figuresIn(text).filter((f) => !known(f));
}

/**
 * A table with every row that states an unaccounted figure removed. A cell
 * may be a total of the table's own grounded cells - a TOTAL row the model
 * added - so those totals are computed from the cells the source vouches for
 * and nothing else.
 */
function groundTable(
  b: Extract<Block, { type: "table" }>,
  inPool: (f: string) => boolean,
  drop: (where: string, value: unknown) => void
): Extract<Block, { type: "table" }> | null {
  const header = b.headers.find((h) => unknownIn(h, inPool).length > 0);
  if (header !== undefined) {
    drop("table header", header);
    return null;
  }

  const sourced = b.rows.filter((r) => r.every((cell) => unknownIn(cell, inPool).length === 0));
  const totals = derivable([{ ...b, rows: sourced }]);
  const known = (f: string) => inPool(f) || totals.has(f);

  const keep: boolean[] = b.rows.map((r) => {
    const bad = r.find((cell) => unknownIn(cell, known).length > 0);
    if (bad === undefined) return true;
    drop("table row", r.join(" | "));
    return false;
  });
  if (!keep.includes(true)) return null;

  // Band positions count rows, so each is moved up by the rows dropped
  // before it. A band whose label states an unknown figure goes; its rows
  // stay, as they passed on their own.
  const removedBefore = (i: number) => keep.slice(0, i).filter((k) => !k).length;
  const groups = b.groups
    ?.filter((g) => {
      if (unknownIn(g.label, inPool).length === 0) return true;
      drop("table group", g.label);
      return false;
    })
    .map((g) => ({ ...g, afterRow: g.afterRow - removedBefore(g.afterRow) }));

  return { ...b, rows: b.rows.filter((_, i) => keep[i]), ...(groups ? { groups } : {}) };
}

/**
 * A section label with the parts that state an unaccounted figure removed:
 * "Billing - 42 meters audited" becomes "Billing". The label is the slide's
 * title, so it is trimmed rather than dropped; if nothing survives, the
 * section key stands in.
 */
function groundLabel(label: string, key: string, inPool: (f: string) => boolean): string {
  if (unknownIn(label, inPool).length === 0) return label;
  const kept = label
    .split(/\s+[-–—|:]\s+|\s*\(|\)\s*/)
    .map((p) => p.trim())
    .filter((p) => p && unknownIn(p, inPool).length === 0);
  return kept.join(" - ") || key;
}

export function groundDeck(
  deck: CuratedDeck,
  source: string
): { deck: CuratedDeck; issues: GroundingIssue[] } {
  const pool = sourcePool(source);
  const inPool = (f: string) => pool.has(f);
  const issues: GroundingIssue[] = [];
  const allDerived = new Set<string>();

  const sections = deck.sections.map((section) => {
    const drop = (where: string, value: unknown) =>
      issues.push({ section: section.key, where, value: String(value), action: "dropped" });

    // Pass 1: the blocks that carry the parts - tables, then charts. A table
    // is checked against the source alone. A chart may plot a total of a
    // surviving table in this section, but never of itself.
    const tables = new Map<Block, Block | null>();
    for (const b of section.blocks) {
      if (b.type === "table") tables.set(b, groundTable(b, inPool, drop));
    }
    const tableTotals = derivable([...tables.values()].filter((t): t is Block => t !== null));
    const chartKnown = (f: string) => inPool(f) || tableTotals.has(f);

    const firstPass: Block[] = [];
    for (const b of section.blocks) {
      if (b.type === "table") {
        const t = tables.get(b);
        if (t) firstPass.push(t);
        continue;
      }
      if (b.type !== "chart") {
        firstPass.push(b);
        continue;
      }
      const values = b.series.flatMap((s) => s.values.map(String)).filter((v) => unknownIn(v, chartKnown).length > 0);
      const labels = [...b.categories, ...b.series.map((s) => s.name), b.unit ?? ""]
        .filter((t) => unknownIn(t, inPool).length > 0);
      if (values.length > 0 || labels.length > 0) {
        drop("chart", [...values, ...labels].join(", "));
        continue;
      }
      if (b.title && unknownIn(b.title, chartKnown).length > 0) {
        drop("chart title", b.title);
        firstPass.push({ ...b, title: undefined });
        continue;
      }
      firstPass.push(b);
    }

    // Pass 2: everything that may state a total, checked against totals of
    // what survived pass 1 only.
    const derived = derivable(firstPass);
    derived.forEach((d) => allDerived.add(d));
    const known = (f: string) => inPool(f) || derived.has(f);

    const blocks: Block[] = [];
    for (const b of firstPass) {
      switch (b.type) {
        case "kpis": {
          const items = b.items.filter((k) => {
            if (unknownIn(k.value, known).length === 0 && unknownIn(k.label, known).length === 0) return true;
            drop(`kpi "${k.label}"`, k.value);
            return false;
          });
          if (items.length > 0) blocks.push({ ...b, items });
          break;
        }
        case "bullets": {
          const items = b.items.filter((t) => {
            if (unknownIn(t, known).length === 0) return true;
            drop("bullet", t);
            return false;
          });
          if (items.length > 0) blocks.push({ ...b, items });
          break;
        }
        case "narrative":
        case "note": {
          if (unknownIn(b.text, known).length > 0) {
            drop(b.type, b.text);
            continue;
          }
          blocks.push(b);
          break;
        }
        default:
          blocks.push(b);
      }
    }

    const label = groundLabel(section.label, section.key, known);
    if (label !== section.label) drop("section title", section.label);
    return { ...section, label, blocks };
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
  // The glance tiles are meant to be copied from the sections; checked like
  // any KPI, value and label both.
  const kpis = (deck.summary.kpis ?? []).filter((k) => {
    if (summaryOk(k.value) && summaryOk(k.label)) return true;
    issues.push({ section: "summary", where: `kpi "${k.label}"`, value: k.value, action: "dropped" });
    return false;
  });

  return {
    deck: {
      ...deck,
      summary: { headline, points, ...(kpis.length > 0 ? { kpis } : {}) },
      sections: sections.filter((s) => s.blocks.length > 0),
    },
    issues,
  };
}
