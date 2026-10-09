import { groundDeck, type GroundingIssue } from "./grounding";
import { CUT_SHORT } from "./parsing/xlsx";

export type ChecklistItem = {
  key: string;
  label: string;
  askType: "upload" | "text" | "photos";
  /** Columns this parameter expects, from the real MMR data template.
   *  Lets Claude ask for exactly the right fields and check that a
   *  submitted table has the right shape - not merely that the section
   *  exists somewhere. */
  fields?: string[];
};

const CORE_RULES = `
What counts as genuinely present:
- A populated table, substantive narrative text, OR content photography. A section whose only content is photographs IS present — photo evidence of completed work is real content.
- A heading with nothing under it is NOT present.
- A name appearing only in a table of contents or agenda list is NOT present — that's an index, not content.
- A category name appearing only as a column header inside another section's table (e.g. "Total Attendance" inside a Training table) does NOT make that category present.
- Categories with overlapping words are distinct: "Security Training" and "Training" are separate parameters; one never satisfies the other.
- If the same data appears more than once (duplicate embedded workbooks, or overlapping files), treat it as one thing.

Absolute rule: never invent, estimate, round, or infer a number, name, or date that is not in the source content.
The only figures you may produce that are not written in the source are: the total of figures you show in the same section (a chart series, or a table column or row), and the count of rows or categories you show. Nothing else - no percentage, average, rate or difference the source does not state. Every figure is checked against the source after you respond, and any that cannot be accounted for is removed.
Each figure belongs to the period its own source row states. Where a workbook row carries a date, that date - not a slide heading - is the figure's period. Never place a figure under a month its own row does not name.
`;

function checklistBlock(items: ChecklistItem[]): string {
  return items
    .map((i) => {
      const fields = i.fields && i.fields.length > 0 ? ` [expects: ${i.fields.join(", ")}]` : "";
      return `${i.key} — ${i.label}${fields}`;
    })
    .join("\n");
}

// ---------------------------------------------------------------------
// Pass 1: strict site + month validation
// ---------------------------------------------------------------------
const VALIDATE_SYSTEM = `You are validating whether an uploaded submission actually belongs to the site and report month it was submitted against. The expected site and month are known for certain — they come from a single-use link issued to that site's manager. Your job is to catch a submission that belongs to a DIFFERENT site or month.

Read the whole content, not just titles. Site names and months appear in slide titles, headers, footers, table captions, sheet names, and body text.

Decide one of four verdicts:
- "match" — the content names this site and/or this month, and nothing contradicts either.
- "mismatch" — the content clearly names a DIFFERENT site, or a DIFFERENT report month. This is the case to catch. Quote the exact text that contradicts.
- "unidentified" — the content never names any site or month at all (for example a raw data workbook of pure figures). This is NOT a failure: with nothing to contradict the expected values, it passes. Treat it as acceptable.
- "period_conflict" — the slides or titles name this site and this month, but the dated figures in the data belong to a different period: workbook rows whose dates (shown as e.g. "Apr 2025" or "1 Apr 2025") do not include the report month at all. Typical cause: last period's deck copied forward with its embedded workbook left unchanged. Quote the dated rows and say which period they cover. A rolling comparison that includes the report month is NOT a conflict.

Be strict about genuine conflicts and relaxed about silence. Prefer "mismatch" when the report itself is plainly for another site or month; use "period_conflict" only when the report's own titles are right and the figures underneath are from another period. Do not report a mismatch because the site name is merely absent. Do not report a mismatch on a nearby month if the content is plainly a rolling comparison (e.g. a consumption table showing previous months alongside the current one) — only when the report's own subject month differs.

Respond with ONLY valid JSON, no fences:
{"verdict":"match","evidence":"Slide 1 title reads 'MMR - JUNE 2025 - AHUJA TOWER'"}
or
{"verdict":"mismatch","evidence":"Slide 1 title reads 'Altimus - May 2026', but this link is for Ahuja Tower, August 2026","conflictType":"site"}
or
{"verdict":"unidentified","evidence":"No site name or month appears anywhere in the content"}
or
{"verdict":"period_conflict","evidence":"Slides are titled August 2026, but the electricity, water and fit-out workbook rows are dated Apr 2025 to Jun 2025"}`;

export type ValidationResult = {
  verdict: "match" | "mismatch" | "unidentified" | "period_conflict";
  evidence: string;
  conflictType?: "site" | "month" | "both";
};

// ---------------------------------------------------------------------
// Pass 2: review against the site's resolved checklist
// ---------------------------------------------------------------------
function reviewSystem(items: ChecklistItem[]): string {
  return `You are reviewing content extracted from a facility management site's Monthly Management Report submission, which may span several files.

The parameters required for THIS site (they vary by client):
${checklistBlock(items)}
${CORE_RULES}
For each parameter, decide whether genuine content exists.

Where a parameter lists expected fields, use them: content that covers most of those fields is present; content naming the topic but missing the substantive figures is not. If the data is there but clearly incomplete against those fields, mark it present and note in "request" which fields are missing so they can top it up.

For anything absent, write a short specific request naming the exact fields needed - "Send the compliance register with Particular, Frequency, Validity Date and Days Before Expiry" rather than "send compliance data".

Respond with ONLY valid JSON, no fences:
{"sections":[{"key":"manpower","present":true},{"key":"compliance","present":false,"request":"Attach the compliance register showing licence validity dates for this month."}]}
Include every parameter listed above, in that order.`;
}

export type ReviewSection = { key: string; present: boolean; request?: string };

// ---------------------------------------------------------------------
// Pass 3: curate the deck
// ---------------------------------------------------------------------
function curateSystem(items: ChecklistItem[]): string {
  return `You are curating the final content of a facility management Monthly Management Report deck. You have the extracted source content, a manifest of available photographs, and any answers the site manager supplied for gaps.

The parameters relevant to this site:
${checklistBlock(items)}
${CORE_RULES}

IMPORTANT — what to include: build the deck ONLY from parameters that have genuine content. Omit anything absent entirely. Do not emit a section, a placeholder, or a "not available" slide for missing data — a shorter deck of real substance is the goal, not a fixed structure with holes. Order the sections you do include by what matters most for this month rather than by the checklist order, and lead with what a client would most want to see.

Chart what can be charted. Wherever a section's figures form a series — across months, across locations, across categories, or as parts of one total — give that section a chart block. A section of real series data shown only as a table reads as a data dump rather than a report. Chart only from data you have in full: a sheet ending "${CUT_SHORT}" has been cut short, so never count or total it.

Each section becomes one slide, holding several blocks stacked in order:
- {"type":"table","headers":[...],"rows":[[...]]} — real headers and rows exactly as in the source. At most 10 rows so it fits; if the source has more, choose the most representative and add a "note" block saying the table shows a selection. Give the source's total number of rows only where the source itself states it. Leave out any column that is empty in every row you show. A table has no "note" field - a caveat is always its own note block.
- {"type":"table","headers":[...],"rows":[[...]],"groups":[{"label":"Housekeeping","afterRow":0},{"label":"Technical","afterRow":5}]} — a table with category band rows. Use whenever rows fall into natural groups (service line, location, floor); grouped tables read far better than flat ones.
- {"type":"narrative","text":"..."} — 1-4 sentences, only what the source states.
- {"type":"bullets","items":["..."]} — 2-6 short points.
- {"type":"kpis","items":[{"label":"PPM Completion","value":"96%"}]} — 2-4 headline figures. Use where a section has standout numbers worth calling out above the detail.
- {"type":"chart","chartType":"bar","title":"Attendance by month","categories":["Jun","Jul","Aug"],"series":[{"name":"Deployed","values":[52,54,53]}],"unit":"staff"} — a native chart. Use one wherever a section's figures form a series: a metric across months, a quantity across locations or categories, or a breakdown of one total. Consumption, complaints, attendance, PPM and uptime sections almost always contain one.
  chartType is "bar" for comparing quantities across categories, "line" for a trend over time, "pie" for parts of one whole.
  "categories" and every series' "values" must be the same length: 2-12 categories, 1-3 series. "values" are plain numbers — no units, no commas, no percent signs; put the unit in "unit" instead. "title" is a short caption; omit it if the slide heading already says it.
  Put only series of comparable magnitude and unit on one chart. A series forty times larger than another flattens the smaller one into a straight line at the axis — split those into two charts. "pie" takes exactly one series, whose values are parts of a single total.
  A chart and a table of the same figures belong together on a slide when the exact per-row values are also worth reading — chart first, table beneath it. Use a table alone only when the rows are not comparable numbers at all.
  Chart only figures the source states. Do not interpolate a missing month, estimate a bar to complete a trend, or compute a percentage the source does not give; leave a category out rather than filling it in.
- {"type":"photos","imageIds":["img3","img4"]} — 2-6 ids from the manifest only. Choose photos whose source slide genuinely relates to this section.
- {"type":"note","text":"..."} — a short caveat for the client about what a figure covers, rendered small and muted.

Notes, the headline and summary points are read by the client. Anything about the submission itself - that it is demo or sample data, carried over from an earlier month, missing a month, or contradicts itself - goes in "submitterWarnings" instead, never in a note, the headline or a summary point. The submitter sees those warnings before filing; the client never does.

Also nominate a cover photograph as "coverImageId", choosing from the manifest the image that best represents the property as a whole - a building exterior, entrance, or lobby. Prefer a photo from a slide about the site itself over one from a maintenance or activity section; a photo of equipment or a work-in-progress makes a poor cover. Omit the field if nothing suitable exists.

Also write an opening "at a glance" summary: one headline sentence and 3-5 of the month's most significant, specific facts with real figures.

Respond with ONLY valid JSON, no fences:
{"summary":{"headline":"...","points":["..."]},"submitterWarnings":["..."],"coverImageId":"img3","sections":[{"key":"ppm","label":"PPM & Maintenance Schedule","blocks":[{"type":"kpis","items":[{"label":"Completed","value":"16 of 16"}]},{"type":"table","headers":["Location","Equipment"],"rows":[["10th Floor","LTG panel"]]}]}]}
Every section you include must have at least one block.`;
}

export type Block =
  | { type: "table"; headers: string[]; rows: string[][]; groups?: { label: string; afterRow: number }[] }
  | { type: "narrative"; text: string }
  | { type: "bullets"; items: string[] }
  | { type: "kpis"; items: { label: string; value: string }[] }
  | {
      type: "chart";
      chartType: "bar" | "line" | "pie";
      title?: string;
      categories: string[];
      series: { name: string; values: number[] }[];
      unit?: string;
    }
  | { type: "photos"; imageIds: string[] }
  | { type: "note"; text: string };

export type CuratedSection = { key: string; label: string; blocks: Block[] };
export type CuratedDeck = {
  summary: { headline: string; points: string[] };
  /** Photo id for the cover panel, chosen from the manifest. */
  coverImageId?: string;
  sections: CuratedSection[];
  /** Figures the grounding check removed or flagged. Not rendered; shown to the submitter. */
  grounding?: GroundingIssue[];
  /** About the submission, not the site: shown to the submitter, never drawn in the deck. */
  submitterWarnings?: string[];
};

async function callClaude(step: string, system: string, user: string, maxTokens: number): Promise<string> {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("ANTHROPIC_API_KEY isn't set in this deployment.");

  // A dropped connection surfaces from fetch as a bare "fetch failed", which
  // names neither the step nor the cause; both are added here.
  let res: Response;
  try {
    res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: maxTokens,
        system,
        messages: [{ role: "user", content: user }],
      }),
    });
  } catch (e) {
    const cause = (e as { cause?: { message?: string; code?: string } }).cause;
    throw new Error(
      `The ${step} request to Claude failed before any response arrived: ` +
        `${(e as Error).message}${cause ? ` (${cause.code ?? cause.message})` : ""}`
    );
  }

  if (!res.ok) throw new Error(`Claude API call for ${step} failed (${res.status}): ${(await res.text()).slice(0, 300)}`);

  const data = await res.json();
  const text = data.content?.find((b: { type: string }) => b.type === "text")?.text ?? "";
  if (!text) throw new Error(`Claude returned an empty response for ${step}.`);
  if (data.stop_reason === "max_tokens") {
    throw new Error(`Claude's ${step} response was cut off by the output limit before the JSON finished. Raise max_tokens in lib/curate.ts.`);
  }
  return text;
}

function parseJson<T>(raw: string, what: string): T {
  const cleaned = raw.replace(/```json/gi, "").replace(/```/g, "").trim();
  try {
    return JSON.parse(cleaned) as T;
  } catch {
    throw new Error(`Claude's ${what} response wasn't valid JSON. First 300 chars: ${cleaned.slice(0, 300)}`);
  }
}

export async function validateSiteMonth(
  combinedText: string,
  expectedSite: string,
  expectedMonth: string
): Promise<ValidationResult> {
  const user = `Expected site: ${expectedSite}\nExpected report month: ${expectedMonth}\n\n--- SUBMITTED CONTENT ---\n${combinedText}`;
  return parseJson<ValidationResult>(await callClaude("validation", VALIDATE_SYSTEM, user, 1500), "validation");
}

export async function reviewSubmission(
  combinedText: string,
  imageManifest: string,
  items: ChecklistItem[]
): Promise<ReviewSection[]> {
  const user = `--- AVAILABLE PHOTOGRAPHY ---\n${imageManifest}\n\n--- EXTRACTED SOURCE CONTENT ---\n${combinedText}`;
  return parseJson<{ sections: ReviewSection[] }>(await callClaude("review", reviewSystem(items), user, 8000), "review").sections;
}

/**
 * Coerces one charted value to a number, or null if it isn't one.
 *
 * Accepts the shapes a model reaches for despite being told not to — "1,240",
 * "96%", " 53 " — because stripping a comma or a percent sign returns the
 * figure the source already stated. It does NOT accept "N/A", "-" or an empty
 * cell: there is no number there, and substituting zero would put a figure in
 * a client report that nobody submitted.
 */
function chartValue(v: unknown): number | null {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v !== "string") return null;
  const cleaned = v.replace(/[,\s]/g, "").replace(/%$/, "");
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : null;
}

/**
 * Brings the model's blocks into the shapes the renderer can draw honestly.
 *
 * Charts are where a malformed response degrades into a plausible-looking
 * lie rather than a visible break: a series shorter than its categories
 * silently slides every remaining value onto the wrong month. So a chart
 * that does not hold together is never drawn as a chart. It used to be
 * dropped, and when it was a section's only block the whole section went
 * with it - one "NA" in a series could remove a parameter from the deck. Now
 * it becomes a table of exactly what the model gave, which shows a gap as a
 * gap. Every change is logged: a dropped or reshaped block is otherwise
 * invisible in the finished deck.
 */
export function sanitiseBlocks(deck: CuratedDeck): CuratedDeck {
  const sections = deck.sections.map((section) => {
    const out: Block[] = [];
    for (const block of section.blocks) {
      // The model attaches a "note" to tables and charts although neither
      // type has one, and nothing drew it: a real deck lost "AMC expired and
      // renewal not done for Fire system" and the caveat that its electricity
      // figures were from 2025. A note becomes its own block straight after,
      // so it is drawn, previewed and checked like any other.
      const attached = (block as { note?: unknown }).note;
      if (block.type !== "note" && typeof attached === "string" && attached.trim()) {
        // sanitiseBlock rebuilds each block from its known fields only.
        sanitiseBlock(block, section.key, out);
        out.push({ type: "note", text: attached.trim() });
        continue;
      }
      sanitiseBlock(block, section.key, out);
    }
    return { ...section, blocks: out };
  });

  const submitterWarnings = (Array.isArray(deck.submitterWarnings) ? deck.submitterWarnings : [])
    .map((w) => String(w ?? "").trim())
    .filter(Boolean);
  return { ...deck, submitterWarnings, sections: sections.filter((s) => s.blocks.length > 0) };
}

/**
 * A table whose every row has exactly one cell per header, with no column
 * that is empty in every row. A short row used to leave its last cells
 * unfilled and a long one drew a broken table over the slide title; an empty
 * column is a placeholder, and absent content is omitted, never shown.
 */
function sanitiseTable(block: Extract<Block, { type: "table" }>, key: string): Block | null {
  const text = (v: unknown) => (v == null ? "" : String(v));
  let headers = (Array.isArray(block.headers) ? block.headers : []).map(text);
  let rows = (Array.isArray(block.rows) ? block.rows : [])
    .filter(Array.isArray)
    .map((r) => r.map(text));
  const width = headers.length;
  if (width === 0 || rows.length === 0) {
    console.warn(`[curate] table in section "${key}" dropped: no headers or no rows`);
    return null;
  }
  rows = rows.map((r) => {
    if (r.length === width) return r;
    if (r.length < width) return [...r, ...Array(width - r.length).fill("")];
    // Cells past the last header are kept, in the last column, rather than lost.
    console.warn(`[curate] table in section "${key}": a row had ${r.length} cells for ${width} headers`);
    return [...r.slice(0, width - 1), r.slice(width - 1).filter((c) => c.trim()).join(" ")];
  });

  const used = headers.map((_, c) => rows.some((r) => r[c].trim() !== ""));
  if (used.includes(false)) {
    console.warn(
      `[curate] table in section "${key}": removed empty column(s) ` +
        headers.filter((_, c) => !used[c]).map((h) => `"${h}"`).join(", ")
    );
    headers = headers.filter((_, c) => used[c]);
    rows = rows.map((r) => r.filter((_, c) => used[c]));
  }
  if (headers.length === 0) return null;

  const groups = Array.isArray(block.groups)
    ? block.groups
        .filter((g) => g && typeof g.label === "string" && g.label.trim() && Number.isFinite(Number(g.afterRow)))
        .map((g) => ({ label: g.label.trim(), afterRow: Math.min(rows.length, Math.max(0, Math.floor(Number(g.afterRow)))) }))
    : undefined;
  return { type: "table", headers, rows, ...(groups && groups.length > 0 ? { groups } : {}) };
}

/** One block, made drawable, appended to `out` - or nothing, if it cannot be. */
function sanitiseBlock(block: Block, key: string, out: Block[]): void {
  switch (block.type) {
    case "table": {
      const t = sanitiseTable(block, key);
      if (t) out.push(t);
      return;
    }
    case "kpis": {
      // The renderer draws four cards a row; a fifth used to vanish. Split
      // instead, so every figure the model chose is still shown.
      const items = (Array.isArray(block.items) ? block.items : [])
        .filter((k) => k && k.value != null && String(k.value).trim())
        .map((k) => ({ label: String(k.label ?? ""), value: String(k.value) }));
      if (items.length > 4) {
        console.warn(`[curate] kpis in section "${key}": split ${items.length} cards into rows of four`);
      }
      for (let i = 0; i < items.length; i += 4) out.push({ type: "kpis", items: items.slice(i, i + 4) });
      return;
    }
    case "bullets": {
      const items = (Array.isArray(block.items) ? block.items : []).map((t) => String(t ?? "").trim()).filter(Boolean);
      if (items.length > 0) out.push({ type: "bullets", items });
      return;
    }
    case "narrative":
    case "note": {
      const text = typeof block.text === "string" ? block.text.trim() : "";
      if (text) out.push({ type: block.type, text });
      return;
    }
    case "chart":
      sanitiseChart(block, key, out);
      return;
    default:
      out.push(block);
  }
}

function sanitiseChart(block: Extract<Block, { type: "chart" }>, key: string, out: Block[]): void {
  // A unit is drawn as text, so anything else is either dropped or, for a
  // number, made text; the renderer called .trim() on it and a numeric unit
  // failed the build after the submitter had approved the preview.
  let unit = typeof block.unit === "string" || typeof block.unit === "number" ? String(block.unit).trim() : "";
  // chartValue strips a percent sign the model wrote into a value. The sign is
  // part of the figure, so when the chart names no unit it becomes the unit
  // rather than disappearing: 96% must not be drawn as a bare 96.
  const series0 = Array.isArray(block.series) ? block.series : [];
  if (!unit && series0.some((s) => Array.isArray(s?.values) && s.values.some((v) => typeof v === "string" && /%\s*$/.test(v)))) {
    unit = "%";
  }
  const title = typeof block.title === "string" && block.title.trim() ? block.title.trim() : undefined;
  const cats = Array.isArray(block.categories) ? block.categories.map((c) => String(c ?? "")) : [];

  // Shown as a table, the chart keeps everything the reader needs to read
  // its figures: the title over the label column, the unit beside each series.
  const asTable = (why: string) => {
    console.warn(`[curate] chart in section "${key}" shown as a table: ${why}`);
    if (cats.length === 0 || series0.length === 0) return;
    const head = (name: string) => (unit ? (name ? `${name} (${unit})` : unit) : name);
    const t = sanitiseTable(
      {
        type: "table",
        headers: [title ?? "", ...series0.map((s) => head(String(s?.name ?? "")))],
        rows: cats.map((c, i) => [
          c,
          ...series0.map((s) => (Array.isArray(s?.values) && s.values[i] != null ? String(s.values[i]) : "")),
        ]),
      },
      key
    );
    if (t) out.push(t);
  };

  if (!["bar", "line", "pie"].includes(block.chartType)) return asTable(`unknown chart type ${JSON.stringify(block.chartType)}`);
  if (series0.length === 0) return asTable("no series");
  if (cats.length < 2 || cats.length > 12) return asTable(`${cats.length} categories, needs 2-12`);

  // A pie is parts of one whole; extra series are meaningless on it.
  const max = block.chartType === "pie" ? 1 : 3;
  if (series0.length > max) {
    console.warn(
      `[curate] chart in section "${key}": kept ${max} of ${series0.length} series, ` +
        `dropped ${series0.slice(max).map((s) => `"${s?.name}"`).join(", ")}`
    );
  }

  const clean: { name: string; values: number[] }[] = [];
  for (const s of series0.slice(0, max)) {
    if (!Array.isArray(s?.values) || s.values.length !== cats.length) {
      return asTable(`series "${s?.name}" has ${s?.values?.length} values for ${cats.length} categories`);
    }
    const values = s.values.map(chartValue);
    const bad = values.findIndex((v) => v === null);
    if (bad !== -1) return asTable(`series "${s?.name}" value ${JSON.stringify(s.values[bad])} is not a number`);
    clean.push({ name: String(s.name ?? ""), values: values as number[] });
  }
  out.push({ type: "chart", chartType: block.chartType, title, categories: cats, series: clean, unit: unit || undefined });
}


/**
 * Keeps only photographs that exist.
 *
 * The model picks ids from a manifest, but nothing stopped it naming one that
 * is not there, or the same one twice. The renderer skipped unknown ids while
 * the preview drew a tile for each, so the submitter approved a photo grid
 * that came out smaller in the filed deck. Checked once here, both agree.
 */
export function sanitisePhotos(deck: CuratedDeck, knownIds: string[]): CuratedDeck {
  const known = new Set(knownIds);
  const sections = deck.sections
    .map((section) => ({
      ...section,
      blocks: section.blocks.flatMap((b): Block[] => {
        if (b.type !== "photos") return [b];
        const ids = [...new Set((Array.isArray(b.imageIds) ? b.imageIds : []).filter((id) => known.has(id)))].slice(0, 6);
        const dropped = (b.imageIds ?? []).length - ids.length;
        if (dropped > 0) console.warn(`[curate] photos in "${section.key}": ${dropped} id(s) unknown, repeated or over six`);
        return ids.length > 0 ? [{ ...b, imageIds: ids }] : [];
      }),
    }))
    .filter((s) => s.blocks.length > 0);
  const coverImageId = deck.coverImageId && known.has(deck.coverImageId) ? deck.coverImageId : undefined;
  return { ...deck, coverImageId, sections };
}

export async function curateDeck(
  combinedText: string,
  imageManifest: string,
  items: ChecklistItem[],
  siteName: string,
  reportMonth: string,
  userAnswers: { label: string; answer: string }[]
): Promise<CuratedDeck> {
  const answers =
    userAnswers.length > 0
      ? userAnswers.map((a) => `- ${a.label}: ${a.answer}`).join("\n")
      : "(no gap answers were supplied)";
  const user = `Site: ${siteName}\nReport month: ${reportMonth}\n\n--- AVAILABLE PHOTOGRAPHY ---\n${imageManifest}\n\n--- ANSWERS SUPPLIED FOR GAPS ---\n${answers}\n\n--- EXTRACTED SOURCE CONTENT ---\n${combinedText}`;
  const deck = sanitiseBlocks(parseJson<CuratedDeck>(await callClaude("curation", curateSystem(items), user, 20000), "curation"));

  // The submitter's typed answers are source material too: a figure they
  // supplied for a gap is theirs, not the model's.
  const source = [combinedText, ...userAnswers.map((a) => a.answer)].join("\n");
  const { deck: grounded, issues } = groundDeck(deck, source);
  for (const i of issues) {
    console.warn(`[grounding] ${i.action} ${i.section} / ${i.where}: ${i.value.slice(0, 120)}`);
  }
  return { ...grounded, grounding: issues };
}
