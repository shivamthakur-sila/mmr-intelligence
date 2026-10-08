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

Decide one of three verdicts:
- "match" — the content names this site and/or this month, and nothing contradicts either.
- "mismatch" — the content clearly names a DIFFERENT site, or a DIFFERENT report month. This is the case to catch. Quote the exact text that contradicts.
- "unidentified" — the content never names any site or month at all (for example a raw data workbook of pure figures). This is NOT a failure: with nothing to contradict the expected values, it passes. Treat it as acceptable.

Be strict about genuine conflicts and relaxed about silence. Do not report a mismatch because the site name is merely absent. Do not report a mismatch on a nearby month if the content is plainly a rolling comparison (e.g. a consumption table showing previous months alongside the current one) — only when the report's own subject month differs.

Respond with ONLY valid JSON, no fences:
{"verdict":"match","evidence":"Slide 1 title reads 'MMR - JUNE 2025 - AHUJA TOWER'"}
or
{"verdict":"mismatch","evidence":"Slide 1 title reads 'Altimus - May 2026', but this link is for Ahuja Tower, August 2026","conflictType":"site"}
or
{"verdict":"unidentified","evidence":"No site name or month appears anywhere in the content"}`;

export type ValidationResult = {
  verdict: "match" | "mismatch" | "unidentified";
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

Chart what can be charted. Wherever a section's figures form a series — across months, across locations, across categories, or as parts of one total — give that section a chart block. A deck of this kind normally carries several; one made entirely of tables reads as a data dump rather than a report.

Each section becomes one slide, holding several blocks stacked in order:
- {"type":"table","headers":[...],"rows":[[...]]} — real headers and rows exactly as in the source. At most 10 rows so it fits; if the source has more, choose the most representative and add a "note" block stating the true total.
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
- {"type":"note","text":"..."} — a short caveat, rendered small and muted.

Also nominate a cover photograph as "coverImageId", choosing from the manifest the image that best represents the property as a whole - a building exterior, entrance, or lobby. Prefer a photo from a slide about the site itself over one from a maintenance or activity section; a photo of equipment or a work-in-progress makes a poor cover. Omit the field if nothing suitable exists.

Also write an opening "at a glance" summary: one headline sentence and 3-5 of the month's most significant, specific facts with real figures.

Respond with ONLY valid JSON, no fences:
{"summary":{"headline":"...","points":["..."]},"coverImageId":"img3","sections":[{"key":"ppm","label":"PPM & Maintenance Schedule","blocks":[{"type":"kpis","items":[{"label":"Completed","value":"16 of 16"}]},{"type":"table","headers":["Location","Equipment"],"rows":[["10th Floor","LTG panel"]]}]}]}
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
};

async function callClaude(system: string, user: string, maxTokens: number): Promise<string> {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("ANTHROPIC_API_KEY isn't set in this deployment.");

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "content-type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    body: JSON.stringify({
      model: "claude-sonnet-5",
      max_tokens: maxTokens,
      system,
      messages: [{ role: "user", content: user }],
    }),
  });

  if (!res.ok) throw new Error(`Claude API call failed (${res.status}): ${(await res.text()).slice(0, 300)}`);

  const data = await res.json();
  const text = data.content?.find((b: { type: string }) => b.type === "text")?.text ?? "";
  if (!text) throw new Error("Claude returned an empty response.");
  if (data.stop_reason === "max_tokens") {
    throw new Error("Claude's response was cut off by the output limit before the JSON finished. Raise max_tokens in lib/curate.ts.");
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
  return parseJson<ValidationResult>(await callClaude(VALIDATE_SYSTEM, user, 1500), "validation");
}

export async function reviewSubmission(
  combinedText: string,
  imageManifest: string,
  items: ChecklistItem[]
): Promise<ReviewSection[]> {
  const user = `--- AVAILABLE PHOTOGRAPHY ---\n${imageManifest}\n\n--- EXTRACTED SOURCE CONTENT ---\n${combinedText}`;
  return parseJson<{ sections: ReviewSection[] }>(await callClaude(reviewSystem(items), user, 8000), "review").sections;
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
 * Drops chart blocks the renderer cannot draw honestly.
 *
 * A chart is the one block type where a malformed response degrades into a
 * plausible-looking lie rather than a visible break: a series shorter than its
 * categories silently slides every remaining value onto the wrong month. So a
 * chart that doesn't hold together is removed outright — the documented
 * behaviour for content that isn't really there — rather than patched up.
 */
function sanitiseCharts(deck: CuratedDeck): CuratedDeck {
  const sections = deck.sections.map((section) => ({
    ...section,
    blocks: section.blocks.filter((block) => {
      if (block.type !== "chart") return true;

      // A dropped chart is invisible in the finished deck, so say why.
      const drop = (why: string) => {
        console.warn(`[curate] dropped chart in section "${section.key}": ${why}`);
        return false;
      };

      const cats = Array.isArray(block.categories) ? block.categories.map((c) => String(c ?? "")) : [];
      if (cats.length < 2 || cats.length > 12) return drop(`${cats.length} categories, needs 2-12`);
      if (!Array.isArray(block.series) || block.series.length === 0) return drop("no series");

      // A pie is parts of one whole; extra series are meaningless on it.
      const series = block.chartType === "pie" ? block.series.slice(0, 1) : block.series.slice(0, 3);

      const clean: { name: string; values: number[] }[] = [];
      for (const s of series) {
        if (!Array.isArray(s?.values) || s.values.length !== cats.length) {
          return drop(`series "${s?.name}" has ${s?.values?.length} values for ${cats.length} categories`);
        }
        const values = s.values.map(chartValue);
        const bad = values.findIndex((v) => v === null);
        if (bad !== -1) return drop(`series "${s?.name}" value ${JSON.stringify(s.values[bad])} is not a number`);
        clean.push({ name: String(s.name ?? ""), values: values as number[] });
      }

      block.categories = cats;
      block.series = clean;
      return true;
    }),
  }));

  // A section whose only block was a rejected chart has nothing left to show.
  return { ...deck, sections: sections.filter((s) => s.blocks.length > 0) };
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
  return sanitiseCharts(parseJson<CuratedDeck>(await callClaude(curateSystem(items), user, 20000), "curation"));
}