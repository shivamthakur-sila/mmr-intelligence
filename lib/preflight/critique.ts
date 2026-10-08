import Anthropic from "@anthropic-ai/sdk";
import fs from "fs";
import { RUBRIC, type Severity } from "./rubric";
import type { RenderedSlide } from "./render";

/**
 * Looks at the rendered slides and reports what is wrong with them.
 *
 * The deck is sent as images, not as the spec it was built from, on purpose:
 * the spec is what we already believed, and every defect this generator has
 * shipped was one the spec looked fine about. The model is asked to judge the
 * picture.
 */

export type Finding = {
  slide: number;
  severity: Severity;
  rule: string;
  issue: string;
  fix: string;
};

const FINDINGS_SCHEMA = {
  type: "object" as const,
  additionalProperties: false,
  required: ["findings"],
  properties: {
    findings: {
      type: "array" as const,
      items: {
        type: "object" as const,
        additionalProperties: false,
        required: ["slide", "severity", "rule", "issue", "fix"],
        properties: {
          slide: { type: "number" as const, description: "1-based slide number as labelled in the prompt" },
          severity: { type: "string" as const, enum: ["blocker", "major", "minor"] },
          rule: { type: "string" as const, description: "Rubric code, e.g. C1" },
          issue: { type: "string" as const, description: "What is visibly wrong, in one sentence" },
          fix: { type: "string" as const, description: "The smallest change that would resolve it" },
        },
      },
    },
  },
};

/** Slides per request. Keeps any one call well inside the image limit and
 *  stops a long deck from blurring into one undifferentiated judgement. */
const BATCH = 6;

const SEVERITIES: readonly Severity[] = ["blocker", "major", "minor"];

/**
 * Checks the parsed response against the shape the schema promised. The
 * schema constrains generation, but this result decides whether a deck goes
 * to a client, so it is checked rather than trusted: a malformed finding
 * would otherwise be counted as no finding at all.
 */
function toFindings(parsed: unknown, batch: RenderedSlide[], range: string): Finding[] {
  const findings = (parsed as { findings?: unknown } | null)?.findings;
  if (typeof parsed !== "object" || parsed === null || !Array.isArray(findings)) {
    throw new Error(`Preflight for ${range} returned JSON without a findings array.`);
  }
  const pages = new Set(batch.map((s) => s.page));
  const text = (v: unknown) => typeof v === "string" && v.trim() !== "";
  return findings.map((f, i) => {
    const item = f as Partial<Record<keyof Finding, unknown>> | null;
    // A slide number outside this batch cannot be attributed to any slide the
    // model was shown, so it is as unusable as a missing one.
    const ok =
      typeof item === "object" && item !== null &&
      typeof item.slide === "number" && pages.has(item.slide) &&
      SEVERITIES.includes(item.severity as Severity) &&
      text(item.rule) && text(item.issue) && text(item.fix);
    if (!ok) {
      throw new Error(`Preflight for ${range} returned a malformed finding at index ${i}: ${JSON.stringify(f)}`);
    }
    return item as Finding;
  });
}

async function critiqueBatch(
  client: Anthropic,
  batch: RenderedSlide[],
  model: string
): Promise<Finding[]> {
  const content: Anthropic.ContentBlockParam[] = [];
  for (const s of batch) {
    content.push({ type: "text", text: `Slide ${s.page}:` });
    content.push({
      type: "image",
      source: { type: "base64", media_type: "image/png", data: fs.readFileSync(s.file).toString("base64") },
    });
  }
  content.push({
    type: "text",
    text:
      "Inspect each slide above against the rubric. Report every fault you can " +
      "see, using the slide number given before each image. If a slide is sound, " +
      "report nothing for it.",
  });

  const response = await client.messages.create({
    model,
    max_tokens: 16000,
    system: RUBRIC,
    output_config: {
      effort: "high",
      format: { type: "json_schema", schema: FINDINGS_SCHEMA },
    },
    messages: [{ role: "user", content }],
  });

  // Every failure below throws. This is a gate, so an answer that cannot be
  // read must stop the deck rather than pass it with no findings.
  const first = batch[0].page;
  const last = batch[batch.length - 1].page;
  const range = first === last ? `slide ${first}` : `slides ${first}-${last}`;

  if (response.stop_reason === "refusal") {
    throw new Error(`Preflight was declined for ${range}: ${response.stop_details?.explanation ?? "no explanation"}`);
  }
  if (response.stop_reason === "max_tokens") {
    throw new Error(`Preflight response for ${range} hit the output limit before the findings finished.`);
  }
  if (response.stop_reason !== "end_turn") {
    throw new Error(`Preflight for ${range} stopped early (stop_reason: ${response.stop_reason}), so its findings are incomplete.`);
  }

  const text = response.content
    .filter((b): b is Anthropic.TextBlock => b.type === "text")
    .map((b) => b.text)
    .join("");
  if (!text.trim()) {
    throw new Error(`Preflight for ${range} returned no text, so those slides were not judged.`);
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    throw new Error(`Preflight for ${range} returned text that is not valid JSON (${(e as Error).message}): ${text.slice(0, 200)}`);
  }
  return toFindings(parsed, batch, range);
}

/** `client` is a parameter so the response handling can be exercised
 *  without calling the API; production passes nothing. */
export async function critiqueSlides(
  slides: RenderedSlide[],
  model = "claude-opus-5-5",
  client: Anthropic = new Anthropic()
): Promise<Finding[]> {
  const out: Finding[] = [];
  for (let i = 0; i < slides.length; i += BATCH) {
    out.push(...(await critiqueBatch(client, slides.slice(i, i + BATCH), model)));
  }
  const rank: Record<Severity, number> = { blocker: 0, major: 1, minor: 2 };
  return out.sort((a, b) => rank[a.severity] - rank[b.severity] || a.slide - b.slide);
}
