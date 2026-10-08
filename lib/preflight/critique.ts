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

  if (response.stop_reason === "refusal") {
    throw new Error(`Preflight was declined: ${response.stop_details?.explanation ?? "no explanation"}`);
  }
  if (response.stop_reason === "max_tokens") {
    throw new Error("Preflight response hit the output limit before the findings finished.");
  }

  const text = response.content.find((b): b is Anthropic.TextBlock => b.type === "text")?.text ?? "";
  if (!text.trim()) return [];
  return (JSON.parse(text) as { findings: Finding[] }).findings ?? [];
}

export async function critiqueSlides(
  slides: RenderedSlide[],
  model = "claude-opus-5-5"
): Promise<Finding[]> {
  const client = new Anthropic();
  const out: Finding[] = [];
  for (let i = 0; i < slides.length; i += BATCH) {
    out.push(...(await critiqueBatch(client, slides.slice(i, i + BATCH), model)));
  }
  const rank: Record<Severity, number> = { blocker: 0, major: 1, minor: 2 };
  return out.sort((a, b) => rank[a.severity] - rank[b.severity] || a.slide - b.slide);
}
