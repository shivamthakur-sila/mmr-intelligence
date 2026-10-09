/**
 * The chapters a deck's sections are grouped into, and the order they run in.
 *
 * Human MMRs open with an index of numbered chapters, each introduced by a
 * divider slide, in a fixed operational order - Altimus's runs through ten.
 * A section's key is the checklist parameter it reports on (lib/curate.ts
 * asks for exactly that), so the mapping lives here, in code, rather than
 * being chosen by the model each month: a chapter mislabelled in the index is
 * the kind of error a human deck has and this one must not.
 *
 * Pure data, safe to import from a client component.
 */

export type Chapter = { name: string; order: number };

const CHAPTERS: { name: string; keys: string[]; words: RegExp }[] = [
  { name: "Highlights", keys: ["highlights", "lowlights", "achievements", "events"], words: /highlight|lowlight|achievement|accolade|challenge|event/i },
  { name: "People & Training", keys: ["manpower", "training", "security_training", "rewards"], words: /manpower|attendance|training|reward|recognition|staff/i },
  { name: "Helpdesk & Complaints", keys: ["complaints", "customer_centricity", "footfall", "traffic"], words: /complaint|helpdesk|help desk|service request|customer|footfall|visitor|traffic/i },
  { name: "Technical & PPM", keys: ["ppm", "amc", "equipment_uptime", "technical_activity", "lift_breakdowns", "hvac_peak"], words: /ppm|preventive|amc|uptime|breakdown|technical|maintenance|lift|hvac|elevator/i },
  { name: "Utilities & Energy", keys: ["electricity", "water", "diesel", "solar", "stp", "ghg"], words: /electric|power|energy|water|diesel|\bdg\b|solar|stp|ghg|carbon|utilit/i },
  { name: "Soft Services", keys: ["soft_service_activity", "pest", "hkmachinery", "waste"], words: /soft service|housekeeping|pest|cleaning|waste|machinery/i },
  { name: "HSE & Compliance", keys: ["compliance", "incidents", "mock_drills", "security_incidents", "ambulance", "safety_remarks", "monsoon"], words: /complian|statutory|incident|safety|hse|drill|security|monsoon|fire|ambulance/i },
  { name: "Fit-outs & Projects", keys: ["fitout", "hoto", "capex", "occupancy", "clubhouse"], words: /fit-?out|hoto|capex|project|occupancy|tenant|clubhouse/i },
  { name: "Commercials", keys: ["payments", "cam_budget"], words: /payment|outstanding|budget|cam\b|invoice|cost/i },
  { name: "Way Forward", keys: ["wayforward", "mom"], words: /way forward|next steps|minutes|plan for|upcoming/i },
];

const OTHER: Chapter = { name: "Other Updates", order: CHAPTERS.length };

/** The chapter a section belongs to: by its checklist key, else by words in its label. */
export function chapterOf(section: { key: string; label: string }): Chapter {
  const byKey = CHAPTERS.findIndex((c) => c.keys.includes(section.key));
  if (byKey !== -1) return { name: CHAPTERS[byKey].name, order: byKey };
  const byWords = CHAPTERS.findIndex((c) => c.words.test(`${section.key.replace(/_/g, " ")} ${section.label}`));
  if (byWords !== -1) return { name: CHAPTERS[byWords].name, order: byWords };
  return OTHER;
}

/**
 * Sections in chapter order. Within a chapter they keep the order the model
 * gave them, which leads with what mattered most that month.
 */
export function inChapterOrder<T extends { key: string; label: string }>(sections: T[]): T[] {
  return sections
    .map((s, i) => ({ s, i, c: chapterOf(s).order }))
    .sort((a, b) => a.c - b.c || a.i - b.i)
    .map((x) => x.s);
}

/**
 * Whether a deck gets chapter dividers. They are navigation: worth a slide
 * when chapters run to several pages, as in the reference decks (Altimus has
 * ten dividers in 54 slides), and padding when most chapters are a single
 * slide - a nine-section deck in six chapters went from 13 slides to 19.
 */
export function wantsDividers(chapters: number, sectionSlides: number): boolean {
  return chapters >= 4 && sectionSlides >= 2 * chapters;
}
