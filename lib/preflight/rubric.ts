/**
 * What a finished slide has to satisfy.
 *
 * Written down rather than left to "does this look right" so the check is
 * repeatable and so a disagreement about output is a disagreement about a
 * numbered rule. Sources: SILA's brand guide (colours, type scale, logo,
 * closing slide), the house decks it is drawn from, and the defects this
 * generator has actually produced — every FATAL rule below exists because
 * something shipped that broke it.
 */

export const BRAND = {
  fmBlue: "#264170",
  sunshine: "#F7A328",
  slate: "#2D2D2D",
  hairline: "#E5E5E4",
} as const;

export type Severity = "blocker" | "major" | "minor";

export const RUBRIC = `
You are inspecting rendered slides from a SILA Facility Management Monthly
Management Report, before it is sent to a client. Judge only what you can see
in the image.

FRAME — every content slide
F1. A solid blue (#264170) band runs across the very top.
F2. The SILA logo sits in that band, at the right.
F3. The slide title is top-left, in a serif face, large — clearly a title and
    not a caption — with a short orange (#F7A328) rule directly beneath it.
F4. A small orange box sits at the bottom-right carrying the page number.
F5. The background is white. Orange is never a full-slide background.

CONTENT
C1. Nothing is cut off at any edge, and nothing overlaps anything else.
    A table, chart or photo running past the bottom of the slide is a blocker.
C2. No element is squashed to the point of illegibility — a chart under about
    a fifth of the slide height, or axis labels colliding, is a major fault.
C3. Text is readable against whatever sits behind it. Dark text on a dark fill
    or light text on a pale fill is a major fault.
C4. Photographs keep their proportions. A face or a building stretched or
    squeezed is a blocker.
C5. The slide is not empty, and not a lone sentence floating in white space.
C6. Charts: bars start at a zero baseline; every series is distinguishable;
    no two adjacent slices or bars share a colour; data labels are legible and
    do not overlap each other or the plot.
C7. Tables: header row is filled blue with white text; no column is so narrow
    its text wraps to a sliver; no row is clipped.

PALETTE AND TYPE
P1. Only brand colour appears: FM Blue #264170, Sunshine #F7A328, Seashell
    Grey #D9D9D9, white, near-black text, light grey rules — plus lighter
    tints and darker shades of FM Blue and Sunshine, and neutral greys, which
    are sanctioned and are how a chart of three or more series or slices is
    coloured. A tint, shade or grey is not a fault; judge whether
    neighbouring series are TELLABLE APART, not whether the colour is one of
    the two base hues. Do flag genuinely foreign colour: a default Office
    blue/orange/grey/yellow chart palette, a dark grey chart background, a
    3-D bar effect, or a hue outside the blue/orange/neutral family.
P2. Headings are serif; body and labels are sans. Body text is not smaller
    than roughly 10pt on a 13.3-inch-wide slide.

SUBSTANCE
S1. No placeholder text, no "lorem", no "undefined", no "null", no "NaN", no
    "[object Object]", no visible template markers.
S2. Nothing that looks machine-leftover: a caption baked into a source image
    that contradicts the slide, a stray "Series1"/"Series2" legend, a label
    that is plainly an internal key rather than a human phrase.

Severities:
- "blocker": a client must not see this. Clipping, overlap, distortion,
  unreadable text, a fabricated-looking artefact.
- "major": visibly wrong or off-brand, but the slide still communicates.
- "minor": a polish point.

Report only what is actually visible in the image. Do not infer problems from
what you imagine the underlying data to be, and do not report the same fault
twice for one slide.
`.trim();
