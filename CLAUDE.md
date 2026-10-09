
# CLAUDE.md — MMR Intelligence

Project-specific rules. General engineering philosophy lives in
`~/.claude/CLAUDE.md`; this file is only what is true of *this* codebase.

## Stack

Next.js App Router (TypeScript, strict) · Supabase Postgres + Storage ·
Anthropic API · pptxgenjs · sharp · deployed on Vercel.

n8n owns Agent 0 and the confirmation email. It is not in this repo but
shares the database — see *Contract with n8n* below.

## The rule that outranks everything

**Never invent a number, name or date.** If it is not in the submitted
source or supplied by the submitter, it does not appear in any output. A
fabricated figure in a client-facing report is the worst failure this
system can produce. Every prompt that touches content restates this; keep
it that way.

Corollaries: content that is absent is omitted from the deck entirely,
never rendered as a placeholder. A section whose only content is
photographs *is* present — photo evidence is real content.

## Hard constraints, each learned by something breaking

**Files never transit an API request body.** Vercel rejects bodies over
~4.5MB at the edge before any code runs; real submissions reach 26MB. The
browser uploads directly to Supabase Storage via signed URLs
(`lib/upload-via-storage.ts`) and only paths are posted to the API. Do not
add a multipart upload endpoint.

**Nothing may persist on local disk between requests.** Each route is a
separate serverless invocation with its own disposable `/tmp`. Session
state, extracted photographs and built decks all live in Storage
(`lib/session.ts`). Writing to `/tmp` is fine *within* one invocation —
`materializeImages()` does this deliberately because pptxgenjs embeds
images from file paths.

**pptxgenjs `sizing` cannot be trusted.** When `w` and `h` are both passed,
`sizing` is ignored and the image is stretched. Verified by rendering a
known circle through every combination. Always centre-crop the image to the
box ratio first (`cropToRatio`), then place it with plain `w`/`h`.

**A chart palette must be at least as long as its data.** pptxgenjs colours
pie slices — and the bars of a *single-series* bar chart — per data point,
and once the index runs past the end of `chartColors` it fills the rest with
`Math.random()` picks. A five-slice pie against the two brand colours was
generated six times and produced six different colour assignments, some with
touching slices the same colour. So `chartPalette(n)` returns exactly one
colour per point, and a single-series bar chart is passed exactly one colour
(more than one and every bar is coloured separately, as if the months were
different categories). Verified by rendering. Distinct hex codes are not
enough: the twelve colours are chosen so every pair is at least 10 apart in
CIEDE2000, because tints of two hues alone gave a 10-slice pie two oranges a
reader could not tell apart.

**`dataLabelPosition: "outEnd"` works on a pie, not on a doughnut.** On a
doughnut it is ignored and labels are forced inside the ring, where no single
colour is readable against both the dark and the pale ends of the palette.
That is the only reason breakdowns render as a pie.

**A section that overflows spills onto a continuation slide; it is never
truncated.** `lib/deck-layout.ts` measures the blocks and decides where a
section breaks, and both the renderer and the browser preview use it, so the
slide the submitter approves is the slide that gets filed. This exists
because the overflow used to be dropped in silence: on a real submission two
of the five charts curation had chosen never appeared in the deck, with
nothing in the file or the logs to say so. If you change a block's height
estimate, change it there — a table row is costed by how many lines its
widest cell wraps to, not as a flat per-row figure.

**A line chart needs its axis bounds set explicitly.** Left automatic the
value axis started at zero: consumption moving 27,240 to 28,920 rendered as a
dead flat rule. `niceAxis()` in `lib/chart-style.ts` pads the real range and rounds outwards.
A truncated axis is honest on a line, which encodes by position — but not on
a bar, which encodes by length and therefore keeps `valAxisMinVal: 0`. Axis
labels take their decimal places from the step (`axisDecimals`), never from
the data: gridlines 0.0005 apart labelled to three places read 0.995, 0.995,
0.996 — duplicate ticks and a top label nobody submitted.

**Slide XML text must be parsed as text.** fast-xml-parser's defaults turn a
run that is only a number into a number and trim a run that is only a space.
The text collector keeps strings, so every numeric cell in a slide table — a
Sr. No. column, a reading, a headcount — reached the model empty. `pptx.ts`
sets `parseTagValue: false, trimValues: false`; keep it.

**SheetJS comes from its CDN, not npm.** The npm `xlsx` package stopped at
0.18.5, which has open advisories (prototype pollution, ReDoS — we parse
uploaded files) and decodes ISO date cells in the server's local timezone, so
1 Apr 2025 became March on an IST machine. `package.json` points at the
0.20.3 tarball on cdn.sheetjs.com.

**An extractor annotation must never state a figure it cannot state
correctly.** Sheets used to be headed "(N real data rows)", with N counting
the header, title and TOTAL rows; once a sheet was cut short that wrong count
was the only total the model saw, and grounding accepted it because it was
in the source. Sheets now carry no count, a cut one ends with `CUT_SHORT`,
and grounding strips the extractor's own slide numbers and file names.

**Escape apostrophes before interpolating into SQL.** n8n workflows build
SQL by string interpolation, and an apostrophe in an email body has already
broken production once. In this repo, prefer the Supabase client's
parameterised methods; where raw SQL is unavoidable, escape explicitly.

## Data and security

- `SUPABASE_SERVICE_ROLE_KEY` and `ANTHROPIC_API_KEY` are server-side only.
  Never prefix either with `NEXT_PUBLIC_`.
- The token flow uses the service-role client because there is no user
  session. It therefore must validate the token itself and scope every
  query to the site and month that token resolves to. Never widen a query
  in that flow beyond the resolved token.
- `resolveToken()` is the single gate for the upload flow. Every route
  under `app/api/upload/[token]/` calls it first, including `download`.
  It refuses unknown, used, expired, and already-filed.
- Agent 2 must resolve scope from `user_site_access` *before* retrieval.
  Never retrieve broadly and filter afterwards.
- Google sign-in is restricted to `@silagroup.co.in`. The `hd` parameter on
  the login page is a hint only; the real check is server-side in the
  callback and must stay.

## Contract with n8n — do not break silently

n8n reads and writes these directly. Changing them breaks workflows that
live in another repo, with no compile error to warn you:

| Column | Who | Meaning |
|---|---|---|
| `mmr_status.state = 'received'` | written here, read by Agent 0 | Stops the chaser. The single most important write in the system. |
| `mmr_status.confirmation_sent_at` | written by n8n | Null means filed but not yet confirmed |
| `upload_tokens.used_at` | written here | Burns the link |
| `reminder_log` | written by Agent 0 | Prevents duplicate chasing |
| `site_contacts.role` | read by both | `person_in_charge` / `site_head` / `hod` |

## Working with the model

Three passes, deliberately separate (`lib/curate.ts`): validate site and
month, review against the checklist, then curate the deck. Keep them
separate — a wrong-site submission must be refused before any effort is
spent reviewing it.

- **Checklists are data, not prompt text.** Parameters come from
  `site_config` via `resolve_site_config`, resolved site → client →
  default, filtered on `enabled`. Never hardcode a parameter list.
- **Pass field specs into prompts.** Each parameter carries the columns it
  expects. This is what makes a gap request say *"send Particular,
  Frequency, Validity Date"* rather than *"send compliance data"*.
- **Validation is strict about conflicts, relaxed about silence.** Content
  naming a different site or month is refused. Content naming neither
  passes — a raw figures workbook is legitimate.
- **Output limits are a real failure mode.** Truncated JSON has broken
  production. Check `stop_reason === 'max_tokens'` and surface it as
  itself, never as a JSON parse error. Current limits: validate 1500,
  review 8000, curate 20000.
- **Model output is a structural contract.** Curation returns JSON the
  renderer consumes directly. Adding a block type means changing the
  prompt, the TypeScript type, the renderer, and the browser preview
  together, or it fails at runtime.

## Verification

**The deck is verified by looking at it.** Type-checking and a clean build
prove nothing about a slide. Every bug in the generator — stretched photos,
a dropped table, a vanished photo grid, an invisible logo, a deleted
`addImage` — was found by rendering and looking, and none would have been
caught any other way.

```bash
# once: the render image with the real deck fonts
docker build -t mmr-render render/
# render, contact sheets, and how full each slide is
python render/bench.py deck.pptx out-dir
```

Then actually view the images. Render only with `mmr-render`: the stock
LibreOffice image substitutes Noto for Georgia and Calibri, about 10-15%
wider, and a layout judged on it is judged on text PowerPoint never draws.
The bench's fill figures catch emptiness only - Sattva's medians (vertical
fill 0.97, ink 0.36) are the bar - so the slides are still looked at.

- **Verify against real submissions, not fixtures.** The real files contain
  grouped shapes, duplicate embedded workbooks, portrait photographs and
  table-of-contents slides that a clean fixture will not.
- **When a heuristic is involved, test the case you know the answer to.**
  There is a real MMR with a blank Attendance slide and a real PPM table;
  any change to matching must still get both right.
- **Errors must name the failing step.** "Couldn't reach the server" hid a
  413 for hours. Surface the actual status and message.
- **Verify against the live schema before writing to it.** A column name
  guessed from memory (`file_name` for `source_filename`) would have failed
  at the final step of the user journey. Check `information_schema`, or test
  the insert inside a transaction and roll it back.

## Local development

```bash
npm install
npm run dev          # localhost:3000
```

`.env.local` needs: `NEXT_PUBLIC_SUPABASE_URL`,
`NEXT_PUBLIC_SUPABASE_ANON_KEY`, `SUPABASE_SERVICE_ROLE_KEY`,
`ANTHROPIC_API_KEY`, `NEXT_PUBLIC_AGENT2_WEBHOOK_URL`.

The upload flow needs a live token. Reset a site for testing with the four
deletes in `docs/reset-site.sql`, then trigger Agent 0.

Deploy with `vercel --prod`. Plain `vercel` builds a preview and does not
update the live domain — the emailed links point at the live domain, so a
preview deploy looks like a 404 to anyone testing.

<!-- BEGIN:nextjs-agent-rules -->

# This is NOT the Next.js you know

This version has breaking changes — APIs, conventions, and file structure may all differ from your training data. Read the relevant guide in `node_modules/next/dist/docs/` (resolved from this file's directory; in monorepos the `next` package may not be visible from the repo root) before writing any code. Heed deprecation notices.

This block is written and re-added by `next dev` — verify at `node_modules/next/dist/server/lib/generate-agent-files.js`. Removing it from a diff only re-creates the uncommitted change; committing it with your work keeps the tree clean.

<!-- END:nextjs-agent-rules -->
