# MMR Intelligence — Project Blueprint

## Status

This is a **working system being extended**, not a greenfield build. Agent
1 runs end to end in production today: a real 8.7MB submission has been
chased, uploaded, validated, curated, previewed, filed, and confirmed to
have stopped the chaser. Agent 2 exists as a shell only.

Treat existing behaviour as load-bearing. Several parts of it look odd and
are deliberate; the reasons are recorded in `CLAUDE.md` because they were
each learned by something breaking in production.

## The approach, and why

**Deterministic rendering, model-driven content.**

Claude decides everything about *what* a deck says: which parameters have
real content, what each table contains, how rows group, which photographs
belong to which section, which photo becomes the cover, and the opening
summary. Code decides everything about *how* it looks: geometry, palette,
typography, image handling.

The alternative — having Claude emit slide-layout code at runtime — was
considered and rejected for now, for two concrete reasons:

1. **No visual feedback loop on the current host.** Rendering slides to
   images to check them requires LibreOffice, which will not run on Vercel
   serverless. Model-written layout code with no way to see its own output
   is how distorted images reach a client.
2. **Consistency is a product requirement.** This is a recurring monthly
   report across ~1,000 sites. Layout varying per run undermines the
   standardisation the system exists to create.

This decision is worth revisiting if the deck-generation path moves to a
container (see Risk 1). At that point the render-and-critique loop becomes
possible and the tradeoff genuinely changes.

**The nearer-term win is widening what Claude composes with.** The real
limitation today is six block types. A real client MMR examined during
design contained 23 native charts; the system currently has none.

## Architecture

```
Browser ──► Supabase Storage        (files go direct; never via the API)
   │
   ▼
Next.js API routes (Vercel)
   │   review    extract → validate site/month → check against checklist
   │   generate  curate → build .pptx → store
   │   commit    register → mark received → burn token
   ▼
Supabase Postgres  ◄──  n8n (Agent 0 + confirmations; separate repo)
   │
   ▼
Anthropic API      validation, review, curation
```

### Key components

| Area | Location | Responsibility |
|---|---|---|
| Token resolution + registry gate | `lib/resolve-token.ts` | Resolves a token to site and month; refuses used, expired, or already-filed |
| Extraction | `lib/parsing/*`, `lib/extract-all.ts` | pptx text, tables, grouped shapes, embedded OLE workbooks, photographs; xlsx sheets; de-duplication |
| Checklist resolution | `lib/checklist.ts` | Site → client → default, filtered to enabled |
| Model calls | `lib/curate.ts` | Three passes: validate site/month, review against checklist, curate deck |
| Deck generation | `lib/generate-deck.ts` | Renders the curated spec into a branded `.pptx` |
| Block measurement | `lib/deck-layout.ts` | Block heights and where a section breaks onto a continuation slide; shared by the renderer and the preview so both agree |
| Session state | `lib/session.ts` | Extracted text, photos and built deck, in object storage |
| Browser upload | `lib/upload-via-storage.ts` | Signed-URL direct upload |

### Data model

14 tables. The ones that matter most:

- `sites`, `subgroups`, `groups` — the account hierarchy
- `site_contacts` — who to chase, by role
- `site_config` — the parameter catalogue and brand settings, resolved by
  `resolve_site_config(p_site_id, p_kind)`
- `mmr_status` — per site and month; `state='received'` is the signal that
  stops Agent 0
- `upload_tokens`, `document_registry`, `reminder_log`
- `narrative_chunks` — pgvector, for Agent 2; unused so far

## Risks, most uncertain first

**Risk 1 — Generation may exceed the serverless time limit on the largest
decks.** The largest real submission seen is 26MB with 41 photographs.
Generation on an 8.7MB deck already takes 80–100 seconds against a 300
second ceiling, and cost scales with photo count because every image is
cropped individually. This has not been tested at 26MB.

*Trigger and response:* the first timeout on a real site means deck
generation moves to a container worker (Railway, Fly or Render) with the
app staying on Vercel. Do not pre-emptively migrate — two of the three
Vercel workarounds already built (direct-to-storage upload, stateless
sessions) are correct architecture regardless of host and should be kept.

*Cheap mitigations to do first:* crop images in parallel rather than
serially, cap photographs per deck, and skip cropping when a source image
already matches the target ratio closely.

**Risk 2 — Charts (retired).** Closed by Slice 1: a `chart` block type runs
end to end — prompt contract, type, renderer, preview. Proven on a real
submission, not a fixture: the live three-pass model path produced five
charts across bar, horizontal bar, line and pie from the Ahuja Tower August
2026 deck, all 34 charted values traceable to the extracted source, each
slide rendered through LibreOffice and inspected.

Four behaviours found by rendering are recorded in `CLAUDE.md`: a palette
shorter than its data gets random colours, a doughnut ignores an outside
label position, a line chart's axis has to be set explicitly, and an
overflowing section must spill rather than be truncated. None would have
been caught by type-checking, and the first curation run produced no charts
at all until the prompt was rebalanced — which only reading the output
revealed.

**Risk 3 — Agent 2 is blocked on an unmade decision.** The ring-fenced
retrieval query has been proven correct against synthetic data. It cannot
proceed without an embedding provider chosen and a key issued, and the
`narrative_chunks.embedding` column must be sized to that provider's output
dimension before any content is indexed. Changing it later means re-indexing
everything.

**Risk 4 — Checklist coverage is inferred, not authoritative.** The 22
default-enabled parameters were derived from frequency across eight real
MMRs. Nobody at SILA has confirmed a mandatory set. A large client whose
requirements differ will need per-client configuration, which is supported
but unpopulated.

**Risk 5 — Model output shape is a runtime contract.** Curation returns JSON
that the renderer consumes structurally. A malformed or truncated response
fails the whole build. Token limits have already caused this once.

## Build order

Thin vertical slices. Each ends with something demonstrable against a real
file, not a unit test in isolation.

**Slice 1 — Charts. Done.** A `chart` block type end to end: prompt
contract, type, renderer, preview. Bar, line and pie, verified by rendering
to image and inspecting. Closes the largest gap against a human-made deck.

**Slice 2 — Generation robustness at 26MB.** Run the largest real
submission through generation. Parallelise cropping, cap photo counts, and
measure. Resolves Risk 1 with evidence rather than assumption.

**Slice 3 — Agent 2 retrieval.** Choose the embedding provider, size the
vector column, index filed decks, and wire the existing ask surface to
real ring-fenced retrieval with citations.

**Slice 4 — Agent 2 access control.** Real RLS policies against
`user_site_access`, so scope is enforced by the database rather than by
application code.

**Slice 5 — Richer composition.** Timelines, two-column splits,
before/after pairs, KPI callouts — the layouts real client decks use that
the six current block types cannot express.

**Slice 6 — Operability.** Admin visibility into submissions, failures and
per-site configuration. Currently all of this is SQL.

## Deliberately deferred

- Moving off Vercel (see Risk 1 for the trigger).
- Claude generating layout code at runtime (revisit after a container move).
- A UI for editing the parameter catalogue.
