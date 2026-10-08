# MMR Intelligence — Requirements

## Goal

SILA Facility Management operates roughly 1,000 sites across India. Every
site must submit a Monthly Management Report (MMR) — a slide deck covering
manpower, compliance, maintenance, utilities, incidents and activity
evidence. Today this is chased by hand over email, assembled by hand in
PowerPoint, and tracked by hand in spreadsheets.

MMR Intelligence automates the loop: chase the sites that haven't
submitted, validate and assemble what they send, record it, and let the
business ask questions of everything that has been filed.

The system is three agents:

- **Agent 0 — Chaser.** Scheduled. Finds sites whose MMR is due and not
  received, and emails the person in charge a single-use upload link.
  Escalates to site head, then HOD, as the deadline passes.
- **Agent 1 — Validator & Builder.** The upload link opens this. Accepts
  whatever the site has (finished deck, raw workbooks, or both), verifies
  it belongs to that site and month, checks it against the parameters that
  site is required to report, asks for anything missing, then assembles a
  branded deck and records it as filed.
- **Agent 2 — Insight.** A login-gated surface for asking questions across
  everything filed, answering only from real filed content and always
  citing its source.

## Scope of this repository

This repo is **the web application only** — Agent 1's entire flow and
Agent 2's surface.

Agent 0 and the confirmation-email workflow run in **n8n** and are outside
this repo. They are not owned here, but they share the database, so their
contract is binding and documented in `CLAUDE.md`. Changing the columns
they depend on breaks them silently.

## Functional requirements

### Agent 1 — submission

- **FR-1.** An upload link carries a single-use token. The token alone
  identifies the site and report month; no account or password is required.
  This is deliberate: site managers are occasional, non-technical users and
  login friction was judged the main adoption risk.
- **FR-2.** A submission is refused before any processing if that site and
  month is already recorded as received.
- **FR-3.** A submission may contain several files. `.pptx` and `.xlsx` are
  read for content; `.jpg`/`.png` are accepted as photographic evidence.
- **FR-4.** Extraction must recover content that naive parsing misses:
  native tables, text nested inside grouped shapes, workbooks embedded as
  OLE objects, and photographs. All four occur in real submissions.
- **FR-5.** The submission is rejected if its content names a different
  site or a different report month than the token expects. Content that
  names neither is accepted — absence is not a conflict.
- **FR-6.** Content is checked against the parameters required for that
  specific site, resolved per site, then per client, then falling back to a
  shared default set.
- **FR-7.** For each missing parameter the submitter may type an answer,
  attach a file, or explicitly skip it.
- **FR-8.** The generated deck contains only parameters with real content.
  Anything absent or skipped is omitted entirely — no placeholder slides.
- **FR-9.** The submitter previews the deck in the browser before filing,
  and may download it.
- **FR-10.** On confirmation the system stores the deck, registers it,
  marks the month received, and burns the token. Marking it received is
  what stops Agent 0 chasing that site.

### Agent 2 — insight

- **FR-11.** Access requires sign-in with a `@silagroup.co.in` Google
  account, enforced server-side.
- **FR-12.** A signed-in user may only query sites they have access to.
  Scope is resolved from the database before retrieval — never by filtering
  results afterwards.
- **FR-13.** Answers are drawn only from filed content and cite the site,
  month and document they came from. Where the filed content does not
  answer the question, that is the answer.

### Cross-cutting

- **FR-14. Never invent.** No number, name or date may appear in any output
  unless it is present in the submitted source or was supplied by the
  submitter. This is the single most important rule in the system; a
  fabricated figure in a client-facing report is the worst failure
  available to it.

## Non-functional requirements

- **NFR-1. Large files.** Real submissions reach 26MB. Files must never
  transit an API request body; the browser uploads directly to object
  storage and only references are passed on.
- **NFR-2. Stateless request handling.** Nothing may rely on local disk
  persisting between requests. Session state, extracted media and generated
  decks live in object storage.
- **NFR-3. Latency and feedback.** Review takes roughly 30 seconds and
  generation one to two minutes. Any operation over two seconds must show
  live, honest progress — elapsed time and a description of the work, never
  a fabricated percentage.
- **NFR-4. Secrets.** Service-role database credentials and the Anthropic
  API key are server-side only and must never reach the browser.
- **NFR-5. Least privilege.** The token flow deliberately runs with
  elevated database access because there is no user session; it must
  therefore validate the token itself and scope every query to the site and
  month that token resolves to.
- **NFR-6. Idempotency.** Re-running any scheduled or repeatable step must
  not duplicate emails, submissions or registry rows.
- **NFR-7. Brand fidelity.** Generated decks follow SILA's Facility
  Management brand: FM Blue `#264170`, Sunshine `#F7A328`. Geometry and
  palette are guaranteed by code, not left to model output.
- **NFR-8. Legible failure.** Errors state what actually failed. Generic
  messages have repeatedly cost hours of misdiagnosis on this project.

## Out of scope

- Agent 0 and the confirmation workflow (they live in n8n).
- Authoring or editing MMR content on behalf of a site.
- Any interface for configuring a site's parameters. The catalogue is
  seeded and changed by SQL; a management UI is a later project.
- Mobile-native applications.
- Migrating historical MMRs into the system.
- Sending email from this application. All email is n8n's.

## What "done" looks like

The project is done when all of the following hold:

1. A site manager who receives a reminder can complete a submission end to
   end — open the link, upload real files, resolve gaps, preview, file —
   without assistance, and Agent 0 stops chasing them automatically.
2. A deck generated from a real submission is good enough to send to a
   client without manual editing, including charts and photographs.
3. A submission for the wrong site or month is refused with an explanation
   naming the conflicting content.
4. A 26MB submission completes without error.
5. Every figure in a generated deck is traceable to the submitted source or
   to an answer the submitter typed.
6. A signed-in user can ask a question across filed MMRs and receives an
   answer citing its sources, covering only the sites they may see.
7. A fresh engineer can run the system locally and deploy it from the
   documentation in this repository alone.