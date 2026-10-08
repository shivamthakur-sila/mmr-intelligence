-- Reset ONE site for testing.
--
-- Clears the submission trail for a single site so the upload flow can be
-- exercised again from the top: Agent 0 sees the month as outstanding,
-- issues a fresh token, and the link opens instead of being refused as
-- already-filed.
--
-- This deliberately leaves the uploaded files in Storage. They are harmless
-- orphans, keeping them costs nothing, and deleting them would destroy the
-- only copy of a real submission that is useful to test against. Clear the
-- bucket by hand if you actually want the space back.
--
-- Scoped to one site on purpose: every statement filters on the same
-- site_id. Run it against a site whose history you are willing to lose —
-- there is no undo. Set the id once, below.
--
-- Written as a DO block so the id appears exactly once and so the four
-- deletes commit together. It runs as-is in the Supabase SQL editor and in
-- psql; a psql \set would work in only one of them.
--
-- site_id is text, not a uuid — it is a readable slug ('exora',
-- 'one-trade-tower'), and it is text in all four of these tables. Checked
-- against the live schema.

DO $$
DECLARE
  target_site_id text := 'REPLACE-WITH-SITE-ID';
BEGIN
  -- Filed decks for this site.
  DELETE FROM document_registry WHERE site_id = target_site_id;

  -- Per-month state. Removing the 'received' row is what makes Agent 0
  -- treat the month as outstanding again.
  DELETE FROM mmr_status WHERE site_id = target_site_id;

  -- Chase history. Left in place, Agent 0 suppresses the reminder it has
  -- already sent and no new link is ever issued. (Written by n8n rather
  -- than by this repo; its site_id column is confirmed against the live
  -- schema.)
  DELETE FROM reminder_log WHERE site_id = target_site_id;

  -- Upload links, including ones already burnt.
  DELETE FROM upload_tokens WHERE site_id = target_site_id;
END $$;
