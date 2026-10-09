-- One latest final deck per site and month, enforced by the database.
--
-- Why: app/api/upload/[token]/commit/route.ts files a deck with two REST
-- calls - insert the new row, then retire older ones. The REST client cannot
-- open a transaction, so two filings landing in the same instant can both
-- stay is_latest. This migration closes that: a partial unique index makes a
-- second latest final deck impossible, and file_final_deck() retires and
-- inserts inside one transaction, serialised per site and month.
--
-- The commit route calls file_final_deck() when it exists and falls back to
-- its two-call path when it does not, so this can be applied at any time.
-- Apply it in the Supabase SQL editor. Safe to run more than once.

begin;

-- 1. If duplicates already exist, keep the newest (highest id) as latest.
--    The index below cannot be created while any remain.
update document_registry d
   set is_latest = false
 where d.format_role = 'final_deck'
   and d.is_latest
   and exists (
     select 1 from document_registry n
      where n.site_id = d.site_id
        and n.report_month = d.report_month
        and n.format_role = 'final_deck'
        and n.is_latest
        and n.id > d.id
   );

-- 2. At most one latest final deck per site and month.
create unique index if not exists document_registry_one_latest_final_deck
  on document_registry (site_id, report_month)
  where is_latest and format_role = 'final_deck';

-- 3. Retire the previous latest deck and register the new one atomically.
--    The advisory lock serialises filings for the same site and month, so a
--    second caller waits for the first to commit and then retires its row,
--    instead of failing on the index.
create or replace function file_final_deck(
  p_site_id text,
  p_report_month text,
  p_source_id text,
  p_submission_id uuid,
  p_source_filename text,
  p_storage_path text
) returns bigint
language plpgsql
security invoker
set search_path = public
as $$
declare
  new_id bigint;
begin
  perform pg_advisory_xact_lock(hashtextextended(p_site_id || '/' || p_report_month, 0));

  update document_registry
     set is_latest = false
   where site_id = p_site_id
     and report_month = p_report_month
     and format_role = 'final_deck'
     and is_latest;

  insert into document_registry (
    source_id, submission_id, site_id, report_month, source_filename,
    file_type, format_role, review_status, ingestion_status, is_latest, storage_path
  ) values (
    p_source_id, p_submission_id, p_site_id, p_report_month, p_source_filename,
    'pptx', 'final_deck', 'approved', 'success', true, p_storage_path
  )
  returning id into new_id;

  return new_id;
end;
$$;

-- Only the server (service role) files decks. The upload flow has no user
-- session, and nothing in the browser may call this.
revoke all on function file_final_deck(text, text, text, uuid, text, text) from public;
do $$
begin
  if exists (select 1 from pg_roles where rolname = 'anon') then
    revoke all on function file_final_deck(text, text, text, uuid, text, text) from anon;
  end if;
  if exists (select 1 from pg_roles where rolname = 'authenticated') then
    revoke all on function file_final_deck(text, text, text, uuid, text, text) from authenticated;
  end if;
  if exists (select 1 from pg_roles where rolname = 'service_role') then
    grant execute on function file_final_deck(text, text, text, uuid, text, text) to service_role;
  end if;
end $$;

commit;
