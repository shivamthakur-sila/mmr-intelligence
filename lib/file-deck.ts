import crypto from "crypto";

/** The part of a Supabase client this needs; a PostgrestClient fits too. */
type Db = {
  rpc: (fn: string, args: Record<string, unknown>) => PromiseLike<{ data: unknown; error: { code?: string; message: string } | null }>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  from: (table: string) => any;
};

export type FiledDeck = { ok: true; warning?: string } | { ok: false; error: string };

/**
 * Registers a built deck as the latest final deck for its site and month.
 *
 * Exactly one final deck per site and month may be the latest. With
 * docs/migrations/001-one-latest-final-deck.sql applied, the database
 * enforces that: file_final_deck() retires the previous deck and inserts the
 * new one in one transaction, serialised per site and month, and a partial
 * unique index refuses a second latest deck. Thirty concurrent filings were
 * run against it and left exactly one latest row, the newest.
 *
 * Until the migration is applied, the function does not exist and PostgREST
 * answers PGRST202; only then does this fall back to two REST calls - insert
 * the new row, then retire rows with a lower id - which cannot be made atomic
 * from here. Any other error from the function is a real failure and is
 * returned as one: once the index exists, the fallback's insert would be
 * refused anyway.
 */
export async function fileFinalDeck(
  db: Db,
  deck: { siteId: string; reportMonth: string; submissionId: string; storagePath: string | null }
): Promise<FiledDeck> {
  const sourceId = crypto.randomUUID();
  const fileName = `MMR_${deck.reportMonth}.pptx`;

  const { error: rpcError } = await db.rpc("file_final_deck", {
    p_site_id: deck.siteId,
    p_report_month: deck.reportMonth,
    p_source_id: sourceId,
    p_submission_id: deck.submissionId,
    p_source_filename: fileName,
    p_storage_path: deck.storagePath,
  });
  if (!rpcError) return { ok: true };
  if (rpcError.code !== "PGRST202") {
    return { ok: false, error: `Couldn't register the deck: ${rpcError.message}` };
  }

  // Column names verified against the live schema: it's source_filename
  // (not file_name), and source_id is NOT NULL and UNIQUE with no default.
  const { data: inserted, error: regError } = await db
    .from("document_registry")
    .insert({
      source_id: sourceId,
      submission_id: deck.submissionId,
      site_id: deck.siteId,
      report_month: deck.reportMonth,
      source_filename: fileName,
      file_type: "pptx",
      format_role: "final_deck",
      review_status: "approved",
      ingestion_status: "success",
      is_latest: true,
      storage_path: deck.storagePath,
    })
    .select("id")
    .single();
  if (regError || !inserted) {
    return { ok: false, error: `Couldn't register the deck: ${regError?.message ?? "no row returned"}` };
  }

  const { error: supersedeError } = await db
    .from("document_registry")
    .update({ is_latest: false })
    .eq("site_id", deck.siteId)
    .eq("report_month", deck.reportMonth)
    // Final decks only. The registry also holds the month's source files,
    // each with its own is_latest, and a new deck does not supersede them.
    .eq("format_role", "final_deck")
    .eq("is_latest", true)
    .lt("id", inserted.id);
  if (supersedeError) {
    // The new deck is registered and correct; an older one is still flagged
    // latest beside it. Reported rather than failed - the month is filed.
    return { ok: true, warning: `The previous deck record could not be superseded: ${supersedeError.message}` };
  }
  return { ok: true };
}
