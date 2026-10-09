import { NextRequest, NextResponse } from "next/server";
import crypto from "crypto";
import { resolveToken, TokenError } from "@/lib/resolve-token";
import { createServiceClient } from "@/lib/supabase/service";
import { loadDeck } from "@/lib/session";

export const maxDuration = 120;

export async function POST(req: NextRequest, { params }: { params: Promise<{ token: string }> }) {
  try {
    const { token } = await params;
    const resolved = await resolveToken(token);

    const { sessionId } = await req.json();
    if (!sessionId) return NextResponse.json({ error: "Missing session id." }, { status: 400 });

    const supabase = createServiceClient();
    const buf = await loadDeck(sessionId, resolved);
    const submissionId = crypto.randomUUID();
    const storagePath = `${resolved.siteId}/${resolved.reportMonth}/MMR_${resolved.reportMonth}.pptx`;

    // Storage is best-effort: a missing bucket shouldn't block the
    // registry from recording that the month is filed. The failure is
    // reported back rather than swallowed.
    let storageWarning: string | null = null;
    const upload = await supabase.storage
      .from("mmr-decks")
      .upload(storagePath, buf, {
        contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        upsert: true,
      });
    if (upload.error) storageWarning = upload.error.message;

    // Column names verified against the live schema: it's source_filename
    // (not file_name), and source_id is NOT NULL and UNIQUE with no
    // default, so it has to be supplied.
    //
    // Committing twice is reachable: if the mmr_status write below fails,
    // this route returns 500 without burning the token, so the submitter can
    // press Confirm again. The table carries `is_latest`, so a new deck
    // supersedes the old one rather than replacing it, and exactly one final
    // deck per site and month should be the latest.
    //
    // The new row is written first and only rows older than it are retired
    // afterwards. The other order - retire, then insert - left the month
    // with no latest deck at all whenever the insert failed, and two
    // concurrent commits each retired nothing of the other's and both stayed
    // latest. Now a failed insert changes nothing, and of two commits the
    // newer row - the higher id, which the database assigns at insert -
    // always retires the older, whichever finishes first. Making it atomic
    // outright needs a transaction, which the REST client cannot open; the
    // window left is two inserts whose ids and commits land in opposite
    // orders within the same instant.
    const { data: inserted, error: regError } = await supabase
      .from("document_registry")
      .insert({
        source_id: crypto.randomUUID(),
        submission_id: submissionId,
        site_id: resolved.siteId,
        report_month: resolved.reportMonth,
        source_filename: `MMR_${resolved.reportMonth}.pptx`,
        file_type: "pptx",
        format_role: "final_deck",
        review_status: "approved",
        ingestion_status: "success",
        is_latest: true,
        storage_path: storageWarning ? null : storagePath,
      })
      .select("id")
      .single();
    if (regError || !inserted) {
      return NextResponse.json(
        { error: `Couldn't register the deck: ${regError?.message ?? "no row returned"}` },
        { status: 500 }
      );
    }

    const { error: supersedeError } = await supabase
      .from("document_registry")
      .update({ is_latest: false })
      .eq("site_id", resolved.siteId)
      .eq("report_month", resolved.reportMonth)
      // Final decks only. The registry also holds the source files and
      // other formats for the month, each with its own is_latest, and those
      // are not superseded by a new deck.
      .eq("format_role", "final_deck")
      .eq("is_latest", true)
      .lt("id", inserted.id);
    if (supersedeError) {
      // The new deck is registered and correct; the older one is still
      // flagged latest beside it. Reported rather than failed, for the same
      // reason as the token burn below - the month is filed.
      console.error("[commit] filed, but the previous deck record could not be superseded", {
        siteId: resolved.siteId,
        reportMonth: resolved.reportMonth,
        error: supersedeError.message,
      });
    }

    // This single write is what makes Agent 0 stop chasing this site -
    // its due-check query already excludes anything marked received.
    const { error: statusError } = await supabase.from("mmr_status").upsert(
      {
        site_id: resolved.siteId,
        report_month: resolved.reportMonth,
        state: "received",
        deck_received_at: new Date().toISOString(),
        current_submission_id: submissionId,
      },
      { onConflict: "site_id,report_month" }
    );
    if (statusError) {
      return NextResponse.json({ error: `Couldn't mark it complete: ${statusError.message}` }, { status: 500 });
    }

    // Burn the token so the link can't be reused. Reported rather than
    // thrown: the month is filed by this point, and failing the request over
    // it would tell the submitter their submission didn't go through when it
    // did. The link is not left open either way — resolveToken refuses a
    // month already marked received — but a burn that silently failed used to
    // leave nothing behind to explain it.
    const { error: burnError } = await supabase
      .from("upload_tokens")
      .update({ used_at: new Date().toISOString() })
      .eq("token", token);
    if (burnError) {
      // The submitter is told too, but the browser is not a log: this is
      // the record an operator would actually look for.
      console.error("[commit] filed, but token burn failed", {
        siteId: resolved.siteId,
        reportMonth: resolved.reportMonth,
        error: burnError.message,
      });
    }

    return NextResponse.json({
      filed: true,
      siteName: resolved.siteName,
      reportMonth: resolved.reportMonth,
      storageWarning,
      tokenWarning: burnError
        ? `Filed, but the link couldn't be marked used: ${burnError.message}`
        : null,
    });
  } catch (err) {
    if (err instanceof TokenError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
