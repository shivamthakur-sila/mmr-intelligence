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
    const buf = await loadDeck(sessionId);
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
    // Any earlier row for this site and month stops being the latest before
    // the new one claims it. Committing twice is reachable: if the mmr_status
    // write below fails, this route returns 500 without burning the token, so
    // the submitter can press Confirm again — and that used to leave two
    // registry rows for one month both flagged is_latest, with nothing to say
    // which one anything downstream should read. The table carries `version`
    // and `is_latest`, so superseding rather than replacing is what it is
    // shaped for; a retry costs one redundant row, never an ambiguous one.
    //
    // Deliberately a plain update rather than an upsert keyed on source_id:
    // that would be tidier, but it needs a UNIQUE constraint on source_id
    // that nothing here has verified against the live schema, and a wrong
    // guess fails the final step of the submission.
    const { error: supersedeError } = await supabase
      .from("document_registry")
      .update({ is_latest: false })
      .eq("site_id", resolved.siteId)
      .eq("report_month", resolved.reportMonth)
      .eq("is_latest", true);
    if (supersedeError) {
      return NextResponse.json(
        { error: `Couldn't supersede the previous deck record: ${supersedeError.message}` },
        { status: 500 }
      );
    }

    const { error: regError } = await supabase.from("document_registry").insert(
      {
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
      }
    );
    if (regError) {
      return NextResponse.json({ error: `Couldn't register the deck: ${regError.message}` }, { status: 500 });
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
