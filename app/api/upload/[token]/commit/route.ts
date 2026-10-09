import { NextRequest, NextResponse } from "next/server";
import crypto from "crypto";
import { resolveToken, TokenError } from "@/lib/resolve-token";
import { createServiceClient } from "@/lib/supabase/service";
import { loadDeck } from "@/lib/session";
import { fileFinalDeck } from "@/lib/file-deck";

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

    // One latest final deck per site and month; see lib/file-deck.ts.
    // Committing twice is reachable: if the mmr_status write below fails,
    // this route returns 500 without burning the token, so the submitter can
    // press Confirm again, and the new deck supersedes the earlier one.
    const filed = await fileFinalDeck(supabase, {
      siteId: resolved.siteId,
      reportMonth: resolved.reportMonth,
      submissionId,
      storagePath: storageWarning ? null : storagePath,
    });
    if (!filed.ok) return NextResponse.json({ error: filed.error }, { status: 500 });
    if (filed.warning) {
      console.error("[commit] filed, but", { siteId: resolved.siteId, reportMonth: resolved.reportMonth, warning: filed.warning });
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
