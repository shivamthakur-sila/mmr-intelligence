import { createServiceClient } from "./supabase/service";

export type ResolvedToken = {
  token: string;
  siteId: string;
  siteName: string;
  reportMonth: string;
};

export class TokenError extends Error {
  status: number;
  constructor(message: string, status = 400) {
    super(message);
    this.status = status;
  }
}

/**
 * Resolves an upload token to its site and month, and refuses anything
 * that shouldn't proceed. The registry check comes first by design: a
 * month already marked received is never reprocessed, which is what
 * "marked complete" actually has to mean to be worth anything.
 */
export async function resolveToken(token: string): Promise<ResolvedToken> {
  const supabase = createServiceClient();

  const { data: row, error } = await supabase
    .from("upload_tokens")
    .select("token, site_id, report_month, expires_at, used_at")
    .eq("token", token)
    .maybeSingle();

  if (error) throw new TokenError(`Couldn't check that link: ${error.message}`, 500);
  if (!row) throw new TokenError("This link isn't recognised.");
  if (row.used_at) throw new TokenError("This link has already been used.");
  if (new Date(row.expires_at) < new Date()) throw new TokenError("This link has expired.");

  const { data: status } = await supabase
    .from("mmr_status")
    .select("state")
    .eq("site_id", row.site_id)
    .eq("report_month", row.report_month)
    .maybeSingle();

  if (status?.state === "received") {
    throw new TokenError(
      `${row.report_month}'s MMR for this site is already filed. Nothing further is needed.`,
      409
    );
  }

  const { data: site } = await supabase
    .from("sites")
    .select("site_name")
    .eq("site_id", row.site_id)
    .maybeSingle();

  return {
    token: row.token,
    siteId: row.site_id,
    siteName: site?.site_name ?? row.site_id,
    reportMonth: row.report_month,
  };
}
