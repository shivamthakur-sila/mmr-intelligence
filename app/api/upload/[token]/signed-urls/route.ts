import { NextRequest, NextResponse } from "next/server";
import crypto from "crypto";
import { resolveToken, TokenError } from "@/lib/resolve-token";
import { createServiceClient } from "@/lib/supabase/service";

export const BUCKET = "mmr-uploads";

/**
 * Hands the browser a signed URL per file so it can upload directly to
 * Supabase Storage. This exists because Vercel rejects request bodies
 * over roughly 4.5MB before any of our code runs, and real MMR decks
 * run well past that (one site's is 26MB). Only the resulting storage
 * paths travel through Vercel afterwards, which is a few hundred bytes.
 */
export async function POST(req: NextRequest, { params }: { params: Promise<{ token: string }> }) {
  try {
    const { token } = await params;
    const resolved = await resolveToken(token);

    const { files } = (await req.json()) as { files: { name: string; size: number }[] };
    if (!Array.isArray(files) || files.length === 0) {
      return NextResponse.json({ error: "No files listed." }, { status: 400 });
    }

    const supabase = createServiceClient();

    // Create the bucket on first use rather than making this a manual
    // setup step someone has to remember.
    const { data: buckets } = await supabase.storage.listBuckets();
    if (!buckets?.some((b) => b.name === BUCKET)) {
      const { error } = await supabase.storage.createBucket(BUCKET, { public: false });
      // A concurrent request may have created it between the check and
      // here; that's fine and not worth failing over.
      if (error && !/already exists/i.test(error.message)) {
        return NextResponse.json(
          { error: `Couldn't prepare storage: ${error.message}` },
          { status: 500 }
        );
      }
    }

    const batch = crypto.randomBytes(6).toString("hex");
    const issued: { name: string; path: string; token: string }[] = [];

    for (const f of files) {
      // Keep the original extension so the extractor can tell a .pptx
      // from a .xlsx, but strip the rest of the name to avoid path
      // characters causing trouble in storage keys.
      const ext = (f.name.match(/\.(pptx|xlsx|jpe?g|png)$/i)?.[0] ?? "").toLowerCase();
      const safe = `${crypto.randomBytes(4).toString("hex")}${ext}`;
      const path = `${resolved.siteId}/${resolved.reportMonth}/${batch}/${safe}`;

      const { data, error } = await supabase.storage.from(BUCKET).createSignedUploadUrl(path);
      if (error || !data) {
        return NextResponse.json(
          { error: `Couldn't create an upload URL: ${error?.message ?? "unknown"}` },
          { status: 500 }
        );
      }
      issued.push({ name: f.name, path: data.path, token: data.token });
    }

    return NextResponse.json({ bucket: BUCKET, files: issued });
  } catch (err) {
    if (err instanceof TokenError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
