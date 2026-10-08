import { NextRequest, NextResponse } from "next/server";
import { resolveToken, TokenError } from "@/lib/resolve-token";
import { extractAllFiles } from "@/lib/extract-all";
import { loadChecklist } from "@/lib/checklist";
import { validateSiteMonth, reviewSubmission } from "@/lib/curate";
import { saveSession, newSessionId } from "@/lib/session";
import { createServiceClient } from "@/lib/supabase/service";
import { BUCKET } from "../signed-urls/route";

export const maxDuration = 300;

export async function POST(req: NextRequest, { params }: { params: Promise<{ token: string }> }) {
  try {
    const { token } = await params;
    const resolved = await resolveToken(token);

    // Files arrive as storage paths, not as an upload. The browser has
    // already sent them straight to Supabase Storage, because Vercel
    // rejects request bodies over ~4.5MB and real MMR decks are far
    // bigger. Downloading them here is a server-to-server fetch with no
    // such limit.
    const { files } = (await req.json()) as { files: { name: string; path: string }[] };
    if (!Array.isArray(files) || files.length === 0) {
      return NextResponse.json({ error: "No files were uploaded." }, { status: 400 });
    }

    const supabaseStorage = createServiceClient();
    const buffers: { buffer: Buffer; name: string }[] = [];
    for (const f of files) {
      const { data, error } = await supabaseStorage.storage.from(BUCKET).download(f.path);
      if (error || !data) {
        return NextResponse.json(
          { error: `Couldn't read back "${f.name}" from storage: ${error?.message ?? "not found"}` },
          { status: 500 }
        );
      }
      buffers.push({ buffer: Buffer.from(await data.arrayBuffer()), name: f.name });
    }

    const { combinedText, images, imageManifest } = await extractAllFiles(buffers);
    if (!combinedText.trim()) {
      return NextResponse.json(
        { error: "Couldn't read any content from those files. Supported formats are .pptx and .xlsx." },
        { status: 400 }
      );
    }

    // Strict check first — no point reviewing a checklist against a
    // submission that belongs to a different site or month.
    const validation = await validateSiteMonth(combinedText, resolved.siteName, resolved.reportMonth);
    if (validation.verdict === "mismatch") {
      return NextResponse.json(
        {
          rejected: true,
          reason: `This submission doesn't appear to be for ${resolved.siteName}, ${resolved.reportMonth}.`,
          evidence: validation.evidence,
          conflictType: validation.conflictType ?? null,
        },
        { status: 422 }
      );
    }

    const checklist = await loadChecklist(resolved.siteId);
    const reviewed = await reviewSubmission(combinedText, imageManifest, checklist);

    const sessionId = newSessionId();
    await saveSession(
      sessionId,
      { combinedText, imageManifest, siteName: resolved.siteName, reportMonth: resolved.reportMonth },
      images
    );

    // Merge Claude's findings with the checklist's own labels and ask
    // types, so the interface knows which control to render per gap.
    const sections = checklist.map((item) => {
      const found = reviewed.find((r) => r.key === item.key);
      return {
        key: item.key,
        label: item.label,
        askType: item.askType,
        present: found?.present ?? false,
        request: found?.request,
      };
    });

    return NextResponse.json({
      sessionId,
      siteName: resolved.siteName,
      reportMonth: resolved.reportMonth,
      validation,
      sections,
      photoCount: images.length,
      fileCount: files.length,
    });
  } catch (err) {
    if (err instanceof TokenError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
