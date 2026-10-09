import fs from "fs";
import path from "path";
import { NextRequest, NextResponse } from "next/server";
import { assertOwnUpload, resolveToken, TokenError } from "@/lib/resolve-token";
import { loadChecklist } from "@/lib/checklist";
import { curateDeck, sanitisePhotos } from "@/lib/curate";
import { coverPreviewDataUrl, generateDeck } from "@/lib/generate-deck";
import { deckOutline } from "@/lib/deck-layout";
import { openSession, addUserImages, materializeImages, saveDeck } from "@/lib/session";
import { extractXlsx, workbookText } from "@/lib/parsing/xlsx";
import { createServiceClient } from "@/lib/supabase/service";
import { BUCKET } from "../signed-urls/route";

export const maxDuration = 300;

export async function POST(req: NextRequest, { params }: { params: Promise<{ token: string }> }) {
  try {
    const { token } = await params;
    const resolved = await resolveToken(token);

    const body = (await req.json()) as {
      sessionId: string;
      answers: { label: string; answer: string }[];
      attachments: { name: string; path: string }[];
    };
    const sessionId = body.sessionId;
    if (!sessionId) return NextResponse.json({ error: "Missing session id." }, { status: 400 });

    const typedAnswers = body.answers ?? [];

    // Before anything is read or written: the session must belong to this
    // link. The id comes from the request body, and every query here runs
    // with the service-role client.
    await openSession(sessionId, resolved);

    // Gap attachments come from storage too, same reasoning as the
    // review step: a photo set can easily exceed Vercel's body limit.
    const storage = createServiceClient();
    const extraText: string[] = [];
    const userPhotos: { name: string; buffer: Buffer }[] = [];

    for (const f of body.attachments ?? []) {
      const own = assertOwnUpload(f.path, resolved);
      const { data, error } = await storage.storage.from(BUCKET).download(own);
      if (error || !data) continue; // a missing attachment shouldn't sink the whole build
      const buffer = Buffer.from(await data.arrayBuffer());
      if (/\.xlsx$/i.test(f.name)) {
        extraText.push(`=== Provided to fill a gap: ${f.name} ===\n${workbookText(extractXlsx(buffer).sheets)}`);
      } else if (/\.(jpe?g|png)$/i.test(f.name)) {
        userPhotos.push({ name: f.name, buffer });
      }
    }
    if (userPhotos.length > 0) await addUserImages(sessionId, userPhotos);

    // Re-read: addUserImages may just have added the site team's photos.
    const session = await openSession(sessionId, resolved);
    // Pull the session's photos down into this invocation's /tmp so the
    // deck builder can embed them from real file paths.
    const sessionDir = await materializeImages(sessionId, session.images);
    // A photo whose download failed is not offered to the model at all.
    // Otherwise the preview drew its tile while the deck silently left it
    // out, and the submitter approved a grid the filed deck did not have.
    const images = session.images.filter((i) => fs.existsSync(path.join(sessionDir, i.file)));
    if (images.length < session.images.length) {
      console.warn(`[generate] ${session.images.length - images.length} photo(s) could not be downloaded and were left out`);
    }
    const checklist = await loadChecklist(resolved.siteId);

    const combined = [session.combinedText, ...extraText].join("\n\n");
    const manifest =
      images.length > 0
        ? images.map((i) => `- ${i.id}: from "${i.slideTitle || "site team upload"}"`).join("\n")
        : "(no photography available)";

    const curated = await curateDeck(
      combined,
      manifest,
      checklist,
      resolved.siteName,
      resolved.reportMonth,
      typedAnswers
    );
    // Only photographs that exist reach the renderer and the preview alike.
    const deck = sanitisePhotos(curated, images.map((i) => i.id));

    const buf = await generateDeck(deck, sessionDir, images, resolved.siteName, resolved.reportMonth);

    // Store the built deck in Storage - not on disk - so the download and
    // commit steps, which are separate invocations, can actually reach it.
    await saveDeck(sessionId, buf);

    return NextResponse.json({
      deck,
      // Counted from the same outline the deck is built to and the preview
      // draws, so the number the submitter sees is the number filed. Adding
      // up sections - and then sections plus two - each went stale as the
      // deck gained continuation, contents and closing slides.
      slideCount: deckOutline(deck.sections).length,
      sizeKb: Math.round(buf.length / 1024),
      coverPreview: await coverPreviewDataUrl(deck, images, sessionDir),
    });
  } catch (err) {
    if (err instanceof TokenError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
