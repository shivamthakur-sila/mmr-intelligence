import { NextRequest, NextResponse } from "next/server";
import { resolveToken, TokenError } from "@/lib/resolve-token";
import { loadChecklist } from "@/lib/checklist";
import { curateDeck } from "@/lib/curate";
import { generateDeck } from "@/lib/generate-deck";
import { sectionSlides } from "@/lib/deck-layout";
import { loadSession, addUserImages, materializeImages, saveDeck } from "@/lib/session";
import { extractXlsx } from "@/lib/parsing/xlsx";
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

    // Gap attachments come from storage too, same reasoning as the
    // review step: a photo set can easily exceed Vercel's body limit.
    const storage = createServiceClient();
    const extraText: string[] = [];
    const userPhotos: { name: string; buffer: Buffer }[] = [];

    for (const f of body.attachments ?? []) {
      const { data, error } = await storage.storage.from(BUCKET).download(f.path);
      if (error || !data) continue; // a missing attachment shouldn't sink the whole build
      const buffer = Buffer.from(await data.arrayBuffer());
      if (/\.xlsx$/i.test(f.name)) {
        extraText.push(`=== Provided to fill a gap: ${f.name} ===\n${extractXlsx(buffer).summary}`);
      } else if (/\.(jpe?g|png)$/i.test(f.name)) {
        userPhotos.push({ name: f.name, buffer });
      }
    }
    if (userPhotos.length > 0) await addUserImages(sessionId, userPhotos);

    const session = await loadSession(sessionId);
    // Pull the session's photos down into this invocation's /tmp so the
    // deck builder can embed them from real file paths.
    const sessionDir = await materializeImages(sessionId, session.images);
    const checklist = await loadChecklist(resolved.siteId);

    const combined = [session.combinedText, ...extraText].join("\n\n");
    const manifest =
      session.images.length > 0
        ? session.images.map((i) => `- ${i.id}: from "${i.slideTitle || "site team upload"}"`).join("\n")
        : "(no photography available)";

    const deck = await curateDeck(
      combined,
      manifest,
      checklist,
      resolved.siteName,
      resolved.reportMonth,
      typedAnswers
    );

    const buf = await generateDeck(deck, sessionDir, session.images, resolved.siteName, resolved.reportMonth);

    // Store the built deck in Storage - not on disk - so the download and
    // commit steps, which are separate invocations, can actually reach it.
    await saveDeck(sessionId, buf);

    return NextResponse.json({
      deck,
      // Counted the way the deck is actually built, not one per section:
      // a section whose blocks overrun the content area is continued onto
      // further slides, so sections and slides are no longer the same
      // number. The cover and the at-a-glance slide are the +2.
      slideCount: 2 + deck.sections.reduce((n, s) => n + sectionSlides(s).length, 0),
      sizeKb: Math.round(buf.length / 1024),
    });
  } catch (err) {
    if (err instanceof TokenError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
