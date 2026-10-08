import fs from "fs";
import path from "path";
import os from "os";
import crypto from "crypto";
import { createServiceClient } from "./supabase/service";
import { TokenError } from "./resolve-token";
import type { ExtractedImage } from "./parsing/images";

export const SESSION_BUCKET = "mmr-uploads";

/**
 * Session state lives in Supabase Storage, NOT on disk.
 *
 * Each API route runs as a separate serverless invocation with its own
 * ephemeral /tmp, so anything the review step writes to disk is gone by
 * the time the generate step runs. Storage is the only thing both
 * invocations can actually see. Images get pulled down into /tmp on
 * demand, inside whichever invocation needs them, because pptxgenjs
 * embeds images from file paths rather than buffers.
 */

export type ImageRecord = { id: string; file: string; slideNumber: number; slideTitle: string };

export type SessionData = {
  /** The site this session belongs to. Every later step checks it against the token. */
  siteId: string;
  combinedText: string;
  imageManifest: string;
  images: ImageRecord[];
  siteName: string;
  reportMonth: string;
};

export function newSessionId(): string {
  return crypto.randomBytes(8).toString("hex");
}

/** Exactly what newSessionId issues. Anything else could steer keyFor elsewhere in the bucket. */
const SESSION_ID = /^[0-9a-f]{16}$/;

function keyFor(id: string, name: string) {
  if (!SESSION_ID.test(id)) throw new TokenError("That session id isn't valid.", 400);
  return `sessions/${id}/${name}`;
}

function contentTypeFor(ext: string) {
  return ext === ".png" ? "image/png" : "image/jpeg";
}

export async function saveSession(
  id: string,
  data: Omit<SessionData, "images">,
  images: ExtractedImage[]
): Promise<void> {
  const supabase = createServiceClient();

  const records: ImageRecord[] = [];
  for (const img of images) {
    const fileName = `${img.id}${img.ext}`;
    const { error } = await supabase.storage
      .from(SESSION_BUCKET)
      .upload(keyFor(id, fileName), img.buffer, {
        contentType: contentTypeFor(img.ext),
        upsert: true,
      });
    if (error) throw new Error(`Couldn't store an extracted image: ${error.message}`);
    records.push({
      id: img.id,
      file: fileName,
      slideNumber: img.slideNumber,
      slideTitle: img.slideTitle,
    });
  }

  const payload: SessionData = { ...data, images: records };
  const { error } = await supabase.storage
    .from(SESSION_BUCKET)
    .upload(keyFor(id, "session.json"), Buffer.from(JSON.stringify(payload)), {
      contentType: "application/json",
      upsert: true,
    });
  if (error) throw new Error(`Couldn't store the session: ${error.message}`);
}

export async function loadSession(id: string): Promise<SessionData> {
  const supabase = createServiceClient();
  const { data, error } = await supabase.storage
    .from(SESSION_BUCKET)
    .download(keyFor(id, "session.json"));
  if (error || !data) {
    throw new Error("That session has expired or wasn't found. Start again from the upload step.");
  }
  return JSON.parse(await data.text()) as SessionData;
}

/**
 * Loads a session and refuses it unless it belongs to the caller's token.
 *
 * The upload routes run with the service-role client and so trust nothing the
 * browser sends. A session id came straight from the request body and was
 * used as given: anyone holding one valid link could name another site's
 * session and have its photographs built into their deck, or file that
 * site's built deck as their own month. Every route that touches a session
 * now goes through here first.
 */
export async function openSession(
  id: string,
  scope: { siteId: string; reportMonth: string }
): Promise<SessionData> {
  const session = await loadSession(id);
  if (session.siteId !== scope.siteId || session.reportMonth !== scope.reportMonth) {
    throw new TokenError("That session doesn't belong to this link.", 403);
  }
  return session;
}

/**
 * Downloads this session's images into a local temp directory and returns
 * its path, so the deck builder in this same invocation can read them.
 */
export async function materializeImages(id: string, images: ImageRecord[]): Promise<string> {
  const supabase = createServiceClient();
  const dir = path.join(os.tmpdir(), `sila-session-${id}`);
  fs.mkdirSync(dir, { recursive: true });

  for (const img of images) {
    const target = path.join(dir, img.file);
    if (fs.existsSync(target)) continue;
    const { data, error } = await supabase.storage
      .from(SESSION_BUCKET)
      .download(keyFor(id, img.file));
    if (error || !data) continue; // one missing photo shouldn't sink the deck
    fs.writeFileSync(target, Buffer.from(await data.arrayBuffer()));
  }

  return dir;
}

/** Adds photos the site team attached while filling a gap. */
export async function addUserImages(
  id: string,
  files: { name: string; buffer: Buffer }[]
): Promise<ImageRecord[]> {
  const supabase = createServiceClient();
  const session = await loadSession(id);
  let counter = session.images.length;
  const added: ImageRecord[] = [];

  for (const f of files) {
    const rawExt = (f.name.match(/\.(jpe?g|png)$/i)?.[0] ?? ".jpg").toLowerCase();
    const ext = rawExt === ".jpeg" ? ".jpg" : rawExt;
    const imgId = `user${++counter}`;
    const fileName = `${imgId}${ext}`;
    const { error } = await supabase.storage
      .from(SESSION_BUCKET)
      .upload(keyFor(id, fileName), f.buffer, {
        contentType: contentTypeFor(ext),
        upsert: true,
      });
    if (error) continue;
    added.push({
      id: imgId,
      file: fileName,
      slideNumber: 0,
      slideTitle: "Provided by the site team",
    });
  }

  session.images.push(...added);
  await supabase.storage
    .from(SESSION_BUCKET)
    .upload(keyFor(id, "session.json"), Buffer.from(JSON.stringify(session)), {
      contentType: "application/json",
      upsert: true,
    });

  return added;
}

/** Stores the built deck so later invocations can serve or file it. */
export async function saveDeck(id: string, deck: Buffer): Promise<void> {
  const supabase = createServiceClient();
  const { error } = await supabase.storage
    .from(SESSION_BUCKET)
    .upload(keyFor(id, "deck.pptx"), deck, {
      contentType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      upsert: true,
    });
  if (error) throw new Error(`Couldn't store the generated deck: ${error.message}`);
}

/**
 * The built deck, for the session this token owns. Takes the scope rather than
 * leaving the check to each caller: commit files whatever this returns as the
 * token's site and month, so a missed check there would file another site's
 * deck.
 */
export async function loadDeck(
  id: string,
  scope: { siteId: string; reportMonth: string }
): Promise<Buffer> {
  await openSession(id, scope);
  const supabase = createServiceClient();
  const { data, error } = await supabase.storage
    .from(SESSION_BUCKET)
    .download(keyFor(id, "deck.pptx"));
  if (error || !data) {
    throw new Error("That generated deck is no longer available. Build it again.");
  }
  return Buffer.from(await data.arrayBuffer());
}
