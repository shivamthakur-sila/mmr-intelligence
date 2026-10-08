import { createClient } from "./supabase/client";

export type UploadedRef = { name: string; path: string };

/**
 * Uploads files straight from the browser to Supabase Storage using
 * short-lived signed URLs issued by our server. Nothing large passes
 * through Vercel, which rejects request bodies over roughly 4.5MB —
 * real MMR decks run to 26MB, so this isn't an edge case.
 */
export async function uploadViaStorage(
  token: string,
  files: File[],
  onProgress?: (done: number, total: number) => void
): Promise<UploadedRef[]> {
  const res = await fetch(`/api/upload/${token}/signed-urls`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ files: files.map((f) => ({ name: f.name, size: f.size })) }),
  });

  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data.error || "Couldn't prepare the upload.");
  }

  const { bucket, files: issued } = (await res.json()) as {
    bucket: string;
    files: { name: string; path: string; token: string }[];
  };

  const supabase = createClient();
  const refs: UploadedRef[] = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const slot = issued[i];

    const { error } = await supabase.storage
      .from(bucket)
      .uploadToSignedUrl(slot.path, slot.token, file);

    if (error) {
      throw new Error(`Couldn't upload "${file.name}": ${error.message}`);
    }

    refs.push({ name: file.name, path: slot.path });
    onProgress?.(i + 1, files.length);
  }

  return refs;
}
