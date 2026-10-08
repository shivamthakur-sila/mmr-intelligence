import { NextRequest, NextResponse } from "next/server";
import { extractPptx } from "@/lib/parsing/pptx";
import { extractXlsx } from "@/lib/parsing/xlsx";

async function extractByName(buffer: Buffer, fileName: string) {
  const lower = fileName.toLowerCase();

  if (lower.endsWith(".pptx")) {
    const { summary } = await extractPptx(buffer);
    return { extractedText: summary, mode: "check" as const };
  }

  if (lower.endsWith(".xlsx")) {
    const { summary } = extractXlsx(buffer);
    return { extractedText: summary, mode: "build" as const };
  }

  if (lower.endsWith(".pdf") || /\.(png|jpe?g|webp)$/.test(lower)) {
    return {
      extractedText: null,
      passThrough: true,
      note: "PDF/image - pass directly to Claude as a document/image content block instead of extracted text.",
    };
  }

  throw Object.assign(new Error(`Unsupported file type: ${fileName}`), { status: 415 });
}

export async function POST(req: NextRequest) {
  const contentType = req.headers.get("content-type") || "";

  try {
    // Direct upload - for testing via Postman (form-data, key "file"),
    // and the real production path from Agent 1's n8n workflow.
    if (contentType.includes("multipart/form-data")) {
      const formData = await req.formData();
      const file = formData.get("file");
      if (!file || !(file instanceof File)) {
        return NextResponse.json({ error: "Expected a form-data field named 'file'" }, { status: 400 });
      }
      // Echoed straight back - these need to survive in THIS branch's
      // output regardless of whether the sender-lookup branch succeeds,
      // since the one case that most needs them (an unrecognized sender,
      // zero rows from the Postgres lookup) is exactly when that branch
      // won't have them at all.
      const passthrough = {
        trigger_source: formData.get("trigger_source") ?? "",
        gmail_message_id: formData.get("gmail_message_id") ?? "",
        user_message: formData.get("user_message") ?? "",
      };
      const buffer = Buffer.from(await file.arrayBuffer());
      const result = await extractByName(buffer, file.name);
      return NextResponse.json({ ...result, ...passthrough });
    }

    // URL - used when n8n forwards a Gmail attachment URL (or an
    // upload-link submission's storage URL) instead of raw binary.
    const body = await req.json();
    const { attachmentUrl, trigger_source, gmail_message_id, user_message } = body;
    const passthrough = { trigger_source: trigger_source ?? "", gmail_message_id: gmail_message_id ?? "", user_message: user_message ?? "" };
    if (!attachmentUrl) {
      return NextResponse.json(
        { error: "Send either multipart/form-data with a 'file' field, or JSON with an attachmentUrl.", ...passthrough },
        { status: 400 }
      );
    }

    const fileRes = await fetch(attachmentUrl);
    if (!fileRes.ok) {
      return NextResponse.json({ error: `Could not fetch attachment: ${fileRes.status}`, ...passthrough }, { status: 502 });
    }
    const buffer = Buffer.from(await fileRes.arrayBuffer());
    const result = await extractByName(buffer, attachmentUrl);
    return NextResponse.json({ ...result, ...passthrough });
  } catch (err) {
    const status = (err as { status?: number }).status ?? 500;
    return NextResponse.json(
      { error: err instanceof Error ? err.message : String(err) },
      { status }
    );
  }
}
