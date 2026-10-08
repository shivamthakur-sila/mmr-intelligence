import { NextRequest, NextResponse } from "next/server";
import { resolveToken, TokenError } from "@/lib/resolve-token";
import { loadDeck } from "@/lib/session";

export async function GET(req: NextRequest, { params }: { params: Promise<{ token: string }> }) {
  try {
    const { token } = await params;
    const resolved = await resolveToken(token);

    const sessionId = new URL(req.url).searchParams.get("session");
    if (!sessionId) return NextResponse.json({ error: "Missing session id." }, { status: 400 });

    const buf = await loadDeck(sessionId, resolved);
    const name = `MMR_${resolved.siteName.replace(/\s+/g, "_")}_${resolved.reportMonth}.pptx`;

    return new NextResponse(new Uint8Array(buf), {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${name}"`,
      },
    });
  } catch (err) {
    if (err instanceof TokenError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
