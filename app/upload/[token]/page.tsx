import { resolveToken, TokenError } from "@/lib/resolve-token";
import Validator from "./validator";

function Notice({ title, body }: { title: string; body?: string }) {
  return (
    <main className="min-h-screen">
      <div style={{ height: 4, background: "var(--blue)" }} />
      <div className="mx-auto flex min-h-[70vh] w-full max-w-lg items-center px-6">
        <div>
          <h1 className="text-[26px] leading-tight">{title}</h1>
          {body && (
            <p className="mt-3 text-[15px] leading-relaxed" style={{ color: "var(--muted)" }}>
              {body}
            </p>
          )}
          <p className="mt-6 text-[13.5px]" style={{ color: "var(--muted)" }}>
            If that doesn&apos;t look right, email ai@silagroup.co.in and someone will sort it out.
          </p>
        </div>
      </div>
    </main>
  );
}

export default async function UploadTokenPage({ params }: { params: Promise<{ token: string }> }) {
  const { token } = await params;

  try {
    const resolved = await resolveToken(token);
    return (
      <Validator
        token={token}
        siteName={resolved.siteName}
        reportMonth={resolved.reportMonth}
      />
    );
  } catch (err) {
    if (err instanceof TokenError) {
      return <Notice title={err.message} />;
    }
    return (
      <Notice
        title="Something went wrong opening this link."
        body={err instanceof Error ? err.message : String(err)}
      />
    );
  }
}
