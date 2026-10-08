"use client";

import { Suspense, useState } from "react";
import { useSearchParams } from "next/navigation";
import { createClient } from "@/lib/supabase/client";

function LoginInner() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const searchParams = useSearchParams();
  // Preserves where the person was headed (e.g. /upload/[token]) across
  // the login round-trip, so they land back where they clicked from.
  const next = searchParams.get("next") ?? "/dashboard";

  async function handleSignIn() {
    setLoading(true);
    setError(null);
    const supabase = createClient();
    const { error } = await supabase.auth.signInWithOAuth({
      provider: "google",
      options: {
        redirectTo: `${window.location.origin}/auth/callback?next=${encodeURIComponent(next)}`,
        queryParams: {
          // Hints Google's account picker to the org's domain. This is a
          // UI hint only, not a hard guarantee - the callback route does
          // the real enforcement server-side after login completes.
          hd: "silagroup.co.in",
        },
      },
    });
    if (error) {
      setError(error.message);
      setLoading(false);
    }
  }

  return (
    <main className="flex min-h-screen items-center justify-center bg-[#F5F5F4] px-4">
      <div className="w-full max-w-sm rounded-lg border border-[#E5E5E4] bg-white p-8 shadow-sm">
        <h1 className="text-xl font-semibold text-[#264170]">MMR Intelligence</h1>
        <p className="mt-2 text-sm text-[#697784]">
          Sign in with your SILA Google account to view pending MMRs or ask a question.
        </p>

        {error && (
          <p className="mt-4 rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">
            {error}
          </p>
        )}

        <button
          onClick={handleSignIn}
          disabled={loading}
          className="mt-6 flex w-full items-center justify-center gap-2 rounded-md bg-[#264170] px-4 py-2.5 text-sm font-medium text-white transition hover:bg-[#1c3157] disabled:opacity-60"
        >
          {loading ? "Redirecting..." : "Sign in with Google"}
        </button>
      </div>
    </main>
  );
}

export default function LoginPage() {
  return (
    <Suspense>
      <LoginInner />
    </Suspense>
  );
}
