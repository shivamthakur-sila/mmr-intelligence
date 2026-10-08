import { NextRequest, NextResponse } from "next/server";
import { createClient } from "@/lib/supabase/server";

export async function GET(request: NextRequest) {
  const { searchParams, origin } = new URL(request.url);
  const code = searchParams.get("code");
  const next = searchParams.get("next") ?? "/dashboard";

  if (code) {
    const supabase = await createClient();
    const { data, error } = await supabase.auth.exchangeCodeForSession(code);

    if (!error && data.user) {
      const email = data.user.email ?? "";
      // The `hd` hint on the login page only shapes Google's account
      // picker - it is not a guarantee. This is the real check: reject
      // and sign out anyone whose verified email isn't actually on the
      // org's domain, regardless of how they got a session.
      if (!email.endsWith("@silagroup.co.in")) {
        await supabase.auth.signOut();
        return NextResponse.redirect(`${origin}/login?error=wrong_domain`);
      }
      return NextResponse.redirect(`${origin}${next}`);
    }
  }

  return NextResponse.redirect(`${origin}/login?error=auth_failed`);
}
