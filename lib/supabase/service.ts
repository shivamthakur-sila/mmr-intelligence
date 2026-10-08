import { createClient as createSupabaseClient } from "@supabase/supabase-js";

/**
 * Service-role client - bypasses RLS entirely. Used ONLY on the server,
 * ONLY for the token-based upload flow, where there is no logged-in user
 * to scope by. The token itself IS the authorization check here, done
 * explicitly in application code (validity, expiry, single-use) - this
 * client just needs to be able to read/write regardless of RLS to do
 * that check and act on it.
 */
export function createServiceClient() {
  return createSupabaseClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.SUPABASE_SERVICE_ROLE_KEY!
  );
}
