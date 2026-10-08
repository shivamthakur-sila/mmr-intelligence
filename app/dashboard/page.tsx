import { redirect } from "next/navigation";
import { createClient } from "@/lib/supabase/server";
import AskPanel from "./ask-panel";

export default async function DashboardPage() {
  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();

  if (!user) redirect("/login");

  return <AskPanel email={user.email ?? ""} />;
}
