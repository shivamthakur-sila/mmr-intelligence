import { createServiceClient } from "./supabase/service";
import type { ChecklistItem } from "./curate";

/**
 * Pulls the checklist that actually applies to this site, using the
 * site -> client -> default resolution proven in resolve_site_config.
 *
 * The catalogue holds every parameter we've seen across real MMRs;
 * `enabled` decides which apply here. A site or client turns one on by
 * inserting a row for the same key at its own tier, and the resolver
 * takes the most specific row per key. Filtering happens here rather
 * than in SQL so the catalogue stays queryable for a future admin UI.
 */
export async function loadChecklist(siteId: string): Promise<ChecklistItem[]> {
  const supabase = createServiceClient();

  const { data, error } = await supabase.rpc("resolve_site_config", {
    p_site_id: siteId,
    p_kind: "checklist_parameter",
  });

  if (error) {
    throw new Error(`Couldn't load this site's checklist: ${error.message}`);
  }

  type Row = {
    key: string;
    value: { ask_type?: string; order?: number; enabled?: boolean; fields?: string[] };
    description: string | null;
  };
  const rows = (data ?? []) as Row[];

  if (rows.length === 0) {
    throw new Error(
      "No checklist parameters are configured. Run seed_checklist_catalogue.sql against the database."
    );
  }

  const enabled = rows.filter((r) => r.value?.enabled === true);

  if (enabled.length === 0) {
    throw new Error(
      `Every checklist parameter is switched off for this site. Enable at least one in site_config.`
    );
  }

  return enabled
    .slice()
    .sort((a, b) => (a.value?.order ?? 999) - (b.value?.order ?? 999))
    .map((r) => ({
      key: r.key,
      label: r.description ?? r.key,
      askType: (r.value?.ask_type as ChecklistItem["askType"]) ?? "text",
      fields: r.value?.fields ?? [],
    }));
}
