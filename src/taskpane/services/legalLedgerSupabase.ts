// src/taskpane/services/legalLedgerSupabase.ts
import { createClient } from "@supabase/supabase-js";
import { LL_SUPABASE_URL, LL_SUPABASE_ANON_KEY } from "./llConfig";

if (!LL_SUPABASE_URL || LL_SUPABASE_URL.includes("PASTE_")) {
  throw new Error("Missing LL_SUPABASE_URL in src/taskpane/services/llConfig.ts");
}
if (!LL_SUPABASE_ANON_KEY || LL_SUPABASE_ANON_KEY.includes("PASTE_")) {
  throw new Error("Missing LL_SUPABASE_ANON_KEY in src/taskpane/services/llConfig.ts");
}

export const llSupabase = createClient(LL_SUPABASE_URL, LL_SUPABASE_ANON_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
  },
});
