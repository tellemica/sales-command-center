import { createClient } from "@supabase/supabase-js";

// Browser uses the PUBLISHABLE key (safe to expose; protected by RLS).
// Values fall back to hardcoded constants so the app works even if the
// Vercel environment variables are not injected at build time.
const FALLBACK_URL = "https://tgeoetccxosekvdvczyc.supabase.co";
const FALLBACK_KEY = "sb_publishable__gKfKtHO8hwbBeTKIb7KIw_AJ2EDcCx";

const url = import.meta.env.VITE_SUPABASE_URL || FALLBACK_URL;
const publishableKey = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || FALLBACK_KEY;

export const supabase = createClient(url, publishableKey);
