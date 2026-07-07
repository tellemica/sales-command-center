import { createClient } from "@supabase/supabase-js";

// These come from your Vercel environment variables (see DEPLOY_GUIDE.md).
// Vite exposes vars prefixed with VITE_ to the browser.
const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // Helpful console message if env vars aren't set yet.
  console.error(
    "Missing Supabase env vars. Set VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY."
  );
}

export const supabase = createClient(url, anonKey);
