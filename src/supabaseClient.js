import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const hasSupabaseConfig = Boolean(
  supabaseUrl &&
  supabaseKey &&
  !supabaseKey.includes("put_your_anon_key_here")
);

export const supabase = hasSupabaseConfig ? createClient(supabaseUrl, supabaseKey) : null;
