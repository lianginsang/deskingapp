import { createClient } from '@supabase/supabase-js';

// This is a "publishable" key (Supabase's newer, explicitly-public anon key
// format) — it's meant to ship in client-side bundles. Access control is
// enforced entirely by the Row Level Security policies on the tables it can
// reach, not by keeping this value secret.
const SUPABASE_URL = 'https://rjysdfrmchfdvgmwdziy.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_JgqXJnfW0p27m5armLXxcA_1JGbKEQ-';

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
