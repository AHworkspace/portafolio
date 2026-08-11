import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL as string | undefined;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined;

export const isSupabaseEnabled = Boolean(supabaseUrl && supabaseAnonKey);

export const supabase = isSupabaseEnabled
  ? createClient(supabaseUrl!, supabaseAnonKey!)
  : null;

export const portfolioId = (import.meta.env.VITE_PORTFOLIO_ID as string | undefined) || 'default';
export const storageBucket = (import.meta.env.VITE_SUPABASE_BUCKET as string | undefined) || 'portfolio-files';
