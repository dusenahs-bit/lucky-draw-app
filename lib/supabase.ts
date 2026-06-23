import { createClient, type SupabaseClient } from '@supabase/supabase-js'

export interface Winner {
  id: string
  name: string
  prize: string
  prize_type: 'survey' | 'lucky'
  confirmed_at: string
}

let _supabase: SupabaseClient | null = null

export function getSupabase(): SupabaseClient {
  if (!_supabase) {
    const url = process.env.NEXT_PUBLIC_SUPABASE_URL!
    const key = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
    _supabase = createClient(url, key)
  }
  return _supabase
}
