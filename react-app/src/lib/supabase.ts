import { createClient } from '@supabase/supabase-js'
import type { AppSettings, Inventory, Metadata, SalesAnalysis } from './types'

const url = import.meta.env.VITE_SUPABASE_URL as string
const key = import.meta.env.VITE_SUPABASE_KEY as string

export const sb = url && key ? createClient(url, key) : null

async function sbLoad(id: string): Promise<unknown | null> {
  if (!sb) return null
  try {
    const r = await sb.from('app_data').select('data').eq('id', id)
    return (r.data as {data: unknown}[])?.[0]?.data ?? null
  } catch { return null }
}

async function sbSave(id: string, data: unknown): Promise<boolean> {
  if (!sb) return false
  try {
    await sb.from('app_data').upsert({ id, data })
    return true
  } catch { return false }
}

// ── Local storage fallback ───────────────────────────────
function lsGet<T>(key: string, fallback: T): T {
  try {
    const v = localStorage.getItem(key)
    return v ? JSON.parse(v) : fallback
  } catch { return fallback }
}

function lsSet(key: string, val: unknown) {
  localStorage.setItem(key, JSON.stringify(val))
}

// ── Public API ────────────────────────────────────────────
export async function loadSettings(): Promise<AppSettings> {
  const d = await sbLoad('settings')
  if (d) return d as AppSettings
  return lsGet('ajw_settings', {
    lead_time_default: 60,
    colors: { main_header: '1F3864', year_23: '2F5597', year_24: '2E75B6', year_25: '155480' },
  })
}

export function saveSettings(s: AppSettings) {
  lsSet('ajw_settings', s)
}

export async function loadMetadata(): Promise<Metadata> {
  const d = await sbLoad('metadata')
  if (d) return d as Metadata
  return lsGet('ajw_metadata', { cable: {}, housing: {} })
}

export function saveMetadata(m: Metadata) {
  lsSet('ajw_metadata', m)
}

export async function loadInventory(): Promise<Inventory> {
  const d = await sbLoad('inventory')
  if (d) return d as Inventory
  return lsGet('ajw_inventory', { cable: {}, housing: {} })
}

export function saveInventory(inv: Inventory) {
  lsSet('ajw_inventory', inv)
}

export async function loadSalesAnalysis(): Promise<SalesAnalysis> {
  const d = await sbLoad('sales')
  if (d) return d as SalesAnalysis
  return lsGet('ajw_sales', {})
}

export function saveSalesAnalysis(s: SalesAnalysis) {
  lsSet('ajw_sales', s)
}

export async function syncToSupabase(
  settings: AppSettings,
  metadata: Metadata,
  inventory: Inventory,
  sales: SalesAnalysis,
): Promise<Record<string, boolean>> {
  const [r1, r2, r3, r4] = await Promise.all([
    sbSave('settings', settings),
    sbSave('metadata', metadata),
    sbSave('inventory', inventory),
    sbSave('sales', sales),
  ])
  return { 설정: r1, '품번 메타': r2, 재고: r3, '판매 분석': r4 }
}
