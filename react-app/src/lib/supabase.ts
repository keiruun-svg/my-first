import { createClient } from '@supabase/supabase-js'
import type { AppSettings, Inventory, Metadata, OjcRules, SalesAnalysis } from './types'
import { DEFAULT_OJC_RULES } from './types'
import type { SalesAggResult } from './aggregate/salesAgg'

const url = import.meta.env.VITE_SUPABASE_URL as string
const key = import.meta.env.VITE_SUPABASE_KEY as string

export const sb = url && key ? createClient(url, key) : null

// MaterialManager 호환성 — 관리자 게이트(비밀번호)로 접근 제어하므로 항상 true
export const CAN_WRITE = true

// ── Supabase 연동 토글 (localStorage 전용) ───────────────────────
// 연동 여부는 로컬에서 결정. 환경(DEV/PROD)과 무관하게 동작.
export function getSyncEnabled(): boolean {
  try { return JSON.parse(localStorage.getItem('ajw_supabase_sync') ?? 'false') === true } catch { return false }
}

export function setSyncEnabled(enabled: boolean): void {
  localStorage.setItem('ajw_supabase_sync', JSON.stringify(enabled))
}

// ── Local storage helpers ────────────────────────────────────────
function lsGet<T>(key: string, fallback: T): T {
  try {
    const v = localStorage.getItem(key)
    return v ? JSON.parse(v) as T : fallback
  } catch { return fallback }
}

function lsSet(key: string, val: unknown) {
  localStorage.setItem(key, JSON.stringify(val))
}

// ── Supabase helpers (연동 ON일 때만 실행) ───────────────────────
async function sbLoad(id: string): Promise<unknown | null> {
  if (!sb || !getSyncEnabled()) return null
  try {
    const r = await sb.from('app_data').select('data').eq('id', id)
    return (r.data as { data: unknown }[])?.[0]?.data ?? null
  } catch { return null }
}

async function sbSave(id: string, data: unknown): Promise<boolean> {
  if (!sb || !getSyncEnabled()) return false
  try {
    await sb.from('app_data').upsert({ id, data })
    return true
  } catch { return false }
}

// ── 로드 패턴: localStorage 우선, Supabase는 연동 ON일 때 갱신 ──
async function load<T>(lsKey: string, sbKey: string, fallback: T): Promise<T> {
  const local = lsGet<T | null>(lsKey, null)
  if (!getSyncEnabled()) return local ?? fallback
  try {
    const remote = await sbLoad(sbKey)
    if (remote != null) {
      lsSet(lsKey, remote)   // 원격값을 로컬에 캐시
      return remote as T
    }
  } catch { /* ignore */ }
  return local ?? fallback
}

// ── Public API ────────────────────────────────────────────────────

export async function loadSettings(): Promise<AppSettings> {
  return load('ajw_settings', 'settings', {
    lead_time_default: 60,
    safety_stock_k: 1.5,
    colors: { main_header: '1F3864', year_23: '2F5597', year_24: '2E75B6', year_25: '155480' },
  })
}

export function saveSettings(s: AppSettings) {
  lsSet('ajw_settings', s)
  void sbSave('settings', s)
}

export async function loadMetadata(): Promise<Metadata> {
  return load('ajw_metadata', 'metadata', { cable: {}, housing: {}, ferrule: {} })
}

export function saveMetadata(m: Metadata) {
  lsSet('ajw_metadata', m)
  void sbSave('metadata', m)
}

export async function loadInventory(): Promise<Inventory> {
  return load('ajw_inventory', 'inventory', { cable: {}, housing: {}, ferrule: {} })
}

export function saveInventory(inv: Inventory) {
  lsSet('ajw_inventory', inv)
  void sbSave('inventory', inv)
}

export async function loadSalesAnalysis(): Promise<SalesAnalysis> {
  return load('ajw_sales', 'sales', {})
}

export function saveSalesAnalysis(s: SalesAnalysis) {
  lsSet('ajw_sales', s)
  void sbSave('sales', s)
}

export async function loadSalesAgg(): Promise<SalesAggResult | null> {
  return load<SalesAggResult | null>('ajw_sales_agg', 'sales_agg', null)
}

export function saveSalesAgg(s: SalesAggResult) {
  lsSet('ajw_sales_agg', s)
  void sbSave('sales_agg', s)
}

export async function loadOjcRules(): Promise<OjcRules> {
  return load('ajw_ojc_rules', 'ojc_rules', DEFAULT_OJC_RULES)
}

export async function saveOjcRules(rules: OjcRules): Promise<void> {
  lsSet('ajw_ojc_rules', rules)
  void sbSave('ojc_rules', rules)
}

// ── 이카운트/EMP 품목 리스트 ─────────────────────────────────
export interface ItemCode {
  code: string
  name: string
  spec: string
}

export async function loadItemCodes(): Promise<ItemCode[]> {
  return load('ajw_item_codes', 'item_codes', [])
}

export function saveItemCodes(codes: ItemCode[]): void {
  lsSet('ajw_item_codes', codes)
  void sbSave('item_codes', codes)
}

// ── 판매 현황 상세 데이터 (판매현황 분석 탭 → 수입 발주계획 공유) ──
export async function loadDetailedSales<T>(): Promise<T[]> {
  return lsGet<T[]>('ajw_detailed_sales', [])
}

export function saveDetailedSales<T>(rows: T[]): void {
  lsSet('ajw_detailed_sales', rows)
}

// ── OJC 생성 품번 이력 ─────────────────────────────────────────
export interface OjcProduct {
  name:       string
  spec:       string
  partNumber: string
  createdAt:  string
}

export async function loadOjcProducts(): Promise<OjcProduct[]> {
  return load('ajw_ojc_products', 'ojc_products', [])
}

export function saveOjcProducts(products: OjcProduct[]): void {
  lsSet('ajw_ojc_products', products)
  void sbSave('ojc_products', products)
}

// ── vendors (업체코드 관리) ────────────────────────────────────
export interface Vendor {
  code: string
  name: string
}

export async function loadVendors(): Promise<Vendor[]> {
  const local = lsGet<Vendor[]>('ajw_vendors', [])
  if (!sb || !getSyncEnabled()) return local
  try {
    const r = await sb.from('vendors').select('*').order('code')
    const data = (r.data as Vendor[]) ?? []
    if (data.length > 0) lsSet('ajw_vendors', data)
    return data.length > 0 ? data : local
  } catch { return local }
}

export async function upsertVendor(v: Vendor): Promise<boolean> {
  const updated = [...lsGet<Vendor[]>('ajw_vendors', []).filter(x => x.code !== v.code), v]
    .sort((a, b) => a.code.localeCompare(b.code))
  lsSet('ajw_vendors', updated)
  if (!sb || !getSyncEnabled()) return true
  try { await sb.from('vendors').upsert(v); return true } catch { return false }
}

export async function deleteVendor(code: string): Promise<boolean> {
  lsSet('ajw_vendors', lsGet<Vendor[]>('ajw_vendors', []).filter(x => x.code !== code))
  if (!sb || !getSyncEnabled()) return true
  try { await sb.from('vendors').delete().eq('code', code); return true } catch { return false }
}

// ── 전체 Supabase 동기화 (수동) ──────────────────────────────────
export async function syncToSupabase(
  settings: AppSettings,
  metadata: Metadata,
  inventory: Inventory,
  sales: SalesAnalysis,
  salesAgg?: SalesAggResult | null,
): Promise<Record<string, boolean>> {
  if (!sb) return { 연결: false }
  const up = (id: string, data: unknown) =>
    Promise.resolve(sb!.from('app_data').upsert({ id, data })).then(() => true).catch(() => false)
  const [r1, r2, r3, r4, r5] = await Promise.all([
    up('settings',  settings),
    up('metadata',  metadata),
    up('inventory', inventory),
    up('sales',     sales),
    salesAgg ? up('sales_agg', salesAgg) : Promise.resolve(true),
  ])
  return { 설정: r1, '품번 메타': r2, 재고: r3, '판매 분석': r4, '판매 집계': r5 }
}

// ── 피그테일 키 마이그레이션 ─────────────────────────────────────
const PIGTAIL_KEY_RENAME: [string, string][] = [
  ['pigtail-연청', 'pigtail-AQUA'],
  ['pigtail-연등', 'pigtail-rose'],
]

function renameCableKeys<T>(obj: Record<string, T>): { result: Record<string, T>; count: number } {
  let count = 0
  const result: Record<string, T> = {}
  for (const [k, v] of Object.entries(obj)) {
    let newKey = k
    for (const [from, to] of PIGTAIL_KEY_RENAME) {
      if (k.includes(from)) { newKey = k.replace(from, to); count++; break }
    }
    result[newKey] = v
  }
  return { result, count }
}

export async function migratePigtailKeys(
  metadata:  Metadata,
  inventory: Inventory,
): Promise<{ metadata: Metadata; inventory: Inventory; changed: number }> {
  const { result: newCableMeta, count: c1 } = renameCableKeys(metadata.cable  ?? {})
  const { result: newCableInv,  count: c2 } = renameCableKeys(inventory.cable ?? {})
  const changed = c1 + c2
  const newMeta: Metadata  = { ...metadata,  cable: newCableMeta }
  const newInv:  Inventory = { ...inventory, cable: newCableInv  }
  if (changed > 0) { saveMetadata(newMeta); saveInventory(newInv) }
  return { metadata: newMeta, inventory: newInv, changed }
}
