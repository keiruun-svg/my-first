import * as XLSX from 'xlsx'
import type { Metadata, Inventory, SalesAnalysis, YearStats } from './types'

const MM_KINDS = new Set(['om1','om1-pigtail','om3'])
const PAI_ORDER: Record<string, number> = { '2.0mm': 0, '3.0mm': 1, '0.9mm': 2 }

function normalPai(p: unknown): string {
  const s = String(p ?? '').trim()
  return s === '0.9' ? '0.9mm' : s
}

function isMultiCore(kind: string): boolean {
  return /^(b3|a1)-\d+c$/i.test(kind)
}

function mc2(kind: unknown, core: unknown): string {
  const k = String(kind ?? '').trim().toLowerCase()
  let c = 1
  try { c = parseInt(String(core)) || 1 } catch {}
  const sd = c === 1 ? 'SP' : c === 2 ? 'DP' : `${c}C`
  const bm: Record<string, string> = {
    'a1':'A1','b3':'B3','om1':'OM1','om3':'OM3',
    'a1-청':'A1_청','a1-녹':'A1_녹','a1-적':'A1_적','a1-자':'A1_자',
  }
  if (bm[k]) return `${bm[k]}-${sd}`
  if (k === 'drop') return 'DROP'
  if (k === 'pigtail' || k === 'om1-pigtail') return 'PIGTAIL'
  if (k === 'a2') return 'Optical cable'
  return k.toUpperCase()
}

function getHousingPrimary(t: unknown, kind: unknown, pai: string): [string,string] | null {
  if (!t) return null
  const ts = String(t).trim(); const p = normalPai(pai); const k = String(kind ?? '').toLowerCase()
  if (ts === 'LC/PC') return MM_KINDS.has(k) && p === '2.0mm' ? [p,'LC/PC 베이지MM'] : [p,'LC/PC 청색']
  const m: Record<string, [string,string]> = {
    'LC/APC':[p,'LC/APC 녹색'],'SC/PC':[p,'SC/PC 청색'],
    'SC/APC':[p,'SC/APC 녹색'],'FC/PC':[p,'FC/PC 흑색'],'FC/APC':[p,'FC/APC 녹색'],
  }
  return m[ts] ?? null
}

function getHousingRed(t: unknown, pai: string): [string,string] | null {
  if (!t) return null
  const ts = String(t).trim(); const p = normalPai(pai)
  const m: Record<string, [string,string]> = {
    'LC/PC':[p,'LC/PC 적색'],'LC/APC':[p,'LC/APC 적색'],'SC/PC':[p,'SC/PC 적색'],
    'SC/APC':[p,'SC/APC 적색'],'FC/PC':[p,'FC/PC 적색'],'FC/APC':[p,'FC/APC 적색'],
  }
  return m[ts] ?? null
}

function ltDays(lt: unknown, def: number): number {
  if (!lt) return def
  const m = String(lt).match(/(\d+)/)
  return m ? parseInt(m[1]) : def
}

export interface Step2Stats {
  cableStats: Record<string, Record<string, YearStats>>
  housingStats: Record<string, Record<string, YearStats>>
  years: string[]
  logs: string[]
}

export function aggregateStats(fileBuffer: ArrayBuffer): Step2Stats {
  const logs: string[] = []
  const wb = XLSX.read(fileBuffer, { type: 'array' })
  const sheets = new Set(wb.SheetNames)

  const activeYears = wb.SheetNames
    .map(s => s.match(/^(\d{2})년[_ ]케이블/)?.[1])
    .filter(Boolean)
    .sort() as string[]

  if (!activeYears.length) {
    logs.push('연도 감지 실패 — 기본값 23/24/25 사용')
    activeYears.push('23','24','25')
  }

  const cableAgg: Record<string, Record<string, number[]>> = {}
  const housingAgg: Record<string, Record<string, number[]>> = {}

  const initKey = (agg: typeof cableAgg, key: string) => {
    if (!agg[key]) agg[key] = {}
    for (const yr of activeYears) if (!agg[key][yr]) agg[key][yr] = new Array(12).fill(0)
  }

  const findSheet = (yr: string, kind: string) =>
    [`${yr}년_${kind}`, `${yr}년 ${kind}`].find(n => sheets.has(n)) ?? null

  for (const yr of activeYears) {
    const cs = findSheet(yr, '케이블')
    const hs = findSheet(yr, '하우징')

    if (cs) {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[cs], { header: 1, defval: null }) as unknown[][]
      for (const row of rows.slice(1)) {
        const kind = row[3], pai = row[4], core = row[5], length = row[6]
        if (!kind || !pai || !length) continue
        const ct = mc2(kind, core); const p = normalPai(pai)
        const len = parseFloat(String(length)) || 0
        if (ct === 'PIGTAIL' && p === '0.9mm') continue
        const key = `${p}|${ct}`; initKey(cableAgg, key)
        for (let i = 0; i < 12; i++) {
          const q = row[9 + i]; if (q) cableAgg[key][yr][i] += parseFloat(String(q)) * len
        }
      }
    }

    if (hs) {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[hs], { header: 1, defval: null }) as unknown[][]
      for (const row of rows.slice(1)) {
        const kind = row[3], pai = row[4], core = row[5], t1 = row[7], t2 = row[8]
        if (!pai) continue
        const ph = normalPai(pai)
        let cps = 1; try { cps = parseInt(String(core)) || 1 } catch {}
        const isDp = cps >= 2 && !isMultiCore(String(kind ?? ''))
        for (let mi = 0; mi < 12; mi++) {
          const qty = row[9 + mi]; if (!qty) continue
          const q = parseFloat(String(qty))
          const ha = getHousingPrimary(t1, kind, ph)
          if (ha) { const hk = `${ha[0]}|${ha[1]}`; initKey(housingAgg, hk); housingAgg[hk][yr][mi] += q * cps }
          if (t2) {
            const hb = isDp ? getHousingRed(t2, ph) : getHousingPrimary(t2, kind, ph)
            if (hb) { const hk = `${hb[0]}|${hb[1]}`; initKey(housingAgg, hk); housingAgg[hk][yr][mi] += q * cps }
          }
        }
      }
    }
  }

  const finalize = (agg: typeof cableAgg): Record<string, Record<string, YearStats>> => {
    const r: Record<string, Record<string, YearStats>> = {}
    for (const [k, yd] of Object.entries(agg)) {
      r[k] = {}
      for (const [yr, m] of Object.entries(yd)) {
        const v = m.map(Math.round)
        r[k][yr] = { monthly: v, annual: v.reduce((a,b)=>a+b,0), peak: Math.max(...v) }
      }
    }
    return r
  }

  logs.push(`감지된 연도: ${activeYears.map(y=>'20'+y+'년').join(', ')}`)
  return { cableStats: finalize(cableAgg), housingStats: finalize(housingAgg), years: activeYears, logs }
}

export interface OrderPlanRow {
  type: 'cable' | 'housing'
  pai: string
  ctype: string
  품번: string
  품명: string
  구매처: string
  리드타임: number
  yearStats: Record<string, YearStats>
  avgAnnual: number
  avgPeak: number
  trend2324: number | null
  trend2425: number | null
  안전재고: number
  현재고: number
  기발주: number
  제안량?: number
}

export function buildOrderPlan(
  stats: Step2Stats,
  metadata: Metadata,
  inventory: Inventory,
  sales: SalesAnalysis,
  ltDefault: number,
): OrderPlanRow[] {
  const rows: OrderPlanRow[] = []
  const { cableStats, housingStats, years } = stats

  const activeYears = years.filter(yr => ['23','24','25'].includes(yr))

  const addRows = (
    statsMap: Record<string, Record<string, YearStats>>,
    type: 'cable' | 'housing',
  ) => {
    const keys = Object.keys(statsMap)
      .filter(k => activeYears.some(yr => (statsMap[k][yr]?.annual ?? 0) > 0))
      .sort((a, b) => {
        const [pa] = a.split('|'); const [pb] = b.split('|')
        return (PAI_ORDER[pa] ?? 9) - (PAI_ORDER[pb] ?? 9) || a.localeCompare(b)
      })

    for (const key of keys) {
      const [pai, ctype] = key.split('|')
      const d = statsMap[key]
      const metaRaw = type === 'cable' ? metadata.cable[key] : metadata.housing[key]
      const metas = metaRaw
        ? (Array.isArray(metaRaw) ? metaRaw : [metaRaw])
        : [{}]
      const invRaw = type === 'cable' ? inventory.cable[key] : inventory.housing[key]
      const invs = invRaw
        ? (Array.isArray(invRaw) ? invRaw : [invRaw])
        : [{ 현재고: 0, 기발주: 0 }]

      const yrs23 = d['23']?.annual ?? 0
      const yrs24 = d['24']?.annual ?? 0
      const yrs25 = d['25']?.annual ?? 0
      const avg = Math.round((yrs23 + yrs24 + yrs25) / 3)
      const avgPeak = Math.round(
        ((d['23']?.peak ?? 0) + (d['24']?.peak ?? 0) + (d['25']?.peak ?? 0)) / 3
      )
      const trend2324 = yrs23 > 0 ? (yrs24 - yrs23) / yrs23 : null
      const trend2425 = yrs24 > 0 ? (yrs25 - yrs24) / yrs24 : null

      for (let i = 0; i < metas.length; i++) {
        const m = metas[i] as Record<string, unknown> ?? {}
        const inv = (invs[i] ?? { 현재고: 0, 기발주: 0 }) as Record<string, number>
        const lt = ltDays(m['리드타임'], ltDefault)
        const 안전재고 = Math.round(avgPeak * lt / 30)

        rows.push({
          type,
          pai,
          ctype,
          품번: String(m['품번'] ?? ''),
          품명: String(m['품명'] ?? ''),
          구매처: String(m['구매처'] ?? ''),
          리드타임: lt,
          yearStats: d,
          avgAnnual: avg,
          avgPeak,
          trend2324,
          trend2425,
          안전재고,
          현재고: Number(inv['현재고'] ?? 0),
          기발주: Number(inv['기발주'] ?? 0),
        })
      }
    }
  }

  addRows(cableStats, 'cable')
  addRows(housingStats, 'housing')

  // Sales-based 제안량
  if (Object.keys(sales).length) {
    // simplified: use 25년 sales trend to project
    for (const row of rows) {
      row.제안량 = Math.round(row.avgAnnual * 1.05)
    }
  }

  return rows
}
