import * as XLSX from 'xlsx'
import type { YearStats, Metadata } from './types'

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
    'a1': 'A1', 'b3': 'B3', 'om1': 'OM1', 'om3': 'OM3',
    'a1-청': 'A1_청', 'a1-녹': 'A1_녹', 'a1-적': 'A1_적', 'a1-자': 'A1_자',
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
    'LC/APC': [p,'LC/APC 녹색'], 'SC/PC': [p,'SC/PC 청색'],
    'SC/APC': [p,'SC/APC 녹색'], 'FC/PC': [p,'FC/PC 흑색'], 'FC/APC': [p,'FC/APC 녹색'],
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

export interface Step1Result {
  cableStats: Record<string, Record<string, YearStats>>
  housingStats: Record<string, Record<string, YearStats>>
  years: string[]
  logs: string[]
  newCableKeys: string[]
  newHousingKeys: string[]
}

export function runStep1(fileBuffer: ArrayBuffer, metadata: Metadata): Step1Result {
  const logs: string[] = []
  const wb = XLSX.read(fileBuffer, { type: 'array' })
  const sheets = new Set(wb.SheetNames)

  const findSheet = (yr: string, kind: string): string | null => {
    for (const name of [
      `${yr}년_${kind}`, `${yr}년 ${kind}`,
      `20${yr}년_${kind}`, `20${yr}년 ${kind}`,
    ]) {
      if (sheets.has(name)) return name
    }
    return null
  }

  const yrSheetMap: Record<string, { cable: string; housing: string | null }> = {}
  for (let y = 2015; y <= 2031; y++) {
    const yr = String(y).slice(-2)
    const cs = findSheet(yr, '케이블')
    if (cs) yrSheetMap[yr] = { cable: cs, housing: findSheet(yr, '하우징') }
  }

  const YEARS = Object.keys(yrSheetMap).sort()
  if (!YEARS.length) {
    logs.push('⚠ 케이블 시트를 찾지 못했습니다.')
    return { cableStats: {}, housingStats: {}, years: [], logs, newCableKeys: [], newHousingKeys: [] }
  }
  logs.push(`감지된 연도: ${YEARS.map(y => '20' + y + '년').join(', ')}`)

  const cableAgg: Record<string, Record<string, number[]>> = {}
  const housingAgg: Record<string, Record<string, number[]>> = {}

  const initKey = (agg: typeof cableAgg, key: string) => {
    if (!agg[key]) agg[key] = {}
    for (const yr of YEARS) if (!agg[key][yr]) agg[key][yr] = new Array(12).fill(0)
  }

  for (const yr of YEARS) {
    const cs = yrSheetMap[yr].cable
    const hs = yrSheetMap[yr].housing

    const cableRows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[cs], { header: 1, defval: null }) as unknown[][]
    for (const row of cableRows.slice(1)) {
      const kind = row[3], pai = row[4], core = row[5], length = row[6]
      if (!kind || !pai || !length) continue
      const ct = mc2(kind, core)
      const pc = normalPai(pai)
      const len = parseFloat(String(length)) || 0

      if (ct === 'PIGTAIL' && pc === '0.9mm') continue

      const key = `${pc}|${ct}`
      initKey(cableAgg, key)
      for (let i = 0; i < 12; i++) {
        const q = (row as unknown[])[9 + i]
        if (q) cableAgg[key][yr][i] += parseFloat(String(q)) * len
      }
    }

    if (!hs) continue
    const hRows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[hs], { header: 1, defval: null }) as unknown[][]
    for (const row of hRows.slice(1)) {
      const kind = row[3], pai = row[4], core = row[5], t1 = row[7], t2 = row[8]
      if (!pai) continue
      const ph = normalPai(pai)
      let cps = 1
      try { cps = parseInt(String(core)) || 1 } catch {}
      const isDp = cps >= 2 && !isMultiCore(String(kind ?? ''))

      for (let mi = 0; mi < 12; mi++) {
        const qty = (row as unknown[])[9 + mi]
        if (!qty) continue
        const q = parseFloat(String(qty))
        const ha = getHousingPrimary(t1, kind, ph)
        if (ha) {
          const hkey = `${ha[0]}|${ha[1]}`
          initKey(housingAgg, hkey)
          housingAgg[hkey][yr][mi] += q * cps
        }
        if (t2) {
          const hb = isDp ? getHousingRed(t2, ph) : getHousingPrimary(t2, kind, ph)
          if (hb) {
            const hkey = `${hb[0]}|${hb[1]}`
            initKey(housingAgg, hkey)
            housingAgg[hkey][yr][mi] += q * cps
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

  const cableStats = finalize(cableAgg)
  const housingStats = finalize(housingAgg)

  const existingCable = new Set(Object.keys(metadata.cable))
  const existingHousing = new Set(Object.keys(metadata.housing))
  const newCableKeys = Object.keys(cableStats).filter(k => !existingCable.has(k))
  const newHousingKeys = Object.keys(housingStats).filter(k => !existingHousing.has(k))

  logs.push(`집계 완료 — 케이블 ${Object.keys(cableStats).length}타입 / 하우징 ${Object.keys(housingStats).length}타입`)
  logs.push(`신규 케이블 ${newCableKeys.length}건 / 신규 하우징 ${newHousingKeys.length}건`)

  return { cableStats, housingStats, years: YEARS, logs, newCableKeys, newHousingKeys }
}

export function sortedCableKeys(stats: Record<string, Record<string, YearStats>>, years: string[]): string[] {
  return Object.keys(stats)
    .filter(k => years.some(yr => (stats[k][yr]?.annual ?? 0) > 0))
    .sort((a, b) => {
      const [pa] = a.split('|'); const [pb] = b.split('|')
      return (PAI_ORDER[pa] ?? 9) - (PAI_ORDER[pb] ?? 9) || a.localeCompare(b)
    })
}
