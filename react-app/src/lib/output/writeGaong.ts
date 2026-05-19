import * as XLSX from 'xlsx'
import type { YearPivot, PivotRow } from '../aggregate/pivot'

// ── 하우징 컬럼 매핑 (1-based 인덱스) ───────────────────────
const MM_KINDS    = new Set(['om1', 'om1-pigtail', 'om3'])
const PRIMARY_20: Record<string, number>   = { 'LC/PC':23,'LC/APC':25,'SC/PC':27,'SC/APC':29,'FC/PC':31,'FC/APC':33 }
const SECONDARY_20: Record<string, number> = { 'LC/PC':24,'LC/APC':26,'SC/PC':28,'SC/APC':30,'FC/PC':32,'FC/APC':34 }
const BEIGE_COL = 35
const PRIMARY_30: Record<string, number>   = { 'LC/PC':36,'LC/APC':38,'SC/PC':40,'SC/APC':42,'FC/PC':44,'FC/APC':46 }
const SECONDARY_30: Record<string, number> = { 'LC/PC':37,'LC/APC':39,'SC/PC':41,'SC/APC':43,'FC/PC':45,'FC/APC':47 }

// 1-based 하우징 컬럼 → 이름
const HOUSING_COL_NAMES: Record<number, string> = {
  23:'2.0mm - LC/PC 청색',  24:'2.0mm - LC/PC 적색',
  25:'2.0mm - LC/APC 녹색', 26:'2.0mm - LC/APC 적색',
  27:'2.0mm - SC/PC 청색',  28:'2.0mm - SC/PC 적색',
  29:'2.0mm - SC/APC 녹색', 30:'2.0mm - SC/APC 적색',
  31:'2.0mm - FC/PC 흑색',  32:'2.0mm - FC/PC 적색',
  33:'2.0mm - FC/APC 녹색', 34:'2.0mm - FC/APC 적색',
  35:'2.0mm - LC/PC 베이지MM',
  36:'3.0mm - LC/PC 청색',  37:'3.0mm - LC/PC 적색',
  38:'3.0mm - LC/APC 녹색', 39:'3.0mm - LC/APC 적색',
  40:'3.0mm - SC/PC 청색',  41:'3.0mm - SC/PC 적색',
  42:'3.0mm - SC/APC 녹색', 43:'3.0mm - SC/APC 적색',
  44:'3.0mm - FC/PC 흑색',  45:'3.0mm - FC/PC 적색',
  46:'3.0mm - FC/APC 녹색', 47:'3.0mm - FC/APC 적색',
}

const MONTHS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']

function isMulticore(kind: string) { return /^(b3|a1)-\d+c$/i.test(kind) }

// ── 케이블 타입 레이블 ───────────────────────────────────────
function kindLabel(kind: string, core: number): string {
  const sd = core === 1 ? 'SP' : core === 2 ? 'DP' : `${core}C`
  const map: Record<string, string> = {
    'a1':'A1','b3':'B3','om1':'OM1','om3':'OM3',
    'a1-청':'A1_청','a1-녹':'A1_녹','a1-적':'A1_적','a1-자':'A1_자',
  }
  if (map[kind])  return `${map[kind]}-${sd}`
  if (kind === 'drop') return `DROP-${sd}`
  if (kind === 'pigtail' || kind === 'om1-pigtail') return 'PIGTAIL'
  if (kind === 'a2') return 'Optical cable'
  return kind.toUpperCase()
}

// ── 하우징 월별 계산 (총합이 아닌 월별 분리) ─────────────────
function calcHousingMonthly(row: PivotRow): Record<number, number[]> {
  const { type1: t1, type2: t2, pai, kind, core } = row
  if (!t1) return {}
  const isMm = MM_KINDS.has(kind), useRed = core >= 2 && !isMulticore(kind)
  const is20 = pai === '2.0mm', is30 = pai === '3.0mm'
  const res: Record<number, number[]> = {}

  const add = (col: number | undefined) => {
    if (!col) return
    if (!res[col]) res[col] = new Array(12).fill(0)
    for (let i = 0; i < 12; i++) res[col][i] += row.monthly[i] * core
  }
  const addPrimary = (t: string) => {
    if (t === 'LC/PC' && isMm && is20) add(BEIGE_COL)
    else if (is20 && PRIMARY_20[t])    add(PRIMARY_20[t])
    else if (is30 && PRIMARY_30[t])    add(PRIMARY_30[t])
  }

  if (t1) addPrimary(t1)
  if (t2) {
    if (useRed) {
      if (is20 && SECONDARY_20[t2])      add(SECONDARY_20[t2])
      else if (is30 && SECONDARY_30[t2]) add(SECONDARY_30[t2])
    } else {
      addPrimary(t2)
    }
  }
  return res
}

// 총합용 (하우징 원본 시트에서 사용)
function calcHousing(row: PivotRow): Record<number, number> {
  const monthly = calcHousingMonthly(row)
  return Object.fromEntries(
    Object.entries(monthly).map(([col, arr]) => [col, arr.reduce((a, b) => a + b, 0)])
  )
}

// ── 집계 시트 빌드 ───────────────────────────────────────────
interface AggEntry { monthly: number[] }

const PIGTAIL_COLORS_AGG = ['청','등','녹','적','황','자','갈','흑','백','회','연청','연등']
const PIGTAIL_COLOR_MAP: Record<string, string> = { '연청': 'AQUA', '연등': 'rose' }

function buildCableAgg(rows: PivotRow[]): Map<string, AggEntry & { pai: string; unit: string }> {
  const agg = new Map<string, AggEntry & { pai: string; unit: string }>()

  for (const row of rows) {
    const isPigtail = row.kind === 'pigtail' || row.kind === 'om1-pigtail'

    // 0.9mm 피그테일: 코어수만큼 색상별 분리 (각 색상이 동일 수량)
    if (isPigtail && row.pai === '0.9mm') {
      const nc = Math.max(1, row.core || 1)
      for (const raw of PIGTAIL_COLORS_AGG.slice(0, nc)) {
        const color = PIGTAIL_COLOR_MAP[raw] ?? raw
        const key = `0.9mm|pigtail-${color}`
        if (!agg.has(key)) agg.set(key, { pai: '0.9mm', unit: '개', monthly: new Array(12).fill(0) })
        const entry = agg.get(key)!
        for (let i = 0; i < 12; i++) entry.monthly[i] += row.monthly[i]
      }
      continue
    }

    const label  = kindLabel(row.kind, row.core)
    const key    = `${row.pai}|${label}`
    const hasLen = row.length != null && row.length > 0 && !isPigtail
    const unit   = hasLen ? 'm' : '개'

    if (!agg.has(key)) agg.set(key, { pai: row.pai, unit, monthly: new Array(12).fill(0) })
    const entry = agg.get(key)!
    for (let i = 0; i < 12; i++) {
      entry.monthly[i] += hasLen
        ? Math.round(row.monthly[i] * (row.length ?? 0))
        : row.monthly[i]
    }
  }
  return agg
}

function buildHousingAgg(rows: PivotRow[]): Map<string, AggEntry> {
  const agg = new Map<string, AggEntry>()

  for (const row of rows) {
    const byCol = calcHousingMonthly(row)
    for (const [col, monthly] of Object.entries(byCol)) {
      const name = HOUSING_COL_NAMES[parseInt(col)]
      if (!name) continue
      if (!agg.has(name)) agg.set(name, { monthly: new Array(12).fill(0) })
      const entry = agg.get(name)!
      for (let i = 0; i < 12; i++) entry.monthly[i] += monthly[i]
    }
  }
  return agg
}

// ── 집계 시트 작성 ───────────────────────────────────────────
const PAI_ORDER: Record<string, number> = { '2.0mm': 0, '3.0mm': 1, '0.9mm': 2 }

function writeCableAggSheet(
  wb: XLSX.WorkBook,
  agg: ReturnType<typeof buildCableAgg>,
  YY: string,
) {
  const header = ['케이블 타입', '파이', '단위', ...MONTHS, '연간 합계', '월 최대', '최대 발생월']
  const data: (string | number | null)[][] = [header]

  const entries = [...agg.entries()].sort((a, b) => {
    const [pa] = a[0].split('|'); const [pb] = b[0].split('|')
    return (PAI_ORDER[pa] ?? 9) - (PAI_ORDER[pb] ?? 9) || a[0].localeCompare(b[0])
  })

  for (const [key, { pai, unit, monthly }] of entries) {
    const label   = key.split('|')[1]
    const annual  = monthly.reduce((a, b) => a + b, 0)
    const peak    = Math.max(...monthly)
    const peakMon = monthly.indexOf(peak) + 1  // 1-based 월
    if (annual === 0) continue
    data.push([label, pai, unit, ...monthly.map(v => v || null), annual || null, peak || null, `${peakMon}월`])
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), `${YY}년_케이블_집계`)
}

function writeHousingAggSheet(
  wb: XLSX.WorkBook,
  agg: ReturnType<typeof buildHousingAgg>,
  YY: string,
) {
  const header = ['하우징 타입', ...MONTHS, '연간 합계', '월 최대', '최대 발생월']
  const data: (string | number | null)[][] = [header]

  // 컬럼 번호 순서대로 정렬 (HOUSING_COL_NAMES의 키 순)
  const colOrder = Object.values(HOUSING_COL_NAMES)
  const entries  = [...agg.entries()].sort(
    (a, b) => colOrder.indexOf(a[0]) - colOrder.indexOf(b[0])
  )

  for (const [name, { monthly }] of entries) {
    const annual  = monthly.reduce((a, b) => a + b, 0)
    const peak    = Math.max(...monthly)
    const peakMon = monthly.indexOf(peak) + 1
    if (annual === 0) continue
    data.push([name, ...monthly.map(v => v || null), annual || null, peak || null, `${peakMon}월`])
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), `${YY}년_하우징_집계`)
}

// ── 원본 시트 헤더 ───────────────────────────────────────────
function cableHeader(YY: string): string[] {
  return [
    '품목코드','품목명','규격명','케이블 종류','파이','코어수','케이블 길이','타입 1','타입2',
    ...Array.from({ length: 12 }, (_, i) => `${YY}년${String(i + 1).padStart(2, '0')}월`),
    '케이블 사용량','월 최대 수량','월 최대 케이블 소요량',
  ]
}

function housingHeader(YY: string): string[] {
  return [
    '품목코드','품목명','규격명','케이블 종류','파이','코어수','케이블 길이','타입 1','타입2',
    ...Array.from({ length: 12 }, (_, i) => `${YY}년${String(i + 1).padStart(2, '0')}월`),
    '합계',
    '2.0MM - LC/PC(청색)','2.0MM - LC/PC(적색)','2.0MM - LC/APC(녹색)','2.0MM - LC/APC(적색)',
    '2.0MM - SC/PC(청색)','2.0MM - SC/PC(적색)','2.0MM - SC/APC(녹색)','2.0MM - SC/APC(적색)',
    '2.0MM - FC/PC(흑색)','2.0MM - FC/PC(적색)','2.0MM - FC/APC(녹색)','2.0MM - FC/APC(적색)',
    '2.0MM - BEIGE(OM1·OM3)',
    '3.0MM - LC/PC(청색)','3.0MM - LC/PC(적색)','3.0MM - LC/APC(녹색)','3.0MM - LC/APC(적색)',
    '3.0MM - SC/PC(청색)','3.0MM - SC/PC(적색)','3.0MM - SC/APC(녹색)','3.0MM - SC/APC(적색)',
    '3.0MM - FC/PC(흑색)','3.0MM - FC/PC(적색)','3.0MM - FC/APC(녹색)','3.0MM - FC/APC(적색)',
    '검증','계산수량',
  ]
}

// ── 메인 ────────────────────────────────────────────────────
export function writeGaong(pivot: YearPivot): ArrayBuffer {
  const wb = XLSX.utils.book_new()

  for (const [year, rows] of [...pivot.entries()].sort()) {
    const YY    = year.slice(-2)
    const cData: (string | number | null)[][] = [cableHeader(YY)]
    const hData: (string | number | null)[][] = [housingHeader(YY)]

    for (const row of rows) {
      const total    = row.monthly.reduce((a, b) => a + b, 0)
      const peak     = Math.max(...row.monthly)
      const len      = row.length ?? 0
      const usage    = Math.round(total * len)
      const peakUsage = Math.round(peak * len)
      const base     = [row.code, row.name, row.spec, row.kind, row.pai, row.core, row.length ?? null, row.type1 || null, row.type2 || null]

      cData.push([...base, ...row.monthly.map(v => v || null), usage || null, peak || null, peakUsage || null])

      const hRow: (string | number | null)[] = [
        ...base,
        ...row.monthly.map(v => v || null),
        total || null,
        ...new Array(25).fill(null),
        null, null,
      ]
      const housing = calcHousing(row)
      for (const [col1, val] of Object.entries(housing))
        hRow[parseInt(col1) - 1] = val || null
      hRow[47] = row.core === 1 ? total * 2 : total * row.core * 2
      hRow[48] = Object.values(housing).reduce((a, b) => a + b, 0) || null
      hData.push(hRow)
    }

    // 집계 시트
    const cableAgg   = buildCableAgg(rows)
    const housingAgg = buildHousingAgg(rows)

    // 시트 순서: 집계(요약) → 원본(검증)
    writeCableAggSheet(wb, cableAgg, YY)
    writeHousingAggSheet(wb, housingAgg, YY)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cData), `${YY}년_케이블`)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(hData), `${YY}년 하우징`)
  }

  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as number[]
  return new Uint8Array(out).buffer as ArrayBuffer
}
