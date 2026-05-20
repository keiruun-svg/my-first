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

// HOUSING_COL_NAMES를 정렬된 "pai|type" 키 배열로 변환 (요약 시트 정렬용)
const HOUSING_KEY_ORDER: string[] = Object.entries(HOUSING_COL_NAMES)
  .sort((a, b) => parseInt(a[0]) - parseInt(b[0]))
  .map(([, name]) => {
    const sep  = name.indexOf(' - ')
    return sep > 0 ? `${name.slice(0, sep)}|${name.slice(sep + 3)}` : name
  })

const MONTHS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']
const PAI_ORDER: Record<string, number> = { '2.0mm': 0, '3.0mm': 1, '0.9mm': 2 }

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

// ── 하우징 월별 계산 ─────────────────────────────────────────
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

function calcHousing(row: PivotRow): Record<number, number> {
  const monthly = calcHousingMonthly(row)
  return Object.fromEntries(
    Object.entries(monthly).map(([col, arr]) => [col, arr.reduce((a, b) => a + b, 0)])
  )
}

// ── 집계 빌더 ────────────────────────────────────────────────
interface AggEntry { monthly: number[] }
interface HousingAggEntry extends AggEntry { pai: string }

const PIGTAIL_COLORS_AGG = ['청','등','녹','적','황','자','갈','흑','백','회','연청','연등']
const PIGTAIL_COLOR_MAP: Record<string, string> = { '연청': 'AQUA', '연등': 'rose' }

function buildCableAgg(rows: PivotRow[]): Map<string, AggEntry & { pai: string; unit: string }> {
  const agg = new Map<string, AggEntry & { pai: string; unit: string }>()

  for (const row of rows) {
    const isPigtail = row.kind === 'pigtail' || row.kind === 'om1-pigtail'

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

function buildHousingAgg(rows: PivotRow[]): Map<string, HousingAggEntry> {
  const agg = new Map<string, HousingAggEntry>()

  for (const row of rows) {
    const byCol = calcHousingMonthly(row)
    for (const [col, monthly] of Object.entries(byCol)) {
      const fullName = HOUSING_COL_NAMES[parseInt(col)]
      if (!fullName) continue
      const sep  = fullName.indexOf(' - ')
      const pai  = sep > 0 ? fullName.slice(0, sep) : ''
      const type = sep > 0 ? fullName.slice(sep + 3) : fullName
      const key  = `${pai}|${type}`
      if (!agg.has(key)) agg.set(key, { pai, monthly: new Array(12).fill(0) })
      const entry = agg.get(key)!
      for (let i = 0; i < 12; i++) entry.monthly[i] += monthly[i]
    }
  }
  return agg
}

// ── 집계 시트 작성 ───────────────────────────────────────────
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
    const peakMon = monthly.indexOf(peak) + 1
    if (annual === 0) continue
    data.push([label, pai, unit, ...monthly.map(v => v || null), annual || null, peak || null, `${peakMon}월`])
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), `${YY}년_케이블_집계`)
}

function writeHousingAggSheet(
  wb: XLSX.WorkBook,
  agg: Map<string, HousingAggEntry>,
  YY: string,
) {
  // 케이블 집계와 동일한 컬럼 구조: 타입 | 파이 | 단위 | 월별12 | 연간합계 | 월최대 | 발생월
  const header = ['하우징 타입', '파이', '단위', ...MONTHS, '연간 합계', '월 최대', '최대 발생월']
  const data: (string | number | null)[][] = [header]

  const entries = [...agg.entries()].sort(
    (a, b) => HOUSING_KEY_ORDER.indexOf(a[0]) - HOUSING_KEY_ORDER.indexOf(b[0])
  )

  for (const [key, { pai, monthly }] of entries) {
    const type    = key.split('|')[1]
    const annual  = monthly.reduce((a, b) => a + b, 0)
    const peak    = Math.max(...monthly)
    const peakMon = monthly.indexOf(peak) + 1
    if (annual === 0) continue
    data.push([type, pai, 'EA', ...monthly.map(v => v || null), annual || null, peak || null, `${peakMon}월`])
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), `${YY}년_하우징_집계`)
}

// ── 코드매핑 시트 ────────────────────────────────────────────
function writeCodeMapSheet(
  wb: XLSX.WorkBook,
  entries: { code: string; kind: string; pai: string; core: number; length: number }[],
) {
  const header = ['품목코드', '케이블종류', '파이', '코어수', '길이(m)', '집계키']
  const data: (string | number | null)[][] = [header]

  for (const { code, kind, pai, core, length } of entries) {
    const label = kindLabel(kind, core)
    const key   = label ? `${pai}|${label}` : null
    data.push([code, kind, pai, core, length, key])
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), '코드매핑')
}

// ── 요약 시트 ────────────────────────────────────────────────
function writeSummarySheet(
  wb: XLSX.WorkBook,
  allCableAggs: Map<string, ReturnType<typeof buildCableAgg>>,
  allHousingAggs: Map<string, Map<string, HousingAggEntry>>,
  years: string[],
) {
  const yrCols = years.flatMap(yr => [`20${yr}년 연간`, `20${yr}년 피크`])
  const hasGrowth = years.length >= 2
  const header = ['타입', '파이', '단위', ...yrCols, ...(hasGrowth ? ['증감률(최근2년)'] : [])]
  const data: (string | number | null)[][] = [header]

  const makeRow = (
    label: string, pai: string, unit: string,
    aggsByYear: (AggEntry | undefined)[],
  ): (string | number | null)[] => {
    const row: (string | number | null)[] = [label, pai, unit]
    let prevAnnual: number | null = null
    for (const entry of aggsByYear) {
      const annual = entry ? entry.monthly.reduce((a, b) => a + b, 0) : null
      const peak   = entry ? Math.max(...entry.monthly) : null
      row.push(annual || null, peak || null)
      prevAnnual = annual
    }
    if (hasGrowth) {
      const last    = aggsByYear[years.length - 1]
      const prev    = aggsByYear[years.length - 2]
      const lastAnn = last ? last.monthly.reduce((a, b) => a + b, 0) : 0
      const prevAnn = prev ? prev.monthly.reduce((a, b) => a + b, 0) : 0
      const growth  = prevAnn > 0 ? `${lastAnn >= prevAnn ? '+' : ''}${(((lastAnn - prevAnn) / prevAnn) * 100).toFixed(1)}%` : null
      row.push(growth)
    }
    return row
  }

  // ── 케이블 ──
  data.push(['── 케이블 ──', ...new Array(header.length - 1).fill(null)])

  const allCableKeys = new Set<string>()
  for (const agg of allCableAggs.values()) for (const k of agg.keys()) allCableKeys.add(k)
  const cableKeysSorted = [...allCableKeys].sort((a, b) => {
    const [pa] = a.split('|'); const [pb] = b.split('|')
    return (PAI_ORDER[pa] ?? 9) - (PAI_ORDER[pb] ?? 9) || a.localeCompare(b)
  })

  for (const key of cableKeysSorted) {
    const [pai, label] = key.split('|', 2)
    let unit = 'm'
    for (const agg of allCableAggs.values()) { const e = agg.get(key); if (e) { unit = e.unit; break } }
    const aggsByYear = years.map(yr => allCableAggs.get(yr)?.get(key))
    const row = makeRow(label, pai, unit, aggsByYear)
    if (row.slice(3).some(v => v != null)) data.push(row)
  }

  // ── 하우징 ──
  data.push(['── 하우징 ──', ...new Array(header.length - 1).fill(null)])

  for (const key of HOUSING_KEY_ORDER) {
    const [pai, type] = key.split('|', 2)
    const aggsByYear = years.map(yr => allHousingAggs.get(yr)?.get(key))
    const row = makeRow(type, pai, 'EA', aggsByYear)
    if (row.slice(3).some(v => v != null)) data.push(row)
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), '📊 요약')
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

  // 1단계: 전체 연도 데이터 수집 (요약 시트 생성을 위해 먼저 모은 후 쓴다)
  const allCableAggs   = new Map<string, ReturnType<typeof buildCableAgg>>()
  const allHousingAggs = new Map<string, Map<string, HousingAggEntry>>()
  const sortedYears:   string[] = []
  const yearRawSheets  = new Map<string, { cData: (string|number|null)[][]; hData: (string|number|null)[][] }>()
  const codeMapEntries: { code: string; kind: string; pai: string; core: number; length: number }[] = []
  const seenCodes = new Set<string>()

  for (const [year, rows] of [...pivot.entries()].sort()) {
    const YY    = year.slice(-2)
    sortedYears.push(YY)
    const cData: (string | number | null)[][] = [cableHeader(YY)]
    const hData: (string | number | null)[][] = [housingHeader(YY)]

    for (const row of rows) {
      const total     = row.monthly.reduce((a, b) => a + b, 0)
      const peak      = Math.max(...row.monthly)
      const len       = row.length ?? 0
      const usage     = Math.round(total * len)
      const peakUsage = Math.round(peak * len)
      const base      = [row.code, row.name, row.spec, row.kind, row.pai, row.core, row.length ?? null, row.type1 || null, row.type2 || null]

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

      // 코드매핑: 길이가 있는 케이블만 수집 (중복 제외)
      if (row.length != null && row.length > 0 && !seenCodes.has(row.code)) {
        seenCodes.add(row.code)
        codeMapEntries.push({ code: row.code, kind: row.kind, pai: row.pai, core: row.core, length: row.length })
      }
    }

    const cableAgg   = buildCableAgg(rows)
    const housingAgg = buildHousingAgg(rows)
    allCableAggs.set(YY, cableAgg)
    allHousingAggs.set(YY, housingAgg)
    yearRawSheets.set(YY, { cData, hData })
  }

  // 2단계: 요약·코드매핑 시트를 먼저 (탭 앞쪽에 위치)
  writeSummarySheet(wb, allCableAggs, allHousingAggs, sortedYears)
  writeCodeMapSheet(wb, codeMapEntries)

  // 3단계: 연도별 집계 + 원본 시트
  for (const YY of sortedYears) {
    writeCableAggSheet(wb,   allCableAggs.get(YY)!,   YY)
    writeHousingAggSheet(wb, allHousingAggs.get(YY)!, YY)
    const { cData, hData } = yearRawSheets.get(YY)!
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cData), `${YY}년_케이블`)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(hData), `${YY}년 하우징`)
  }

  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as number[]
  return new Uint8Array(out).buffer as ArrayBuffer
}
