import * as XLSX from 'xlsx'
import type { SalesAggResult } from '../aggregate/salesAgg'

const KIND_LABEL: Record<string, string> = {
  'a1': 'A1 (SM)', 'a1-청': 'A1 청색', 'a1-녹': 'A1 녹색', 'a1-적': 'A1 적색', 'a1-자': 'A1 자색',
  'b3': 'B3 (SM)', 'om1': 'OM1 (MM)', 'om3': 'OM3 (MM)',
  'drop': 'DROP', 'pigtail': 'PIGTAIL', 'om1-pigtail': 'PIGTAIL (MM)',
  'a2': 'Optical Cable',
}

export function writeSalesAgg(result: SalesAggResult): ArrayBuffer {
  const wb   = XLSX.utils.book_new()
  const { byType, byProduct, salesCagr, productionCagr, years } = result
  const latest = years[years.length - 1]

  // ── Sheet 1: 타입별_분석 (STEP 3에서 읽음) ─────────────────
  const typeHeader = [
    'kind',
    '분류명',
    ...years.map(yr => `${yr}년_판매(EA)`),
    ...years.map(yr => `${yr}년_생산(EA)`),
    ...years.map(yr => `${yr}년_생산비중`),
    '판매CAGR',
    '생산CAGR',
    `${latest}년_생산비중`,
  ]
  const typeRows: (string | number)[][] = [typeHeader]
  for (const [kind, byYear] of Object.entries(byType)) {
    typeRows.push([
      kind,
      KIND_LABEL[kind] ?? kind.toUpperCase(),
      ...years.map(yr => byYear[yr]?.sales      ?? 0),
      ...years.map(yr => byYear[yr]?.production ?? 0),
      ...years.map(yr => byYear[yr]?.ratio      ?? 0),
      salesCagr[kind]      ?? 0,
      productionCagr[kind] ?? 0,
      byYear[latest]?.ratio ?? 0,
    ])
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(typeRows), '타입별_분석')

  // ── Sheet 2: 품목별_상세 (참고용) ──────────────────────────
  const prodHeader = [
    '품목코드', '품목명', 'kind', '분류명',
    ...years.map(yr => `${yr}년_판매(EA)`),
    ...years.map(yr => `${yr}년_생산(EA)`),
  ]
  const prodRows: (string | number)[][] = [prodHeader]
  for (const p of byProduct) {
    prodRows.push([
      p.code, p.name, p.kind, KIND_LABEL[p.kind] ?? p.kind,
      ...years.map(yr => p.byYear[yr]?.sales      ?? 0),
      ...years.map(yr => p.byYear[yr]?.production ?? 0),
    ])
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(prodRows), '품목별_상세')

  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as number[]
  return new Uint8Array(out).buffer as ArrayBuffer
}
