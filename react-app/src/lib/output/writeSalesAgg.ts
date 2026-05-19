import ExcelJS from 'exceljs'
import type { SalesAggResult } from '../aggregate/salesAgg'

const KIND_LABEL: Record<string, string> = {
  'a1': 'A1 (SM)', 'a1-청': 'A1 청색', 'a1-녹': 'A1 녹색', 'a1-적': 'A1 적색', 'a1-자': 'A1 자색',
  'b3': 'B3 (SM)', 'om1': 'OM1 (MM)', 'om3': 'OM3 (MM)',
  'drop': 'DROP', 'pigtail': 'PIGTAIL', 'om1-pigtail': 'PIGTAIL (MM)',
  'a2': 'Optical Cable',
}

const HEADER_BG  = 'FF2E75B6'
const HEADER_FG  = 'FFFFFFFF'
const ALT_BG     = 'FFF0F4FA'

function styleHeader(row: ExcelJS.Row, numCols: number) {
  for (let c = 1; c <= numCols; c++) {
    const cell = row.getCell(c)
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } }
    cell.font = { bold: true, color: { argb: HEADER_FG }, size: 9 }
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    cell.border = {
      top: { style: 'thin' }, bottom: { style: 'thin' },
      left: { style: 'thin' }, right: { style: 'thin' },
    }
  }
  row.height = 28
}

function styleDataRow(row: ExcelJS.Row, numCols: number, isAlt: boolean) {
  for (let c = 1; c <= numCols; c++) {
    const cell = row.getCell(c)
    if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ALT_BG } }
    cell.font = { size: 9 }
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
      bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
      left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
      right: { style: 'thin', color: { argb: 'FFD0D0D0' } },
    }
  }
}

export async function writeSalesAgg(result: SalesAggResult): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook()
  const { byType, byProduct, salesCagr, productionCagr, years } = result
  const latest = years[years.length - 1]
  const pct = (v: number) => `${Math.round(v * 100)}%`
  const cagrStr = (v: number) => { const s = Math.round(v * 100); return s > 0 ? `+${s}%` : `${s}%` }

  // ── Sheet 1: 타입별_분석 ──────────────────────────────────────
  const ws1 = wb.addWorksheet('타입별_분석')
  const s1Cols = 2 + years.length * 3 + 2 + 1
  const hdr1 = ws1.addRow([
    'kind', '분류명',
    ...years.map(yr => `20${yr}년\n판매(EA)`),
    ...years.map(yr => `20${yr}년\n생산(EA)`),
    ...years.map(yr => `20${yr}년\n생산비중`),
    '판매\nCAGR', '생산\nCAGR',
    `20${latest}년\n생산비중`,
  ])
  styleHeader(hdr1, s1Cols)

  ws1.columns = [
    { width: 14 }, { width: 16 },
    ...years.flatMap(() => [{ width: 11 }, { width: 11 }, { width: 9 }]),
    { width: 9 }, { width: 9 }, { width: 9 },
  ]

  const kinds = Object.keys(byType).sort()
  kinds.forEach((kind, i) => {
    const byYear = byType[kind]
    const row = ws1.addRow([
      kind,
      KIND_LABEL[kind] ?? kind.toUpperCase(),
      ...years.map(yr => byYear[yr]?.sales      ?? 0),
      ...years.map(yr => byYear[yr]?.production ?? 0),
      ...years.map(yr => pct(byYear[yr]?.ratio  ?? 0)),
      cagrStr(salesCagr[kind]      ?? 0),
      cagrStr(productionCagr[kind] ?? 0),
      pct(byYear[latest]?.ratio ?? 0),
    ])
    styleDataRow(row, s1Cols, i % 2 === 1)
    // 수량 컬럼 숫자 포맷
    for (let c = 3; c <= 2 + years.length * 2; c++) {
      row.getCell(c).numFmt = '#,##0'
      row.getCell(c).alignment = { horizontal: 'right' }
    }
    row.getCell(1).font = { bold: true, size: 9 }
  })

  // ── Sheet 2: 품목별_상세 ──────────────────────────────────────
  const ws2 = wb.addWorksheet('품목별_상세')
  const s2Cols = 4 + years.length * 2
  const hdr2 = ws2.addRow([
    '품목코드', '품목명', 'kind', '분류명',
    ...years.map(yr => `20${yr}년\n판매(EA)`),
    ...years.map(yr => `20${yr}년\n생산(EA)`),
  ])
  styleHeader(hdr2, s2Cols)

  ws2.columns = [
    { width: 18 }, { width: 36 }, { width: 14 }, { width: 16 },
    ...years.flatMap(() => [{ width: 11 }, { width: 11 }]),
  ]

  byProduct.forEach((p, i) => {
    const row = ws2.addRow([
      p.code, p.name, p.kind, KIND_LABEL[p.kind] ?? p.kind,
      ...years.map(yr => p.byYear[yr]?.sales      ?? 0),
      ...years.map(yr => p.byYear[yr]?.production ?? 0),
    ])
    styleDataRow(row, s2Cols, i % 2 === 1)
    for (let c = 5; c <= s2Cols; c++) {
      row.getCell(c).numFmt = '#,##0'
      row.getCell(c).alignment = { horizontal: 'right' }
    }
  })

  const buf = await wb.xlsx.writeBuffer()
  return buf as ArrayBuffer
}
