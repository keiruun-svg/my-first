import ExcelJS from 'exceljs'
import type { Step3Row } from '../step3Core'
import type { Metadata, Inventory } from '../types'

// ── 스타일 상수 ───────────────────────────────────────────────
const BORDER_THIN: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FF000000' } }
const ALL_BORDERS = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN }
const BORDER_MEDIUM = (color = 'C55A11'): Partial<ExcelJS.Border> => ({ style: 'medium', color: { argb: 'FF' + color } })
const INPUT_BORDERS = () => ({
  top: BORDER_MEDIUM(), bottom: BORDER_MEDIUM(),
  left: BORDER_MEDIUM(), right: BORDER_MEDIUM(),
})

function fill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + argb } }
}
function font(opts: { bold?: boolean; size?: number; color?: string } = {}): Partial<ExcelJS.Font> {
  return { name: 'Arial', bold: opts.bold ?? false, size: opts.size ?? 9, color: { argb: 'FF' + (opts.color ?? '000000') } }
}

// 1-based 열 번호 → Excel 열 문자
function cl(n: number): string {
  let s = ''
  while (n > 0) { s = String.fromCharCode(64 + ((n - 1) % 26 + 1)) + s; n = Math.floor((n - 1) / 26) }
  return s
}

// ── 피그테일 색상 라벨 표시 변환 ──────────────────────────────
const DISPLAY_LABEL_MAP: Record<string, string> = { '연청': '청록', '연등': '분홍' }
function displayLabel(label: string): string {
  // "pigtail-연청" → "pigtail-청록"
  return label.replace(/^(pigtail-)(.+)$/, (_, pfx, color) => pfx + (DISPLAY_LABEL_MAP[color] ?? color))
}

// ── 열 레이아웃 계산 ──────────────────────────────────────────
function makeLayout(nYears: number) {
  const C_NO   = 1, C_PAI = 2, C_TYPE = 3, C_PN = 4, C_NAME = 5, C_VENDOR = 6, C_LT = 7
  const C_ANN_FIRST = 8
  // annual(i) = C_ANN_FIRST + 2*i,  peak(i) = C_ANN_FIRST + 2*i + 1
  const C_AVG      = C_ANN_FIRST + 2 * nYears       // N개년 평균연간
  const C_PEAK_AVG = C_AVG + 1                       // N개년 피크평균
  // growth rates: C_PEAK_AVG+1 .. C_PEAK_AVG+(nYears-1)
  const C_SS    = C_PEAK_AVG + nYears                // 안전재고
  const C_CUR   = C_SS  + 1                          // 현재고
  const C_ORD   = C_CUR + 1                          // 기발주
  const C_TGT   = C_ORD + 1                          // 2026목표 (입력)
  const C_REQ   = C_TGT + 1                          // 필요발주
  const C_NOTE  = C_REQ + 1                          // 비고
  const TOTAL   = C_NOTE
  return { C_NO, C_PAI, C_TYPE, C_PN, C_NAME, C_VENDOR, C_LT, C_ANN_FIRST, C_AVG, C_PEAK_AVG, C_SS, C_CUR, C_ORD, C_TGT, C_REQ, C_NOTE, TOTAL }
}

// ── 메인 시트 (케이블 사용내역 / 하우징 사용내역) ─────────────
function writeMainSheet(
  wb: ExcelJS.Workbook,
  sheetName: string,
  rows: Step3Row[],
  years: string[],
  mainColor: string,
  unit: string,
  typeLabel: string,
  safetyK: number,
) {
  if (!rows.length) return
  const ws = wb.addWorksheet(sheetName)
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]

  const nYears = years.length
  const L = makeLayout(nYears)
  const YEAR_COLORS = ['2F5597', '2E75B6', '155480', '375623', '7030A0']

  const setCell = (r: number, c: number, val: ExcelJS.CellValue) => {
    const cell = ws.getCell(r, c)
    cell.value = val
    cell.font = font()
    cell.border = ALL_BORDERS
    return cell
  }
  const hdrCell = (r: number, c: number, val: string, bg = 'BDD7EE', fg = '1F3864') => {
    const cell = ws.getCell(r, c)
    cell.value = val
    cell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FF' + fg } }
    cell.fill = fill(bg)
    cell.border = ALL_BORDERS
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    return cell
  }
  const band = (r: number, c1: number, c2: number, label: string, bg: string, fg = 'FFFFFF') => {
    if (c1 !== c2) ws.mergeCells(r, c1, r, c2)
    const cell = ws.getCell(r, c1)
    cell.value = label
    cell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FF' + fg } }
    cell.fill = fill(bg)
    cell.border = ALL_BORDERS
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
  }

  // ── 행 1: 타이틀 ─────────────────────────────────────────
  ws.mergeCells(1, 1, 1, L.TOTAL)
  const title = ws.getCell(1, 1)
  title.value = sheetName
  title.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } }
  title.fill = fill(mainColor)
  title.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 26

  // ── 행 2: 구역 밴드 ──────────────────────────────────────
  ws.getRow(2).height = 18
  band(2, L.C_NO, L.C_LT, '기본 정보', '374151')
  for (let i = 0; i < nYears; i++) {
    const c1 = L.C_ANN_FIRST + 2 * i
    band(2, c1, c1 + 1, `20${years[i]}년`, YEAR_COLORS[i % YEAR_COLORS.length])
  }
  band(2, L.C_AVG, L.C_PEAK_AVG + (nYears - 1), '📊 트렌드 분석', '375623')
  band(2, L.C_SS, L.C_SS, '⚠ 안전재고', 'C00000')
  band(2, L.C_CUR, L.C_ORD, '재고 현황', '7030A0')
  band(2, L.C_TGT, L.C_REQ, '✏ 발주 계획', 'C55A11')
  band(2, L.C_NOTE, L.C_NOTE, '비고', '595959')

  // ── 행 3: 헤더 ───────────────────────────────────────────
  ws.getRow(3).height = 42
  const COL_WIDTHS: number[] = [5, 8, 20, 16, 38, 12, 10]
  hdrCell(3, L.C_NO, 'NO'); hdrCell(3, L.C_PAI, '파이'); hdrCell(3, L.C_TYPE, typeLabel)
  hdrCell(3, L.C_PN, '품번'); hdrCell(3, L.C_NAME, '품명'); hdrCell(3, L.C_VENDOR, '구매처')
  hdrCell(3, L.C_LT, `리드타임\n(일)`)
  ws.getColumn(1).width = 5; ws.getColumn(2).width = 8; ws.getColumn(3).width = 20
  ws.getColumn(4).width = 16; ws.getColumn(5).width = 38; ws.getColumn(6).width = 12; ws.getColumn(7).width = 10

  for (let i = 0; i < nYears; i++) {
    const c = L.C_ANN_FIRST + 2 * i
    hdrCell(3, c,     `연간(${unit})\n${years[i]}년`); ws.getColumn(c).width = 12
    hdrCell(3, c + 1, `피크(${unit})\n${years[i]}년`); ws.getColumn(c + 1).width = 12
  }
  hdrCell(3, L.C_AVG,      `${nYears}개년\n평균연간`); ws.getColumn(L.C_AVG).width = 12
  hdrCell(3, L.C_PEAK_AVG, `${nYears}개년\n피크평균`); ws.getColumn(L.C_PEAK_AVG).width = 12
  for (let i = 0; i < nYears - 1; i++) {
    const c = L.C_PEAK_AVG + 1 + i
    const y1 = years[i]; const y2 = years[i + 1]
    hdrCell(3, c, `${y1}→${y2}\n증감률`); ws.getColumn(c).width = 10
  }
  hdrCell(3, L.C_SS,   `안전재고\n(${unit})`); ws.getColumn(L.C_SS).width = 12
  hdrCell(3, L.C_CUR,  `현재고\n(${unit})`);   ws.getColumn(L.C_CUR).width = 10
  hdrCell(3, L.C_ORD,  `기발주\n(참고)`);       ws.getColumn(L.C_ORD).width = 10
  hdrCell(3, L.C_TGT,  `목표수량\n(${unit})`);  ws.getColumn(L.C_TGT).width = 13
  hdrCell(3, L.C_REQ,  `필요발주\n(${unit})`);  ws.getColumn(L.C_REQ).width = 13
  hdrCell(3, L.C_NOTE, '비고');                   ws.getColumn(L.C_NOTE).width = 20

  // ── 데이터 행 ─────────────────────────────────────────────
  let no = 1
  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx]
    const ri  = idx + 4
    const evenFill: ExcelJS.Fill | undefined = idx % 2 === 0 ? undefined : fill('F5F5F5')

    const sc = (c: number, val: ExcelJS.CellValue, right = false, numFmt?: string) => {
      const cell = setCell(ri, c, val)
      if (evenFill) cell.fill = evenFill
      cell.alignment = { horizontal: right ? 'right' : (c <= 2 ? 'center' : 'left'), vertical: 'middle' }
      if (numFmt) cell.numFmt = numFmt
    }

    sc(L.C_NO, no++)
    sc(L.C_PAI, row.pai)
    sc(L.C_TYPE, displayLabel(row.label))
    sc(L.C_PN, row.품번 || '(미등록)')
    sc(L.C_NAME, row.품명 || '')
    sc(L.C_VENDOR, row.구매처 || '')
    sc(L.C_LT, row.리드타임, true, '#,##0')

    // 연도별 연간+피크
    for (let i = 0; i < nYears; i++) {
      const yr  = years[i]
      const ann = row.byYear[yr]?.annual || null
      const pk  = row.byYear[yr]?.peak   || null
      sc(L.C_ANN_FIRST + 2 * i,     ann, true, '#,##0')
      sc(L.C_ANN_FIRST + 2 * i + 1, pk,  true, '#,##0')
    }

    // 연간 평균 — 공식 (연간 합계 평균)
    {
      const annCols = Array.from({ length: nYears }, (_, i) => `${cl(L.C_ANN_FIRST + 2 * i)}${ri}`).join(',')
      const pkCols  = Array.from({ length: nYears }, (_, i) => `${cl(L.C_ANN_FIRST + 2 * i + 1)}${ri}`).join(',')
      const avgAnn = setCell(ri, L.C_AVG, { formula: `ROUND(AVERAGE(${annCols}),0)` })
      avgAnn.numFmt = '#,##0'; avgAnn.alignment = { horizontal: 'right', vertical: 'middle' }
      if (evenFill) avgAnn.fill = evenFill
      const avgPk = setCell(ri, L.C_PEAK_AVG, { formula: `ROUND(AVERAGE(${pkCols}),0)` })
      avgPk.numFmt = '#,##0'; avgPk.alignment = { horizontal: 'right', vertical: 'middle' }
      if (evenFill) avgPk.fill = evenFill
    }

    // 증감률 (연도 쌍별)
    for (let i = 0; i < nYears - 1; i++) {
      const c    = L.C_PEAK_AVG + 1 + i
      const cAnn = cl(L.C_ANN_FIRST + 2 * i); const cAnn2 = cl(L.C_ANN_FIRST + 2 * (i + 1))
      const gr = setCell(ri, c, { formula: `IFERROR((${cAnn2}${ri}-${cAnn}${ri})/${cAnn}${ri},"")` })
      gr.numFmt = '0.0%;[Red]-0.0%'; gr.alignment = { horizontal: 'center', vertical: 'middle' }
      if (evenFill) gr.fill = evenFill
    }

    // 안전재고 = ROUND(피크평균 × LT/30, 0)
    {
      const ssCell = ws.getCell(ri, L.C_SS)
      ssCell.value = { formula: `ROUND(${cl(L.C_PEAK_AVG)}${ri}*${cl(L.C_LT)}${ri}/30,0)` }
      ssCell.font  = { name: 'Arial', bold: true, size: 9 }
      ssCell.fill  = fill('FFF2CC')
      ssCell.border = ALL_BORDERS
      ssCell.numFmt = '#,##0'
      ssCell.alignment = { horizontal: 'right', vertical: 'middle' }
    }

    sc(L.C_CUR, row.현재고 || null, true, '#,##0')
    sc(L.C_ORD, row.기발주 || null, true, '#,##0')

    // 2026목표 — 노란 입력셀 (비워둠)
    {
      const tgtCell = ws.getCell(ri, L.C_TGT)
      tgtCell.value = null
      tgtCell.font  = { name: 'Arial', bold: true, size: 9, color: { argb: 'FF0000FF' } }
      tgtCell.fill  = fill('FFFFC0')
      tgtCell.border = INPUT_BORDERS()
      tgtCell.numFmt = '#,##0'
      tgtCell.alignment = { horizontal: 'right', vertical: 'middle' }
    }

    // 필요발주 = IFERROR(목표 - 현재고 - 기발주, "")
    {
      const tgt = cl(L.C_TGT); const cur = cl(L.C_CUR); const ord = cl(L.C_ORD)
      const reqCell = ws.getCell(ri, L.C_REQ)
      reqCell.value = { formula: `IFERROR(${tgt}${ri}-IFERROR(${cur}${ri},0)-IFERROR(${ord}${ri},0),"")` }
      reqCell.font  = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFC00000' } }
      reqCell.border = ALL_BORDERS
      reqCell.numFmt = '#,##0'
      reqCell.alignment = { horizontal: 'right', vertical: 'middle' }
      if (evenFill) reqCell.fill = evenFill
    }

    // 비고
    const noteCell = setCell(ri, L.C_NOTE, '')
    if (evenFill) noteCell.fill = evenFill
    noteCell.alignment = { horizontal: 'left', vertical: 'middle' }
    ws.getRow(ri).height = 17
  }

  // ── 합계 행 ──────────────────────────────────────────────
  const lastDataRow = rows.length + 3
  const sumRow = lastDataRow + 2
  ws.getRow(sumRow).height = 18
  const sumLabelCell = ws.getCell(sumRow, L.C_TYPE)
  sumLabelCell.value = '합  계'
  sumLabelCell.font  = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
  sumLabelCell.fill  = fill(mainColor)
  sumLabelCell.border = ALL_BORDERS
  sumLabelCell.alignment = { horizontal: 'center', vertical: 'middle' }

  for (let i = 0; i < nYears; i++) {
    for (let d = 0; d < 2; d++) {
      const c = L.C_ANN_FIRST + 2 * i + d
      const sc2 = ws.getCell(sumRow, c)
      sc2.value = { formula: `SUM(${cl(c)}4:${cl(c)}${lastDataRow})` }
      sc2.font  = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
      sc2.fill  = fill(mainColor)
      sc2.border = ALL_BORDERS
      sc2.numFmt = '#,##0'
      sc2.alignment = { horizontal: 'right', vertical: 'middle' }
    }
  }
  for (const c of [L.C_AVG, L.C_PEAK_AVG]) {
    const sc2 = ws.getCell(sumRow, c)
    sc2.value = { formula: `SUM(${cl(c)}4:${cl(c)}${lastDataRow})` }
    sc2.font  = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
    sc2.fill  = fill(mainColor)
    sc2.border = ALL_BORDERS
    sc2.numFmt = '#,##0'
    sc2.alignment = { horizontal: 'right', vertical: 'middle' }
  }
}

const CONN_TYPES = ['LC/PC', 'LC/APC', 'SC/PC', 'SC/APC', 'FC/PC', 'FC/APC'] as const

// ── 품번별 발주 집계 시트 (하우징 + 페롤 섹션) ────────────────
function writeBunhoSheet(
  wb: ExcelJS.Workbook,
  housingRows: Step3Row[],
  years: string[],
  mainColor: string,
  ferruleMeta: Metadata['ferrule'],
  ferruleInv:  Inventory['ferrule'],
) {
  const ws = wb.addWorksheet('📦 품번별 발주 집계')
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]
  const nY = years.length

  // 품번별 합산
  type BunhoEntry = {
    품명: string; 구매처: string; 리드타임: number
    현재고: number; 기발주: number
    annByYear: number[]  // nY 항목
  }
  const byBn = new Map<string, BunhoEntry>()

  for (const row of housingRows) {
    const bn = (row.품번 || '').trim()
    if (!bn) continue
    if (!byBn.has(bn)) {
      byBn.set(bn, {
        품명: row.품명, 구매처: row.구매처, 리드타임: row.리드타임,
        현재고: 0, 기발주: 0, annByYear: new Array(nY).fill(0),
      })
    }
    const e = byBn.get(bn)!
    e.현재고 += row.현재고
    e.기발주 += row.기발주
    for (let i = 0; i < nY; i++) {
      e.annByYear[i] += row.byYear[years[i]]?.annual ?? 0
    }
  }

  const C_NO = 1, C_PN = 2, C_NAME = 3, C_VENDOR = 4, C_LT = 5
  const C_ANN_FIRST = 6
  const C_AVG  = C_ANN_FIRST + nY
  const C_SS   = C_AVG  + 1
  const C_CUR  = C_SS   + 1
  const C_ORD  = C_CUR  + 1
  const C_TGT  = C_ORD  + 1
  const C_REQ  = C_TGT  + 1
  const C_NOTE = C_REQ  + 1
  const TOTAL  = C_NOTE

  const hdr = (r: number, c: number, v: string, w?: number) => {
    const cell = ws.getCell(r, c)
    cell.value = v; cell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
    cell.fill = fill('BDD7EE'); cell.border = ALL_BORDERS
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    if (w) ws.getColumn(c).width = w
  }
  const band2 = (r: number, c1: number, c2: number, label: string, bg: string) => {
    if (c1 !== c2) ws.mergeCells(r, c1, r, c2)
    const cell = ws.getCell(r, c1)
    cell.value = label; cell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
    cell.fill = fill(bg); cell.border = ALL_BORDERS
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
  }

  // 타이틀
  ws.mergeCells(1, 1, 1, TOTAL)
  const titleCell = ws.getCell(1, 1)
  titleCell.value = '2026 연간 발주 계획 — 품번별 집계 (하우징 공용 부품 합산)'
  titleCell.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } }
  titleCell.fill = fill(mainColor)
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 26

  // 구역 밴드
  ws.getRow(2).height = 18
  const YEAR_COLORS = ['2F5597', '2E75B6', '155480', '375623', '7030A0']
  band2(2, C_NO, C_LT, '기본 정보', '374151')
  for (let i = 0; i < nY; i++)
    band2(2, C_ANN_FIRST + i, C_ANN_FIRST + i, `20${years[i]}년 연간`, YEAR_COLORS[i % YEAR_COLORS.length])
  band2(2, C_AVG,  C_AVG,  `${nY}개년 평균`, '375623')
  band2(2, C_SS,   C_SS,   '안전재고',       'C00000')
  band2(2, C_CUR,  C_ORD,  '재고 현황',      '7030A0')
  band2(2, C_TGT,  C_REQ,  '✏ 발주 계획',   'C55A11')
  band2(2, C_NOTE, C_NOTE, '비고',           '595959')

  // 헤더
  ws.getRow(3).height = 42
  hdr(3, C_NO, 'NO', 5); hdr(3, C_PN, '품번', 16); hdr(3, C_NAME, '품명', 38)
  hdr(3, C_VENDOR, '구매처', 14); hdr(3, C_LT, `리드타임\n(일)`, 10)
  for (let i = 0; i < nY; i++) hdr(3, C_ANN_FIRST + i, `20${years[i]}년\n연간(EA)`, 12)
  hdr(3, C_AVG, `${nY}개년\n평균(EA)`, 12); hdr(3, C_SS, `안전재고\n(EA)`, 12)
  hdr(3, C_CUR, `현재고\n(EA)`, 10); hdr(3, C_ORD, `기발주\n(참고)`, 10)
  hdr(3, C_TGT, `목표수량\n(EA)`, 12); hdr(3, C_REQ, `필요발주\n(EA)`, 12)
  hdr(3, C_NOTE, '비고', 20)

  let ri = 4; let no = 1
  for (const [bn, d] of [...byBn.entries()].sort()) {
    const evenFill2: ExcelJS.Fill | undefined = (ri % 2 === 0) ? fill('F5F5F5') : undefined

    const sc = (c: number, val: ExcelJS.CellValue, right = false, numFmt?: string) => {
      const cell = ws.getCell(ri, c)
      cell.value = val; cell.font = font(); cell.border = ALL_BORDERS
      if (evenFill2) cell.fill = evenFill2
      cell.alignment = { horizontal: right ? 'right' : (c <= 1 || c === 5 ? 'center' : 'left'), vertical: 'middle' }
      if (numFmt) cell.numFmt = numFmt
    }

    sc(C_NO, no++); sc(C_PN, bn); sc(C_NAME, d.품명); sc(C_VENDOR, d.구매처); sc(C_LT, d.리드타임, true, '#,##0')
    for (let i = 0; i < nY; i++) sc(C_ANN_FIRST + i, d.annByYear[i] || null, true, '#,##0')

    // N개년 평균
    const annCols = Array.from({ length: nY }, (_, i) => `${cl(C_ANN_FIRST + i)}${ri}`).join(',')
    const avgCell = ws.getCell(ri, C_AVG)
    avgCell.value = { formula: `ROUND(AVERAGE(${annCols}),0)` }
    avgCell.font = font(); avgCell.border = ALL_BORDERS; avgCell.numFmt = '#,##0'
    avgCell.alignment = { horizontal: 'right', vertical: 'middle' }
    if (evenFill2) avgCell.fill = evenFill2

    // 안전재고 = ROUND(평균 × LT/30, 0)
    const ssCell2 = ws.getCell(ri, C_SS)
    ssCell2.value = { formula: `ROUND(${cl(C_AVG)}${ri}*${cl(C_LT)}${ri}/30,0)` }
    ssCell2.font  = font({ bold: true }); ssCell2.border = ALL_BORDERS; ssCell2.numFmt = '#,##0'
    ssCell2.fill  = fill('FFF2CC'); ssCell2.alignment = { horizontal: 'right', vertical: 'middle' }

    sc(C_CUR, d.현재고 || null, true, '#,##0'); sc(C_ORD, d.기발주 || null, true, '#,##0')

    // 2026목표 — 입력셀
    const tgtCell2 = ws.getCell(ri, C_TGT)
    tgtCell2.value = null; tgtCell2.font = font({ bold: true, color: '0000FF' })
    tgtCell2.fill = fill('FFFFC0'); tgtCell2.border = INPUT_BORDERS()
    tgtCell2.numFmt = '#,##0'; tgtCell2.alignment = { horizontal: 'right', vertical: 'middle' }

    // 필요발주 = MAX(목표 - 현재고 - 기발주 + 안전재고, 0)
    const reqCell2 = ws.getCell(ri, C_REQ)
    reqCell2.value = {
      formula: `IFERROR(MAX(${cl(C_TGT)}${ri}-${cl(C_CUR)}${ri}-${cl(C_ORD)}${ri}+${cl(C_SS)}${ri},0),"")`,
    }
    reqCell2.font = font({ bold: true, color: 'C00000' }); reqCell2.border = ALL_BORDERS
    reqCell2.numFmt = '#,##0'; reqCell2.alignment = { horizontal: 'right', vertical: 'middle' }
    if (evenFill2) reqCell2.fill = evenFill2

    const noteCell2 = ws.getCell(ri, C_NOTE)
    noteCell2.value = ''; noteCell2.font = font(); noteCell2.border = ALL_BORDERS
    if (evenFill2) noteCell2.fill = evenFill2

    ws.getRow(ri).height = 17; ri++
  }

  // ── 페롤 섹션 ─────────────────────────────────────────────
  // 커넥터 타입별 하우징 수량 합산 (파이 무관)
  const connAnn = new Map<string, number[]>()  // ct → annByYear[nY]
  for (const row of housingRows) {
    const type = row.key.split('|')[1] ?? ''
    const m = type.match(/^(LC\/PC|LC\/APC|SC\/PC|SC\/APC|FC\/PC|FC\/APC)/)
    if (!m) continue
    const ct = m[1]
    if (!connAnn.has(ct)) connAnn.set(ct, new Array(nY).fill(0))
    const arr = connAnn.get(ct)!
    for (let i = 0; i < nY; i++) arr[i] += row.byYear[years[i]]?.annual ?? 0
  }

  // 페롤 섹션이 있을 때만 구분선 + 데이터 행 추가
  const ferruleEntries = CONN_TYPES.filter(ct => ferruleMeta[ct]?.품번 && connAnn.has(ct))
  if (ferruleEntries.length > 0) {
    // 구분선
    ri++
    ws.mergeCells(ri, 1, ri, TOTAL)
    const sepCell = ws.getCell(ri, 1)
    sepCell.value = '▼ 페롤 (커넥터 타입별 공용 — 파이 무관 합산)'
    sepCell.font  = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
    sepCell.fill  = fill('375623'); sepCell.border = ALL_BORDERS
    sepCell.alignment = { horizontal: 'center', vertical: 'middle' }
    ws.getRow(ri).height = 16; ri++

    for (const ct of ferruleEntries) {
      const fm  = ferruleMeta[ct]!
      const fiv = ferruleInv[ct] ?? { 현재고: 0, 기발주: 0 }
      const ann = connAnn.get(ct)!
      const lt  = (() => { const n = parseInt(String(fm.리드타임 ?? '')); return isNaN(n) ? 60 : n })()
      const evenFillF: ExcelJS.Fill = fill('F0FFF4')

      const scF = (c: number, val: ExcelJS.CellValue, right = false, numFmt?: string) => {
        const cell = ws.getCell(ri, c)
        cell.value = val; cell.font = font(); cell.border = ALL_BORDERS; cell.fill = evenFillF
        cell.alignment = { horizontal: right ? 'right' : (c <= 1 || c === C_LT ? 'center' : 'left'), vertical: 'middle' }
        if (numFmt) cell.numFmt = numFmt
      }

      scF(C_NO, no++); scF(C_PN, fm.품번); scF(C_NAME, fm.품명)
      scF(C_VENDOR, fm.구매처); scF(C_LT, lt, true, '#,##0')
      for (let i = 0; i < nY; i++) scF(C_ANN_FIRST + i, ann[i] || null, true, '#,##0')

      const annColsF = Array.from({ length: nY }, (_, i) => `${cl(C_ANN_FIRST + i)}${ri}`).join(',')
      const avgCellF = ws.getCell(ri, C_AVG)
      avgCellF.value = { formula: `ROUND(AVERAGE(${annColsF}),0)` }
      avgCellF.font = font(); avgCellF.border = ALL_BORDERS; avgCellF.numFmt = '#,##0'
      avgCellF.fill = evenFillF; avgCellF.alignment = { horizontal: 'right', vertical: 'middle' }

      const ssCellF = ws.getCell(ri, C_SS)
      ssCellF.value = { formula: `ROUND(${cl(C_AVG)}${ri}*${cl(C_LT)}${ri}/30,0)` }
      ssCellF.font = font({ bold: true }); ssCellF.border = ALL_BORDERS
      ssCellF.numFmt = '#,##0'; ssCellF.fill = fill('FFF2CC')
      ssCellF.alignment = { horizontal: 'right', vertical: 'middle' }

      scF(C_CUR, fiv.현재고 || null, true, '#,##0')
      scF(C_ORD, fiv.기발주 || null, true, '#,##0')

      const tgtCellF = ws.getCell(ri, C_TGT)
      tgtCellF.value = null; tgtCellF.font = font({ bold: true, color: '0000FF' })
      tgtCellF.fill = fill('FFFFC0'); tgtCellF.border = INPUT_BORDERS()
      tgtCellF.numFmt = '#,##0'; tgtCellF.alignment = { horizontal: 'right', vertical: 'middle' }

      const reqCellF = ws.getCell(ri, C_REQ)
      reqCellF.value = { formula: `IFERROR(MAX(${cl(C_TGT)}${ri}+${cl(C_SS)}${ri},0),"")` }
      reqCellF.font = font({ bold: true, color: 'C00000' }); reqCellF.border = ALL_BORDERS
      reqCellF.numFmt = '#,##0'; reqCellF.fill = evenFillF
      reqCellF.alignment = { horizontal: 'right', vertical: 'middle' }

      // 비고 — 커넥터 타입 표시
      const noteCellF = ws.getCell(ri, C_NOTE)
      noteCellF.value = ct; noteCellF.font = { name: 'Arial', size: 8, color: { argb: 'FF375623' } }
      noteCellF.border = ALL_BORDERS; noteCellF.fill = evenFillF
      noteCellF.alignment = { horizontal: 'left', vertical: 'middle' }

      ws.getRow(ri).height = 17; ri++
    }
  }
}

// ── 월별 발주계획 시트 ────────────────────────────────────────
function writeMonthlySheet(
  wb: ExcelJS.Workbook,
  allRows: Step3Row[],
  years: string[],
  mainColor: string,
) {
  const ws = wb.addWorksheet('2026 월별 발주계획')
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]

  // 타이틀
  ws.mergeCells(1, 1, 1, 20)
  const titleCell = ws.getCell(1, 1)
  titleCell.value = '2026 월별 발주 계획 (과거 계절 패턴 기반 자동 분배)'
  titleCell.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } }
  titleCell.fill = fill(mainColor)
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 26

  // 구역 밴드
  ws.getRow(2).height = 18
  const band3 = (c1: number, c2: number, label: string, bg: string) => {
    if (c1 !== c2) ws.mergeCells(2, c1, 2, c2)
    const cell = ws.getCell(2, c1)
    cell.value = label; cell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
    cell.fill = fill(bg); cell.border = ALL_BORDERS
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
  }
  band3(1, 6,  '기본정보',                         '374151')
  band3(7, 7,  '연간목표',                          'C55A11')
  band3(8, 19, '월별 발주 (연간목표 × 과거 패턴)', '2E75B6')
  band3(20, 20, '합계검증',                         '375623')

  // 헤더
  ws.getRow(3).height = 40
  const MONTH_HDRS = ['NO','분류','파이','종류','품번','단위','연간목표', ...Array.from({length:12},(_,i)=>`${i+1}월`), '합계검증']
  const MONTH_WIDTHS = [5, 10, 8, 22, 16, 6, 14, ...Array(12).fill(9), 10]
  for (const [idx, h] of MONTH_HDRS.entries()) {
    const ci = idx + 1
    const cell = ws.getCell(3, ci)
    cell.value = h; cell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
    cell.fill = fill('BDD7EE'); cell.border = ALL_BORDERS
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    ws.getColumn(ci).width = MONTH_WIDTHS[idx]
  }

  let ri = 4
  for (const [idx, row] of allRows.entries()) {
    const cat    = row.type === 'cable' ? '케이블' : '하우징'
    const unit   = row.unit
    const evenFill3: ExcelJS.Fill | undefined = idx % 2 === 0 ? undefined : fill('F5F5F5')

    // 계절 패턴 계산: 전체 연도 월별 합산
    const combined = new Array(12).fill(0)
    let total = 0
    for (const yr of years) {
      const m = row.byYear[yr]?.monthly ?? []
      for (let i = 0; i < 12; i++) { combined[i] += m[i] ?? 0; total += m[i] ?? 0 }
    }
    const ratios = combined.map(v => total > 0 ? Math.round((v / total) * 1e6) / 1e6 : Math.round((1 / 12) * 1e6) / 1e6)

    const sc = (c: number, val: ExcelJS.CellValue) => {
      const cell = ws.getCell(ri, c)
      cell.value = val; cell.font = font(); cell.border = ALL_BORDERS
      if (evenFill3) cell.fill = evenFill3
      cell.alignment = { horizontal: c <= 4 ? 'center' : 'left', vertical: 'middle' }
    }
    sc(1, idx + 1); sc(2, cat); sc(3, row.pai)
    sc(4, displayLabel(row.label)); sc(5, row.품번 || ''); sc(6, unit)

    // 연간목표 — 입력셀
    const tgtCell3 = ws.getCell(ri, 7)
    tgtCell3.value = null; tgtCell3.font = font({ bold: true, color: '0000FF' })
    tgtCell3.fill = fill('FFFFC0'); tgtCell3.border = INPUT_BORDERS()
    tgtCell3.numFmt = '#,##0'; tgtCell3.alignment = { horizontal: 'right', vertical: 'middle' }

    // 월별 = IFERROR(ROUND($G{ri} × ratio, 0), "")
    for (let mi = 0; mi < 12; mi++) {
      const cell = ws.getCell(ri, 8 + mi)
      cell.value = { formula: `IFERROR(ROUND($G${ri}*${ratios[mi]},0),"")` }
      cell.font = font(); cell.border = ALL_BORDERS; cell.numFmt = '#,##0'
      cell.alignment = { horizontal: 'right', vertical: 'middle' }
      if (evenFill3) cell.fill = evenFill3
    }

    // 합계검증
    const sumCell3 = ws.getCell(ri, 20)
    sumCell3.value = { formula: `IFERROR(SUM(H${ri}:S${ri}),"")` }
    sumCell3.font = font({ bold: true, color: '375623' }); sumCell3.border = ALL_BORDERS
    sumCell3.numFmt = '#,##0'; sumCell3.alignment = { horizontal: 'right', vertical: 'middle' }
    if (evenFill3) sumCell3.fill = evenFill3

    ws.getRow(ri).height = 17; ri++
  }
}

// ── 이상항목 검토 시트 ────────────────────────────────────────
function writeAnomalySheet(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet('⚠ 이상항목 검토')
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]

  ws.mergeCells(1, 1, 1, 4)
  const titleCell = ws.getCell(1, 1)
  titleCell.value = '데이터 이상 항목 검토 (자동 분석)'
  titleCell.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } }
  titleCell.fill = fill('C00000')
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 26

  ws.getRow(3).height = 30
  for (const [ci, h, w] of [
    [1, '구분',        10], [2, '항목',         22],
    [3, '품번',        16], [4, '내용 및 조치 권고', 65],
  ] as [number, string, number][]) {
    const cell = ws.getCell(3, ci)
    cell.value = h; cell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
    cell.fill = fill('BDD7EE'); cell.border = ALL_BORDERS
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    ws.getColumn(ci).width = w
  }

  const ANOMALIES: [string, string, string, string, string][] = [
    ['주의', '2.0mm 자켓 피그테일', '(확인필요)',
      '참고파일에 품번 미등재. 사용 케이블 품번 확인 필요.', 'FFE6C8'],
    ['정보', 'OM4 피그테일 케이블', 'P14-RM-417K',
      '3개년 사용 없음. 재고 보유. 단종 검토 필요.', 'FFD7D7'],
    ['정보', 'OM3 피그테일 케이블', 'P14-RM-417H',
      '23년 648m → 24~25년 0m. 미사용 추세.', 'FFE6C8'],
    ['정보', '피그테일 전체', '(전 색상)',
      '23년 대비 25년 약 81% 급감. 2026 목표량 보수적 설정 권고.', 'D7E8FF'],
  ]

  for (const [ri2, [type, item, bn, desc, bg]] of ANOMALIES.entries()) {
    const row2 = ri2 + 4
    ws.getRow(row2).height = 36
    for (const [ci, val] of [[1, type], [2, item], [3, bn], [4, desc]] as [number, string][]) {
      const cell = ws.getCell(row2, ci)
      cell.value = val; cell.font = { name: 'Arial', size: 9, bold: ci === 1 }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bg } }
      cell.border = ALL_BORDERS
      cell.alignment = ci === 4
        ? { horizontal: 'left', vertical: 'middle', wrapText: true }
        : { horizontal: ci === 1 ? 'center' : 'left', vertical: 'middle' }
    }
  }
}

// ── 진입점 ────────────────────────────────────────────────────
export function buildStep3Workbook(
  rows:      Step3Row[],
  years:     string[],
  mainColor: string,
  safetyK:   number,
  metadata:  Metadata,
  inventory: Inventory,
): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'AJW 발주계획 시스템'

  const cableRows   = rows.filter(r => r.type === 'cable')
  const housingRows = rows.filter(r => r.type === 'housing')

  writeMainSheet(wb, '케이블 사용내역',  cableRows,   years, mainColor, 'm',  '케이블 종류', safetyK)
  writeMainSheet(wb, '하우징 사용내역',  housingRows, years, mainColor, 'EA', '하우징 타입', safetyK)
  writeBunhoSheet(wb, housingRows, years, mainColor, metadata.ferrule ?? {}, inventory.ferrule ?? {})
  writeMonthlySheet(wb, [...cableRows, ...housingRows], years, mainColor)
  writeAnomalySheet(wb)

  return wb
}
