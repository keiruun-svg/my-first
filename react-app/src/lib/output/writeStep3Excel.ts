import ExcelJS from 'exceljs'
import type { Step3Row, CodeCableEntry } from '../step3Core'
import type { Metadata, Inventory } from '../types'
import type { SalesAggResult } from '../aggregate/salesAgg'

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
const DISPLAY_LABEL_MAP: Record<string, string> = { '연청': 'AQUA', '연등': 'rose' }
function displayLabel(label: string): string {
  // "pigtail-연청" → "pigtail-AQUA"  (가공파일이 이미 매핑된 경우 pass-through)
  return label.replace(/^(pigtail-)(.+)$/, (_, pfx, color) => pfx + (DISPLAY_LABEL_MAP[color] ?? color))
}

// ── 수요기반 분석 ─────────────────────────────────────────────
interface DemandEntry {
  제안량:     number  // meters
  생산비중:   number  // 0-1
  판매트렌드: number  // growth rate
}

function computeDemand(
  salesAgg:    SalesAggResult,
  codeToCable: Record<string, CodeCableEntry>,
  years: string[],  // 2-digit step3 years e.g. ["23","24","25"]
): Map<string, DemandEntry> {
  if (!years.length) return new Map()
  const latestYr = '20' + years[years.length - 1]
  const prevYr   = years.length >= 2 ? '20' + years[years.length - 2] : null
  const prev2Yr  = years.length >= 3 ? '20' + years[years.length - 3] : null

  const acc = new Map<string, { suggestSum: number; salesSum: number; ratioWtd: number; trendWtd: number }>()

  for (const prod of salesAgg.byProduct) {
    const entry = codeToCable[prod.code]
    if (!entry || entry.length <= 0) continue
    const latest = prod.byYear[latestYr]
    if (!latest || latest.sales === 0) continue

    const s     = latest.sales
    const ratio = latest.production / s

    let trend = 0
    const sPrev2 = prev2Yr ? (prod.byYear[prev2Yr]?.sales ?? 0) : 0
    const sPrev  = prevYr  ? (prod.byYear[prevYr]?.sales  ?? 0) : 0
    if (sPrev2 > 0)     trend = Math.pow(s / sPrev2, 0.5) - 1
    else if (sPrev > 0) trend = s / sPrev - 1

    const suggest = s * (1 + trend) * ratio * entry.length
    const key = entry.key
    if (!acc.has(key)) acc.set(key, { suggestSum: 0, salesSum: 0, ratioWtd: 0, trendWtd: 0 })
    const a = acc.get(key)!
    a.suggestSum += suggest
    a.salesSum   += s
    a.ratioWtd   += s * ratio
    a.trendWtd   += s * trend
  }

  const result = new Map<string, DemandEntry>()
  for (const [key, a] of acc) {
    if (a.salesSum === 0) continue
    result.set(key, {
      제안량:     Math.round(a.suggestSum),
      생산비중:   a.ratioWtd / a.salesSum,
      판매트렌드: a.trendWtd / a.salesSum,
    })
  }
  return result
}

function riskLabel(ratio: number): string {
  if (ratio < 0.3) return '🔴 고위험'
  if (ratio < 0.7) return '🟠 주의'
  return '🟢 안전'
}

// ── 열 레이아웃 계산 ──────────────────────────────────────────
function makeLayout(nYears: number, demandCols = false) {
  const C_NO   = 1, C_PAI = 2, C_TYPE = 3, C_PN = 4, C_NAME = 5, C_VENDOR = 6, C_LT = 7
  const C_ANN_FIRST = 8
  // annual(i) = C_ANN_FIRST + 3*i,  peak(i) = C_ANN_FIRST + 3*i + 1,  peakMon(i) = C_ANN_FIRST + 3*i + 2
  const C_AVG      = C_ANN_FIRST + 3 * nYears       // N개년 평균연간
  const C_PEAK_AVG = C_AVG + 1                       // N개년 피크평균
  // growth rates: C_PEAK_AVG+1 .. C_PEAK_AVG+(nYears-1)
  const C_SS    = C_PEAK_AVG + nYears                // 안전재고
  const C_CUR   = C_SS  + 1                          // 현재고
  const C_ORD   = C_CUR + 1                          // 기발주
  const C_TGT   = C_ORD + 1                          // 2026목표 (입력)
  const C_REQ   = C_TGT + 1                          // 필요발주
  const C_NOTE  = C_REQ + 1                          // 비고
  // 수요기반 분석 컬럼 (케이블 시트, salesAgg 있을 때만)
  const C_SUGGEST  = demandCols ? C_NOTE + 1 : 0
  const C_VS_AVG   = demandCols ? C_NOTE + 2 : 0
  const C_RATIO    = demandCols ? C_NOTE + 3 : 0
  const C_TREND    = demandCols ? C_NOTE + 4 : 0
  const C_RISK     = demandCols ? C_NOTE + 5 : 0
  const TOTAL      = demandCols ? C_RISK : C_NOTE
  const annC = (i: number) => C_ANN_FIRST + 3 * i
  const pkC  = (i: number) => C_ANN_FIRST + 3 * i + 1
  const pmC  = (i: number) => C_ANN_FIRST + 3 * i + 2
  return { C_NO, C_PAI, C_TYPE, C_PN, C_NAME, C_VENDOR, C_LT, C_ANN_FIRST, C_AVG, C_PEAK_AVG, C_SS, C_CUR, C_ORD, C_TGT, C_REQ, C_NOTE, C_SUGGEST, C_VS_AVG, C_RATIO, C_TREND, C_RISK, TOTAL, annC, pkC, pmC }
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
  demandMap?: Map<string, DemandEntry>,
) {
  if (!rows.length) return
  const ws = wb.addWorksheet(sheetName)
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]

  const nYears = years.length
  const hasDemand = !!(demandMap && demandMap.size > 0)
  const L = makeLayout(nYears, hasDemand)
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

  // ── 안전재고 계수 k 설정 셀 (테이블 우측, 헤더 행에 배치) ──
  const C_K = L.TOTAL + 2
  ws.getColumn(C_K).width = 10

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
  // k 라벨
  const kLabelCell = ws.getCell(2, C_K)
  kLabelCell.value = '안전재고\n계수 k'
  kLabelCell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
  kLabelCell.fill = fill('C00000')
  kLabelCell.border = ALL_BORDERS
  kLabelCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
  for (let i = 0; i < nYears; i++) {
    band(2, L.annC(i), L.pmC(i), `20${years[i]}년`, YEAR_COLORS[i % YEAR_COLORS.length])
  }
  band(2, L.C_AVG, L.C_PEAK_AVG + (nYears - 1), '📊 트렌드 분석', '375623')
  band(2, L.C_SS, L.C_SS, '⚠ 안전재고', 'C00000')
  band(2, L.C_CUR, L.C_ORD, '재고 현황', '7030A0')
  band(2, L.C_TGT, L.C_REQ, '✏ 발주 계획', 'C55A11')
  band(2, L.C_NOTE, L.C_NOTE, '비고', '595959')
  if (hasDemand) band(2, L.C_SUGGEST, L.C_RISK, '🔍 수요기반 분석 (STEP 2 연동)', '1F3864')

  // ── 행 3: 헤더 ───────────────────────────────────────────
  ws.getRow(3).height = 42
hdrCell(3, L.C_NO, 'NO'); hdrCell(3, L.C_PAI, '파이'); hdrCell(3, L.C_TYPE, typeLabel)
  hdrCell(3, L.C_PN, '품번'); hdrCell(3, L.C_NAME, '품명'); hdrCell(3, L.C_VENDOR, '구매처')
  hdrCell(3, L.C_LT, `리드타임\n(일)`)
  ws.getColumn(1).width = 5; ws.getColumn(2).width = 8; ws.getColumn(3).width = 20
  ws.getColumn(4).width = 16; ws.getColumn(5).width = 38; ws.getColumn(6).width = 12; ws.getColumn(7).width = 10

  for (let i = 0; i < nYears; i++) {
    hdrCell(3, L.annC(i), `연간(${unit})\n${years[i]}년`); ws.getColumn(L.annC(i)).width = 12
    hdrCell(3, L.pkC(i),  `피크(${unit})\n${years[i]}년`); ws.getColumn(L.pkC(i)).width  = 12
    hdrCell(3, L.pmC(i),  `피크월\n${years[i]}년`);        ws.getColumn(L.pmC(i)).width  = 7
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
  // k 입력셀 (행 3)
  const kCell = ws.getCell(3, C_K)
  kCell.value = safetyK
  kCell.font = font({ bold: true, color: '0000FF' })
  kCell.fill = fill('FFFFC0')
  kCell.border = INPUT_BORDERS()
  kCell.numFmt = '0.0'
  kCell.alignment = { horizontal: 'center', vertical: 'middle' }
  const kRef = `$${cl(C_K)}$3`  // 절대 참조

  if (hasDemand) {
    hdrCell(3, L.C_SUGGEST, `수요기반\n제안량(${unit})`, 'BDD7EE', '1F3864'); ws.getColumn(L.C_SUGGEST).width = 14
    hdrCell(3, L.C_VS_AVG,  `vs\n3개년평균`,              'BDD7EE', '1F3864'); ws.getColumn(L.C_VS_AVG).width  = 11
    hdrCell(3, L.C_RATIO,   `생산비중\n(맥산)`,           'BDD7EE', '1F3864'); ws.getColumn(L.C_RATIO).width   = 10
    hdrCell(3, L.C_TREND,   `판매\n트렌드`,               'BDD7EE', '1F3864'); ws.getColumn(L.C_TREND).width   = 10
    hdrCell(3, L.C_RISK,    `수입의존\n위험도`,            'BDD7EE', '1F3864'); ws.getColumn(L.C_RISK).width    = 12
  }

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

    if (row.isSubRow) {
      sc(L.C_NO, null)
    } else {
      sc(L.C_NO, no++)
    }
    sc(L.C_PAI, row.pai)
    sc(L.C_TYPE, displayLabel(row.label))
    sc(L.C_PN, row.품번 || '(미등록)')
    sc(L.C_NAME, row.품명 || '')
    sc(L.C_VENDOR, row.구매처 || '')
    sc(L.C_LT, row.리드타임, true, '#,##0')

    // 연도별 연간+피크+피크월
    for (let i = 0; i < nYears; i++) {
      const yr      = years[i]
      const ann     = row.byYear[yr]?.annual || null
      const pk      = row.byYear[yr]?.peak   || null
      const monthly = row.byYear[yr]?.monthly ?? []
      const maxVal  = monthly.length ? Math.max(...monthly) : 0
      const peakMon = maxVal > 0 ? monthly.indexOf(maxVal) + 1 : null
      sc(L.annC(i), ann, true, '#,##0')
      sc(L.pkC(i),  pk,  true, '#,##0')
      const pmCell = setCell(ri, L.pmC(i), peakMon ? `${peakMon}월` : null)
      if (evenFill) pmCell.fill = evenFill
      pmCell.alignment = { horizontal: 'center', vertical: 'middle' }
    }

    // 연간 평균 — 공식 (연간 합계 평균)
    {
      const annCols = Array.from({ length: nYears }, (_, i) => `${cl(L.annC(i))}${ri}`).join(',')
      const pkCols  = Array.from({ length: nYears }, (_, i) => `${cl(L.pkC(i))}${ri}`).join(',')
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
      const cAnn = cl(L.annC(i)); const cAnn2 = cl(L.annC(i + 1))
      const gr = setCell(ri, c, { formula: `IFERROR((${cAnn2}${ri}-${cAnn}${ri})/${cAnn}${ri},"")` })
      gr.numFmt = '0.0%;[Red]-0.0%'; gr.alignment = { horizontal: 'center', vertical: 'middle' }
      if (evenFill) gr.fill = evenFill
    }

    // 안전재고 = ROUND(피크평균 × LT/30, 0)
    {
      const ssCell = ws.getCell(ri, L.C_SS)
      ssCell.value = { formula: `ROUND(${cl(L.C_PEAK_AVG)}${ri}*${cl(L.C_LT)}${ri}/30*${kRef},0)` }
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

    // 필요발주 = MAX(목표 + 안전재고 - 현재고 - 기발주, 0)
    {
      const tgt = cl(L.C_TGT); const cur = cl(L.C_CUR); const ord = cl(L.C_ORD); const ss = cl(L.C_SS)
      const reqCell = ws.getCell(ri, L.C_REQ)
      reqCell.value = { formula: `IFERROR(MAX(${tgt}${ri}+${ss}${ri}-IFERROR(${cur}${ri},0)-IFERROR(${ord}${ri},0),0),"")` }
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

    // 수요기반 분석 컬럼
    if (hasDemand && demandMap) {
      const d = demandMap.get(row.key)
      if (d) {
        // 제안량 (static)
        const sugCell = setCell(ri, L.C_SUGGEST, d.제안량 || null)
        sugCell.numFmt = '#,##0'; sugCell.alignment = { horizontal: 'right', vertical: 'middle' }
        sugCell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FF1F3864' } }
        if (evenFill) sugCell.fill = evenFill

        // vs3개년평균 (formula)
        const vsCell = setCell(ri, L.C_VS_AVG, { formula: `IFERROR(${cl(L.C_SUGGEST)}${ri}/${cl(L.C_AVG)}${ri}-1,"")` })
        vsCell.numFmt = '+0.0%;[Red]-0.0%'; vsCell.alignment = { horizontal: 'center', vertical: 'middle' }
        if (evenFill) vsCell.fill = evenFill

        // 생산비중 (static %)
        const ratCell = setCell(ri, L.C_RATIO, d.생산비중 || null)
        ratCell.numFmt = '0%'; ratCell.alignment = { horizontal: 'center', vertical: 'middle' }
        ratCell.font = font({ color: d.생산비중 < 0.3 ? 'C00000' : d.생산비중 < 0.7 ? 'C55A11' : '375623' })
        if (evenFill) ratCell.fill = evenFill

        // 판매트렌드 (static %)
        const trdCell = setCell(ri, L.C_TREND, d.판매트렌드 || null)
        trdCell.numFmt = '+0.0%;[Red]-0.0%'; trdCell.alignment = { horizontal: 'center', vertical: 'middle' }
        if (evenFill) trdCell.fill = evenFill

        // 수입의존위험도 (text)
        const rskCell = setCell(ri, L.C_RISK, riskLabel(d.생산비중))
        rskCell.alignment = { horizontal: 'center', vertical: 'middle' }
        if (evenFill) rskCell.fill = evenFill
      } else {
        for (const c of [L.C_SUGGEST, L.C_VS_AVG, L.C_RATIO, L.C_TREND, L.C_RISK]) {
          const cell = setCell(ri, c, null)
          cell.alignment = { horizontal: 'center', vertical: 'middle' }
          if (evenFill) cell.fill = evenFill
        }
      }
    }

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
    for (let d = 0; d < 2; d++) {  // annual + peak only (피크월 합산 제외)
      const c = L.annC(i) + d
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

// ── 품번별 발주 집계 시트 (케이블 + 하우징 + 페롤 섹션) ─────────
function writeBunhoSheet(
  wb: ExcelJS.Workbook,
  cableRows:   Step3Row[],
  housingRows: Step3Row[],
  years: string[],
  mainColor: string,
  ferruleMeta: Metadata['ferrule'],
  ferruleInv:  Inventory['ferrule'],
  safetyK: number,
) {
  const ws = wb.addWorksheet('📦 품번별 발주 집계')
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]
  const nY = years.length

  const C_NO = 1, C_PN = 2, C_NAME = 3, C_VENDOR = 4, C_LT = 5, C_UNIT = 6
  const C_ANN_FIRST = 7
  const C_AVG      = C_ANN_FIRST + nY
  const C_PEAK_AVG = C_AVG  + 1
  const C_SS       = C_PEAK_AVG + 1
  const C_CUR      = C_SS   + 1
  const C_ORD      = C_CUR  + 1
  const C_TGT      = C_ORD  + 1
  const C_REQ      = C_TGT  + 1
  const C_NOTE     = C_REQ  + 1
  const TOTAL      = C_NOTE
  const C_K        = TOTAL  + 2

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
  const sectionHeader = (label: string, bg: string) => {
    ws.mergeCells(ri, 1, ri, TOTAL)
    const cell = ws.getCell(ri, 1)
    cell.value = label
    cell.font  = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
    cell.fill  = fill(bg); cell.border = ALL_BORDERS
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
    ws.getRow(ri).height = 16; ri++
  }

  // 타이틀
  ws.mergeCells(1, 1, 1, TOTAL)
  const titleCell = ws.getCell(1, 1)
  titleCell.value = '2026 연간 발주 계획 — 품번별 집계'
  titleCell.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } }
  titleCell.fill = fill(mainColor)
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 26

  // 구역 밴드
  ws.getRow(2).height = 18
  const YEAR_COLORS = ['2F5597', '2E75B6', '155480', '375623', '7030A0']
  band2(2, C_NO, C_UNIT, '기본 정보', '374151')
  for (let i = 0; i < nY; i++)
    band2(2, C_ANN_FIRST + i, C_ANN_FIRST + i, `20${years[i]}년 연간`, YEAR_COLORS[i % YEAR_COLORS.length])
  band2(2, C_AVG,      C_AVG,      `${nY}개년 연간평균`, '375623')
  band2(2, C_PEAK_AVG, C_PEAK_AVG, `${nY}개년 피크평균`, 'C55A11')
  band2(2, C_SS,       C_SS,       '안전재고',           'C00000')
  band2(2, C_CUR,      C_ORD,      '재고 현황',          '7030A0')
  band2(2, C_TGT,      C_REQ,      '✏ 발주 계획',       'C55A11')
  band2(2, C_NOTE,     C_NOTE,     '비고',               '595959')
  // k 라벨
  const kLabelCell2 = ws.getCell(2, C_K)
  kLabelCell2.value = '안전재고\n계수 k'
  kLabelCell2.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } }
  kLabelCell2.fill = fill('C00000'); kLabelCell2.border = ALL_BORDERS
  kLabelCell2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
  ws.getColumn(C_K).width = 10

  // 헤더
  ws.getRow(3).height = 42
  hdr(3, C_NO, 'NO', 5); hdr(3, C_PN, '품번', 16); hdr(3, C_NAME, '품명', 38)
  hdr(3, C_VENDOR, '구매처', 14); hdr(3, C_LT, `리드타임\n(일)`, 10); hdr(3, C_UNIT, '단위', 7)
  for (let i = 0; i < nY; i++) hdr(3, C_ANN_FIRST + i, `20${years[i]}년\n연간`, 12)
  hdr(3, C_AVG,      `${nY}개년\n연간평균`,  12)
  hdr(3, C_PEAK_AVG, `${nY}개년\n피크평균`,  12)
  hdr(3, C_SS,       `안전재고\n(피크×LT/30×k)`, 14)
  hdr(3, C_CUR, `현재고`, 10); hdr(3, C_ORD, `기발주\n(참고)`, 10)
  hdr(3, C_TGT, `목표수량`, 12); hdr(3, C_REQ, `필요발주`, 12)
  hdr(3, C_NOTE, '합산 출처 (가공파일 대조용)', 45)
  // k 입력 셀
  const kCell2 = ws.getCell(3, C_K)
  kCell2.value = safetyK
  kCell2.font = font({ bold: true, color: '0000FF' })
  kCell2.fill = fill('FFFFC0'); kCell2.border = INPUT_BORDERS()
  kCell2.numFmt = '0.0'; kCell2.alignment = { horizontal: 'center', vertical: 'middle' }
  const kRef2 = `$${cl(C_K)}$3`

  let ri = 4; let no = 1

  // 데이터 행 공통 렌더러
  const renderBunhoRow = (
    bn: string, 품명: string, 구매처: string, 리드타임: number,
    unit: string, annByYear: number[], peakByYear: number[], 현재고: number, 기발주: number,
    bgRow?: string, note?: string,
  ) => {
    const evenFill2: ExcelJS.Fill | undefined = bgRow ? fill(bgRow) : ((ri % 2 === 0) ? fill('F5F5F5') : undefined)

    const sc = (c: number, val: ExcelJS.CellValue, right = false, numFmt?: string) => {
      const cell = ws.getCell(ri, c)
      cell.value = val; cell.font = font(); cell.border = ALL_BORDERS
      if (evenFill2) cell.fill = evenFill2
      cell.alignment = { horizontal: right ? 'right' : (c <= 1 || c === C_LT || c === C_UNIT ? 'center' : 'left'), vertical: 'middle' }
      if (numFmt) cell.numFmt = numFmt
    }

    sc(C_NO, no++); sc(C_PN, bn); sc(C_NAME, 품명); sc(C_VENDOR, 구매처)
    sc(C_LT, 리드타임, true, '#,##0'); sc(C_UNIT, unit)
    for (let i = 0; i < nY; i++) sc(C_ANN_FIRST + i, annByYear[i] || null, true, '#,##0')

    const annCols = Array.from({ length: nY }, (_, i) => `${cl(C_ANN_FIRST + i)}${ri}`).join(',')
    const avgCell = ws.getCell(ri, C_AVG)
    avgCell.value = { formula: `ROUND(AVERAGE(${annCols}),0)` }
    avgCell.font = font(); avgCell.border = ALL_BORDERS; avgCell.numFmt = '#,##0'
    avgCell.alignment = { horizontal: 'right', vertical: 'middle' }
    if (evenFill2) avgCell.fill = evenFill2

    const peakAvg = peakByYear.length > 0 ? peakByYear.reduce((s, v) => s + v, 0) / peakByYear.length : 0
    const pkAvgCell = ws.getCell(ri, C_PEAK_AVG)
    pkAvgCell.value = Math.round(peakAvg)
    pkAvgCell.font = font(); pkAvgCell.border = ALL_BORDERS; pkAvgCell.numFmt = '#,##0'
    pkAvgCell.fill = fill('FCE4D6')
    pkAvgCell.alignment = { horizontal: 'right', vertical: 'middle' }

    const ssCell2 = ws.getCell(ri, C_SS)
    ssCell2.value = { formula: `ROUND(${cl(C_PEAK_AVG)}${ri}*${cl(C_LT)}${ri}/30*${kRef2},0)` }
    ssCell2.font  = font({ bold: true }); ssCell2.border = ALL_BORDERS; ssCell2.numFmt = '#,##0'
    ssCell2.fill  = fill('FFF2CC'); ssCell2.alignment = { horizontal: 'right', vertical: 'middle' }

    sc(C_CUR, 현재고 || null, true, '#,##0'); sc(C_ORD, 기발주 || null, true, '#,##0')

    const tgtCell2 = ws.getCell(ri, C_TGT)
    tgtCell2.value = null; tgtCell2.font = font({ bold: true, color: '0000FF' })
    tgtCell2.fill = fill('FFFFC0'); tgtCell2.border = INPUT_BORDERS()
    tgtCell2.numFmt = '#,##0'; tgtCell2.alignment = { horizontal: 'right', vertical: 'middle' }

    const reqCell2 = ws.getCell(ri, C_REQ)
    reqCell2.value = {
      formula: `IFERROR(MAX(${cl(C_TGT)}${ri}-${cl(C_CUR)}${ri}-${cl(C_ORD)}${ri}+${cl(C_SS)}${ri},0),"")`,
    }
    reqCell2.font = font({ bold: true, color: 'C00000' }); reqCell2.border = ALL_BORDERS
    reqCell2.numFmt = '#,##0'; reqCell2.alignment = { horizontal: 'right', vertical: 'middle' }
    if (evenFill2) reqCell2.fill = evenFill2

    const noteCell2 = ws.getCell(ri, C_NOTE)
    noteCell2.value = note ?? ''; noteCell2.font = font(); noteCell2.border = ALL_BORDERS
    if (evenFill2) noteCell2.fill = evenFill2
    noteCell2.alignment = { horizontal: 'left', vertical: 'middle' }

    ws.getRow(ri).height = 17; ri++
  }

  // 출처 요약 생성: "3타입: 2.0mm SC/PC 청색(100K) + 백색(80K) + ..."
  function buildSrcNote(srcs: Map<string, number>): string {
    if (srcs.size === 0) return ''
    const sorted = [...srcs.entries()].sort((a, b) => b[1] - a[1])
    const fmt = (n: number) => n >= 1000 ? `${Math.round(n / 1000)}K` : String(n)
    const parts = sorted.slice(0, 4).map(([lbl, n]) => `${lbl}(${fmt(n)})`)
    const suffix = sorted.length > 4 ? ` 외 ${sorted.length - 4}` : ''
    return `${sorted.length}타입: ${parts.join(' + ')}${suffix}`
  }

  // ── 케이블 섹션 ───────────────────────────────────────────
  // 현재고/기발주는 같은 품번 = 동일 실물 재고이므로 첫 번째 등장 값만 사용
  type BunhoEntry = { 품명: string; 구매처: string; 리드타임: number; 현재고: number; 기발주: number; annByYear: number[]; peakByYear: number[]; srcs: Map<string, number> }
  const byCable = new Map<string, BunhoEntry>()
  const latestYr = years[nY - 1]
  for (const row of cableRows) {
    const bn = (row.품번 || '').trim(); if (!bn) continue
    if (!byCable.has(bn)) byCable.set(bn, { 품명: row.품명, 구매처: row.구매처, 리드타임: row.리드타임, 현재고: row.현재고, 기발주: row.기발주, annByYear: new Array(nY).fill(0), peakByYear: new Array(nY).fill(0), srcs: new Map() })
    const e = byCable.get(bn)!
    const srcKey = `${row.pai} ${row.label}`
    e.srcs.set(srcKey, (e.srcs.get(srcKey) ?? 0) + (row.byYear[latestYr]?.annual ?? 0))
    for (let i = 0; i < nY; i++) {
      e.annByYear[i]  += row.byYear[years[i]]?.annual ?? 0
      e.peakByYear[i] += row.byYear[years[i]]?.peak   ?? 0
    }
  }
  if (byCable.size > 0) {
    sectionHeader('▼ 케이블 (m)', mainColor)
    for (const [bn, d] of [...byCable.entries()].sort())
      renderBunhoRow(bn, d.품명, d.구매처, d.리드타임, 'm', d.annByYear, d.peakByYear, d.현재고, d.기발주, undefined, d.srcs.size > 1 ? buildSrcNote(d.srcs) : '')
    ri++
  }

  // ── 하우징 섹션 ───────────────────────────────────────────
  // 같은 품번이 여러 하우징 타입에 공통 부품으로 등록될 수 있음
  // 연간 사용량: 타입별로 합산(올바름) / 현재고·기발주: 첫 등장값만(합산하면 N배 과대 계상)
  // 비고: 합산에 기여한 하우징 타입 목록 표시 (가공파일 대조용)
  const byHousing = new Map<string, BunhoEntry>()
  for (const row of housingRows) {
    const bn = (row.품번 || '').trim(); if (!bn) continue
    if (!byHousing.has(bn)) byHousing.set(bn, { 품명: row.품명, 구매처: row.구매처, 리드타임: row.리드타임, 현재고: row.현재고, 기발주: row.기발주, annByYear: new Array(nY).fill(0), peakByYear: new Array(nY).fill(0), srcs: new Map() })
    const e = byHousing.get(bn)!
    const srcKey = `${row.pai} ${row.label}`
    e.srcs.set(srcKey, (e.srcs.get(srcKey) ?? 0) + (row.byYear[latestYr]?.annual ?? 0))
    for (let i = 0; i < nY; i++) {
      e.annByYear[i]  += row.byYear[years[i]]?.annual ?? 0
      e.peakByYear[i] += row.byYear[years[i]]?.peak   ?? 0
    }
  }
  if (byHousing.size > 0) {
    sectionHeader('▼ 하우징 공용 부품 (EA)', '7030A0')
    for (const [bn, d] of [...byHousing.entries()].sort())
      renderBunhoRow(bn, d.품명, d.구매처, d.리드타임, 'EA', d.annByYear, d.peakByYear, d.현재고, d.기발주, undefined, buildSrcNote(d.srcs))
    ri++
  }

  // ── 페롤 섹션 ─────────────────────────────────────────────
  const connAnn  = new Map<string, number[]>()
  const connPeak = new Map<string, number[]>()
  for (const row of housingRows) {
    if (row.isSubRow) continue  // 다중 부품 하우징의 2번째+ 행은 제외 (같은 annual이 N번 합산되는 것 방지)
    const type = row.key.split('|')[1] ?? ''
    const m = type.match(/^(LC\/PC|LC\/APC|SC\/PC|SC\/APC|FC\/PC|FC\/APC)/)
    if (!m) continue
    const ct = m[1]
    if (!connAnn.has(ct))  connAnn.set(ct,  new Array(nY).fill(0))
    if (!connPeak.has(ct)) connPeak.set(ct, new Array(nY).fill(0))
    const annArr  = connAnn.get(ct)!
    const peakArr = connPeak.get(ct)!
    for (let i = 0; i < nY; i++) {
      annArr[i]  += row.byYear[years[i]]?.annual ?? 0
      peakArr[i] += row.byYear[years[i]]?.peak   ?? 0
    }
  }

  const ferruleEntries = CONN_TYPES.filter(ct => ferruleMeta[ct]?.품번 && connAnn.has(ct))
  if (ferruleEntries.length > 0) {
    sectionHeader('▼ 페롤 (커넥터 타입별 공용 — 파이 무관 합산) (EA)', '375623')
    for (const ct of ferruleEntries) {
      const fm  = ferruleMeta[ct]!
      const fiv = ferruleInv[ct] ?? { 현재고: 0, 기발주: 0 }
      const ann  = connAnn.get(ct)!
      const peak = connPeak.get(ct) ?? new Array(nY).fill(0)
      const lt  = (() => { const n = parseInt(String(fm.리드타임 ?? '')); return isNaN(n) ? 60 : n })()
      renderBunhoRow(fm.품번!, fm.품명 ?? '', fm.구매처 ?? '', lt, 'EA', ann, peak, fiv.현재고 ?? 0, fiv.기발주 ?? 0, 'F0FFF4', ct)
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
  rows:         Step3Row[],
  years:        string[],
  mainColor:    string,
  safetyK:      number,
  metadata:     Metadata,
  inventory:    Inventory,
  salesAgg?:    SalesAggResult,
  codeToCable?: Record<string, CodeCableEntry>,
): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'AJW 발주계획 시스템'

  const cableRows   = rows.filter(r => r.type === 'cable')
  const housingRows = rows.filter(r => r.type === 'housing')

  const demandMap = (salesAgg && codeToCable && Object.keys(codeToCable).length > 0)
    ? computeDemand(salesAgg, codeToCable, years)
    : undefined

  writeMainSheet(wb, '케이블 사용내역',  cableRows,   years, mainColor, 'm',  '케이블 종류', safetyK, demandMap)
  writeMainSheet(wb, '하우징 사용내역',  housingRows, years, mainColor, 'EA', '하우징 타입', safetyK)
  writeBunhoSheet(wb, cableRows, housingRows, years, mainColor, metadata.ferrule ?? {}, inventory.ferrule ?? {}, safetyK)
  writeMonthlySheet(wb, [...cableRows, ...housingRows], years, mainColor)
  writeAnomalySheet(wb)

  return wb
}
