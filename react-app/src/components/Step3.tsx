import { useState } from 'react'
import ExcelJS from 'exceljs'
import { buildStep3Plan } from '../lib/step3Core'
import type { Step3Row } from '../lib/step3Core'
import type { Metadata, Inventory, AppSettings } from '../lib/types'
import FileUploader from './FileUploader'

interface Props {
  metadata:  Metadata
  inventory: Inventory
  settings:  AppSettings
}

const MAIN_COLOR = 'FF1F3864'
const HEADER_BG  = 'FFBDD7EE'
const BORDER: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FF000000' } }
const ALL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER }
const num = (v: number, unit = '') => v === 0 ? '—' : v.toLocaleString() + (unit ? ' ' + unit : '')

export default function Step3({ metadata, inventory, settings }: Props) {
  const [logs, setLogs]         = useState<string[]>([])
  const [running, setRunning]   = useState(false)
  const [done, setDone]         = useState(false)
  const [gaongFile, setGaongFile] = useState<File | null>(null)
  const [rows, setRows]         = useState<Step3Row[]>([])
  const [years, setYears]       = useState<string[]>([])

  const run = async () => {
    if (!gaongFile) return alert('가공파일을 선택해주세요.')
    setRunning(true); setDone(false); setLogs([]); setRows([])
    try {
      const buf = await gaongFile.arrayBuffer()
      const result = buildStep3Plan(buf, metadata, inventory, settings.lead_time_default)
      setLogs(result.logs)
      setRows(result.rows)
      setYears(result.years)
      setDone(true)
    } catch (e) {
      setLogs([`❌ 오류: ${e}`])
    } finally {
      setRunning(false)
    }
  }

  const exportExcel = async () => {
    if (!rows.length) return
    const wb = new ExcelJS.Workbook()
    wb.creator = 'AJW 발주계획 시스템'
    writeSheet(wb, '케이블 발주계획',  rows.filter(r => r.type === 'cable'),   years, settings.colors.main_header)
    writeSheet(wb, '하우징 발주계획',  rows.filter(r => r.type === 'housing'),  years, settings.colors.main_header)
    const buf  = await wb.xlsx.writeBuffer()
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url  = URL.createObjectURL(blob)
    const a    = document.createElement('a')
    a.href = url; a.download = `발주계획_${new Date().toISOString().slice(0,10)}.xlsx`
    a.click(); URL.revokeObjectURL(url)
  }

  const cableRows   = rows.filter(r => r.type === 'cable')
  const housingRows = rows.filter(r => r.type === 'housing')
  const missingPn   = rows.filter(r => !r.품번).length
  const latestYr    = years[years.length - 1]

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>STEP 3 — 발주계획 생성</b>: STEP 1 가공파일을 업로드하면 자재별 <b>안전재고</b>와
        <b>발주 필요량</b>을 한눈에 파악하고 Excel로 다운로드합니다.
        <span className="text-gray-500 ml-1">안전재고 = 월최대 × 리드타임(일) / 30</span>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-start">
        <div className="md:col-span-2">
          <FileUploader
            label="가공파일.xlsx (STEP 1 결과)"
            fileName={gaongFile?.name ?? ''}
            onFile={setGaongFile}
          />
        </div>
        <div className="bg-[#e8f4fd] rounded-lg p-4 text-sm text-gray-700 space-y-1">
          <div className="font-semibold">현재 등록 데이터</div>
          <div>케이블 품번 <b>{Object.keys(metadata.cable).length}</b>타입</div>
          <div>하우징 품번 <b>{Object.keys(metadata.housing).length}</b>타입</div>
          <div className="text-gray-500 text-xs">리드타임 기본값: {settings.lead_time_default}일</div>
        </div>
      </div>

      <hr className="border-gray-200" />

      <div className="flex gap-2">
        <button
          onClick={run}
          disabled={running || !gaongFile}
          className="flex-1 bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold py-2 px-4 rounded transition text-sm"
        >
          {running ? '⏳ 분석 중...' : '▶ STEP 3 실행 — 발주계획 생성'}
        </button>
        {done && (
          <button
            onClick={exportExcel}
            className="px-4 py-2 text-sm bg-[#2E75B6] hover:bg-[#1a5a9e] text-white font-semibold rounded transition"
          >
            📥 Excel 다운로드
          </button>
        )}
        {(done || logs.length > 0) && (
          <button
            onClick={() => { setDone(false); setLogs([]); setGaongFile(null); setRows([]); setYears([]) }}
            className="px-4 py-2 text-sm border border-gray-300 rounded text-gray-600 hover:bg-gray-100 transition"
          >
            🗑 초기화
          </button>
        )}
      </div>

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs leading-5 max-h-32 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {done && missingPn > 0 && (
        <div className="bg-[#FFF3CD] border border-yellow-300 px-4 py-3 rounded-md text-sm text-[#856404]">
          ⚠ 품번 미등록 <b>{missingPn}건</b> — <b>품번 관리</b> 탭에서 품번·품명·구매처·리드타임을 입력 후 재실행하세요.
        </div>
      )}

      {done && rows.length > 0 && (
        <div className="space-y-4">
          {/* 케이블 */}
          <SectionTable
            title={`케이블 자재 (${cableRows.length}타입)`}
            rows={cableRows}
            years={years}
            latestYr={latestYr}
            unitSuffix="m"
          />
          {/* 하우징 */}
          <SectionTable
            title={`하우징 자재 (${housingRows.length}타입)`}
            rows={housingRows}
            years={years}
            latestYr={latestYr}
            unitSuffix="EA"
          />
        </div>
      )}
    </div>
  )
}

// ── 섹션 테이블 컴포넌트 ─────────────────────────────────────────
function SectionTable({
  title, rows, years, latestYr, unitSuffix
}: {
  title: string; rows: Step3Row[]; years: string[]; latestYr: string; unitSuffix: string
}) {
  if (!rows.length) return null
  return (
    <div>
      <div className="text-sm font-semibold text-gray-700 mb-1.5">{title}</div>
      <div className="overflow-x-auto border border-gray-200 rounded">
        <table className="text-xs w-full whitespace-nowrap">
          <thead className="bg-gray-50">
            <tr>
              <th className="px-3 py-2 text-left border-b font-semibold text-gray-600" rowSpan={2}>타입</th>
              <th className="px-3 py-2 text-left border-b font-semibold text-gray-600" rowSpan={2}>파이</th>
              <th className="px-3 py-2 text-left border-b font-semibold text-gray-600" rowSpan={2}>품번</th>
              <th className="px-3 py-2 text-left border-b font-semibold text-gray-600" rowSpan={2}>품명</th>
              <th className="px-3 py-2 text-left border-b font-semibold text-gray-600" rowSpan={2}>구매처</th>
              <th className="px-2 py-1 text-center border-b font-semibold text-gray-600" rowSpan={2}>LT<br/>(일)</th>
              <th className="px-2 py-1 text-center border-b border-l font-semibold text-blue-700"
                colSpan={years.length}>연간 사용량({unitSuffix})</th>
              <th className="px-2 py-1 text-center border-b font-semibold text-orange-700" rowSpan={2}>월최대<br/>({latestYr}년)</th>
              <th className="px-2 py-1 text-center border-b font-semibold text-red-700" rowSpan={2}>안전재고<br/>({unitSuffix})</th>
              <th className="px-2 py-1 text-center border-b font-semibold text-purple-700" rowSpan={2}>현재고</th>
              <th className="px-2 py-1 text-center border-b font-semibold text-purple-700" rowSpan={2}>기발주</th>
              <th className="px-2 py-1 text-center border-b font-semibold text-green-700" rowSpan={2}>발주<br/>필요량</th>
            </tr>
            <tr>
              {years.map(yr => (
                <th key={yr} className="px-2 py-1 text-right border-b border-l text-blue-600 font-normal">{yr}년</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => {
              const missingPn = !row.품번
              const needsOrder = row.발주필요량 > 0
              return (
                <tr key={row.key} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                  <td className="px-3 py-1.5 font-medium">{row.label}</td>
                  <td className="px-3 py-1.5 text-gray-500">{row.pai}</td>
                  <td className={`px-3 py-1.5 font-mono ${missingPn ? 'text-red-500 font-semibold' : 'text-gray-600'}`}>
                    {row.품번 || '미등록'}
                  </td>
                  <td className="px-3 py-1.5 max-w-[200px] truncate" title={row.품명}>{row.품명 || '—'}</td>
                  <td className="px-3 py-1.5 text-gray-500">{row.구매처 || '—'}</td>
                  <td className="px-2 py-1.5 text-center">{row.리드타임}</td>
                  {years.map(yr => (
                    <td key={yr} className="px-2 py-1.5 text-right border-l">
                      {num(row.byYear[yr]?.annual ?? 0)}
                    </td>
                  ))}
                  <td className="px-2 py-1.5 text-right text-orange-700">{num(row.latestPeak)}</td>
                  <td className="px-2 py-1.5 text-right text-red-700 font-semibold">{num(row.안전재고)}</td>
                  <td className="px-2 py-1.5 text-right text-purple-700">{num(row.현재고)}</td>
                  <td className="px-2 py-1.5 text-right text-purple-700">{num(row.기발주)}</td>
                  <td className={`px-2 py-1.5 text-right font-bold ${needsOrder ? 'text-green-700 bg-green-50' : 'text-gray-400'}`}>
                    {needsOrder ? row.발주필요량.toLocaleString() : '—'}
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
      <div className="text-xs text-gray-400 mt-1">
        발주필요량 = 최근년 연간사용량 + 안전재고 − 현재고 − 기발주
      </div>
    </div>
  )
}

// ── Excel 출력 ────────────────────────────────────────────────────
function writeSheet(
  wb:      ExcelJS.Workbook,
  name:    string,
  rows:    Step3Row[],
  years:   string[],
  color:   string,
) {
  if (!rows.length) return
  const ws      = wb.addWorksheet(name)
  const latestYr = years[years.length - 1]
  const unitCol  = rows[0]?.unit === 'm' ? 'm' : 'EA'
  const nYears   = years.length
  const mainFill  = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF' + color } }
  const cell = (r: number, c: number) => ws.getCell(r, c)
  const hdr  = (r: number, c: number, v: string, w?: number) => {
    const cl = cell(r, c)
    cl.value = v; cl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } }
    cl.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FF1F3864' } }
    cl.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    cl.border = ALL_BORDERS
    if (w) ws.getColumn(c).width = w
  }

  // 타이틀
  const totalCols = 7 + nYears + 4
  ws.mergeCells(1, 1, 1, totalCols)
  const titleCell = cell(1, 1)
  titleCell.value = name
  titleCell.fill = mainFill
  titleCell.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } }
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 26

  // 헤더 행
  ws.getRow(2).height = 36
  const baseHdrs: [string, number][] = [
    ['NO', 5], ['파이', 8], ['타입', 20], ['품번', 16], ['품명', 38], ['구매처', 14], [`LT\n(일)`, 9],
  ]
  baseHdrs.forEach(([h, w], i) => hdr(2, i + 1, h, w))
  let col = 8
  years.forEach(yr => { hdr(2, col, `${yr}년\n연간(${unitCol})`, 12); col++ })
  hdr(2, col++, `피크\n(${latestYr}년)`, 11)
  hdr(2, col++, `안전재고\n(${unitCol})`, 11)
  hdr(2, col++, '현재고', 10)
  hdr(2, col++, `발주\n필요량`, 12)

  // 데이터 행
  rows.forEach((row, idx) => {
    const ri   = idx + 3
    const even = idx % 2 === 0
    const bg   = even ? undefined : { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFF5F5F5' } }
    const setCell = (c: number, v: ExcelJS.CellValue, right = false) => {
      const cl = cell(ri, c)
      cl.value = v
      cl.font = { name: 'Arial', size: 9 }
      cl.border = ALL_BORDERS
      if (bg) cl.fill = bg
      cl.alignment = { horizontal: right ? 'right' : c <= 2 ? 'center' : 'left', vertical: 'middle' }
      if (right) cl.numFmt = '#,##0'
    }
    setCell(1, idx + 1)
    setCell(2, row.pai)
    setCell(3, row.label)
    setCell(4, row.품번 || '(미등록)')
    setCell(5, row.품명 || '')
    setCell(6, row.구매처 || '')
    setCell(7, row.리드타임, true)
    let c2 = 8
    years.forEach(yr => { setCell(c2, row.byYear[yr]?.annual || null, true); c2++ })
    setCell(c2++, row.latestPeak || null, true)
    // 안전재고 — 강조
    const safeCell = cell(ri, c2)
    safeCell.value = row.안전재고 || null
    safeCell.font  = { name: 'Arial', bold: true, size: 9 }
    safeCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } }
    safeCell.border = ALL_BORDERS; safeCell.numFmt = '#,##0'
    safeCell.alignment = { horizontal: 'right', vertical: 'middle' }
    c2++
    setCell(c2++, row.현재고 || null, true)
    // 발주필요량 — 강조
    const orderCell = cell(ri, c2)
    orderCell.value = row.발주필요량 || null
    orderCell.font  = { name: 'Arial', bold: true, size: 9, color: { argb: row.발주필요량 > 0 ? 'FF375623' : 'FF999999' } }
    orderCell.fill  = row.발주필요량 > 0
      ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }
      : (bg ?? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } })
    orderCell.border = ALL_BORDERS; orderCell.numFmt = '#,##0'
    orderCell.alignment = { horizontal: 'right', vertical: 'middle' }
    ws.getRow(ri).height = 17
  })

  ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 2 }]
}
