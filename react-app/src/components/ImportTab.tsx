import { useState, useEffect, useMemo } from 'react'
import ExcelJS from 'exceljs'
import FileUploader from './FileUploader'
import { parseStockFile } from '../lib/parse/parseStockFile'
import { loadDetailedSales, loadItemCodes } from '../lib/supabase'
import type { ItemCode } from '../lib/supabase'
import type { DetailedSalesRow } from '../lib/parse/parseDetailedSales'
import { classifyOjc } from '../lib/ojcFilter'
import { downloadXlsx, today } from '../lib/download'

interface RawRow {
  code:        string
  name:        string
  spec:        string
  category:    string
  stock:       number
  peakMonthly: number
  coverage:    number
}

interface PlanRow extends RawRow {
  orderNeeded: number
}

type RowStatus = 'urgent' | 'warning' | 'ok' | 'no_data'

function rowStatus(row: PlanRow, target: number): RowStatus {
  if (row.peakMonthly === 0) return 'no_data'
  if (row.coverage >= target) return 'ok'
  if (row.coverage >= 1)      return 'warning'
  return 'urgent'
}

const STATUS_LABEL: Record<RowStatus, string> = {
  urgent:  '발주필요',
  warning: '주의',
  ok:      '재고충분',
  no_data: '이력없음',
}

const STATUS_BADGE: Record<RowStatus, string> = {
  urgent:  'text-red-700 bg-red-100',
  warning: 'text-yellow-700 bg-yellow-100',
  ok:      'text-green-700 bg-green-100',
  no_data: 'text-gray-400 bg-gray-100',
}

const STATUS_ROW_BG: Record<RowStatus, string> = {
  urgent:  'bg-red-50',
  warning: 'bg-yellow-50',
  ok:      '',
  no_data: '',
}

export default function ImportTab() {
  const [stockFile,    setStockFile]  = useState<File | null>(null)
  const [salesRows,    setSalesRows]  = useState<DetailedSalesRow[]>([])
  const [itemCodes,    setItemCodes]  = useState<ItemCode[]>([])
  const [targetMonths, setTarget]     = useState(3)
  const [rawRows,      setRawRows]    = useState<RawRow[]>([])
  const [error,        setError]      = useState('')
  const [loading,      setLoading]    = useState(false)

  useEffect(() => {
    loadDetailedSales<DetailedSalesRow>().then(setSalesRows)
    loadItemCodes().then(setItemCodes)
  }, [])

  // Derive final rows whenever raw data or targetMonths changes
  const rows = useMemo<PlanRow[]>(() => {
    const ord: Record<RowStatus, number> = { urgent: 0, warning: 1, ok: 2, no_data: 3 }
    return rawRows
      .map(r => ({
        ...r,
        orderNeeded: r.peakMonthly > 0
          ? Math.max(0, targetMonths * r.peakMonthly - r.stock)
          : 0,
      }))
      .sort((a, b) => {
        const sa = rowStatus(a, targetMonths)
        const sb = rowStatus(b, targetMonths)
        if (ord[sa] !== ord[sb]) return ord[sa] - ord[sb]
        if (a.category !== b.category) return a.category.localeCompare(b.category)
        return a.code.localeCompare(b.code)
      })
  }, [rawRows, targetMonths])

  async function handleRun() {
    if (!stockFile)         { setError('재고 파일을 선택하세요.'); return }
    if (!salesRows.length)  { setError('판매 현황 분석 탭에서 판매 데이터를 먼저 로드하세요.'); return }
    setLoading(true); setError('')
    try {
      const stockMap = await parseStockFile(stockFile)
      const itemMap  = new Map<string, ItemCode>(itemCodes.map(i => [i.code, i]))

      // Aggregate monthly sales per code: (year-month) → qty
      const monthlyMap = new Map<string, Map<string, number>>()
      const nameMap    = new Map<string, string>()
      const catMap     = new Map<string, string>()

      for (const r of salesRows) {
        const cat = classifyOjc(r.name)
        if (!cat) continue
        nameMap.set(r.code, r.name)
        catMap.set(r.code, cat)
        if (!monthlyMap.has(r.code)) monthlyMap.set(r.code, new Map())
        const ym = `${r.year}-${r.month}`
        const m  = monthlyMap.get(r.code)!
        m.set(ym, (m.get(ym) ?? 0) + r.qty)
      }

      // Also add stock-only codes that are OJC by item list
      for (const [code] of stockMap) {
        if (monthlyMap.has(code)) continue
        const item = itemMap.get(code)
        if (!item) continue
        const cat = classifyOjc(item.name)
        if (!cat) continue
        nameMap.set(code, item.name)
        catMap.set(code, cat)
      }

      // Build result
      const allCodes = new Set([...monthlyMap.keys(), ...[...stockMap.keys()].filter(c => nameMap.has(c))])
      const result: RawRow[] = []

      for (const code of allCodes) {
        const item        = itemMap.get(code)
        const name        = item?.name ?? nameMap.get(code) ?? code
        const spec        = item?.spec ?? ''
        const category    = catMap.get(code) ?? ''
        const stock       = stockMap.get(code) ?? 0
        const mMap        = monthlyMap.get(code)
        const peakMonthly = mMap && mMap.size > 0 ? Math.max(...mMap.values()) : 0
        const coverage    = peakMonthly > 0 ? stock / peakMonthly : 0
        result.push({ code, name, spec, category, stock, peakMonthly, coverage })
      }

      setRawRows(result)
    } catch (e) {
      setError(String(e))
    } finally {
      setLoading(false)
    }
  }

  function coverageText(row: PlanRow): string {
    if (row.peakMonthly === 0) return '—'
    return row.coverage.toFixed(1) + '개월'
  }

  async function handleDownload() {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('완제품 수입 발주계획')

    const HEADERS = [
      '카테고리', '품목코드', '품목명', '규격',
      '현재고', '월간최고판매', '커버리지(개월)',
      `목표재고(${targetMonths}개월)`, '발주필요량', '상태',
    ]
    ws.addRow(HEADERS)
    ws.getRow(1).eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } }
      cell.font      = { bold: true, size: 9, color: { argb: 'FFFFFFFF' }, name: 'Arial' }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
      cell.border    = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
    })
    ws.getRow(1).height = 20

    for (const row of rows) {
      const s       = rowStatus(row, targetMonths)
      const xlRow   = ws.addRow([
        row.category,
        row.code,
        row.name,
        row.spec,
        row.stock,
        row.peakMonthly || '',
        row.peakMonthly ? parseFloat(row.coverage.toFixed(1)) : '',
        row.peakMonthly ? row.peakMonthly * targetMonths : '',
        row.orderNeeded || '',
        STATUS_LABEL[s],
      ])
      const fillArgb = s === 'urgent' ? 'FFFFCCCC' : s === 'warning' ? 'FFFFFBCC' : 'FFFFFFFF'
      xlRow.eachCell(cell => {
        cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } }
        cell.font   = { size: 9, name: 'Arial' }
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
      })
    }

    ws.columns = [
      { width: 16 }, { width: 22 }, { width: 40 }, { width: 18 },
      { width: 10 }, { width: 12 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 12 },
    ]
    for (let c = 5; c <= 9; c++) ws.getColumn(c).alignment = { horizontal: 'right' }

    const buf = await wb.xlsx.writeBuffer()
    downloadXlsx(buf as ArrayBuffer, `완제품수입발주계획_${today()}.xlsx`)
  }

  const urgentCount  = rows.filter(r => rowStatus(r, targetMonths) === 'urgent').length
  const warningCount = rows.filter(r => rowStatus(r, targetMonths) === 'warning').length
  const okCount      = rows.filter(r => rowStatus(r, targetMonths) === 'ok').length
  const noDataCount  = rows.filter(r => rowStatus(r, targetMonths) === 'no_data').length

  return (
    <div className="p-6 space-y-5 max-w-full">
      <div>
        <h2 className="text-lg font-semibold text-gray-800">🚢 완제품 수입 발주계획</h2>
        <p className="text-sm text-gray-500 mt-0.5">현재고와 판매 데이터를 기반으로 수입 필요 품목을 자동 도출합니다.</p>
      </div>

      {/* 목표 개월수 설정 */}
      <div className="bg-gray-50 border rounded-lg px-4 py-3 flex items-center gap-4 flex-wrap">
        <span className="text-sm font-medium text-gray-700">목표 재고 개월수</span>
        <input
          type="number" min={1} max={12} value={targetMonths}
          onChange={e => setTarget(Math.max(1, parseInt(e.target.value) || 1))}
          className="w-16 border rounded px-2 py-1 text-sm text-center"
        />
        <span className="text-sm text-gray-500">개월</span>
        <span className="text-xs text-gray-400 border-l pl-4">
          커버리지 = 현재고 ÷ 월간최고  /  발주필요량 = (월간최고 × 목표개월) − 현재고
        </span>
      </div>

      {/* 재고 파일 업로드 */}
      <FileUploader
        label="재고 파일 (이카운트 수불부 또는 EMP 재고현황 — 자동 감지)"
        fileName={stockFile?.name ?? ''}
        onFile={setStockFile}
      />

      {/* 판매 데이터 상태 */}
      <div className={`text-sm rounded-lg px-4 py-2.5 flex items-center gap-2 border ${
        salesRows.length
          ? 'bg-green-50 text-green-700 border-green-200'
          : 'bg-yellow-50 text-yellow-700 border-yellow-200'
      }`}>
        {salesRows.length
          ? `✅ 판매 현황 분석 데이터 로드됨 — ${salesRows.length.toLocaleString()}행`
          : '⚠️ 판매 데이터 없음 — 판매 현황 분석 탭에서 파일을 먼저 업로드하세요.'}
      </div>

      {/* 실행 버튼 */}
      <button
        onClick={handleRun}
        disabled={loading || !stockFile || !salesRows.length}
        className="px-6 py-2 bg-blue-600 text-white text-sm font-semibold rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition"
      >
        {loading ? '계산 중…' : '📊 발주계획 생성'}
      </button>

      {error && <p className="text-sm text-red-600">{error}</p>}

      {/* 결과 */}
      {rows.length > 0 && (
        <div className="space-y-3">
          {/* 요약 배너 */}
          <div className="flex items-center gap-3 bg-gray-50 border rounded-lg px-4 py-2.5 flex-wrap text-sm">
            <span className="font-medium text-gray-700">총 {rows.length}개 품목</span>
            {urgentCount  > 0 && <span className="text-red-600 font-medium">🔴 발주필요 {urgentCount}</span>}
            {warningCount > 0 && <span className="text-yellow-600 font-medium">⚠️ 주의 {warningCount}</span>}
            {okCount      > 0 && <span className="text-green-600">✅ 재고충분 {okCount}</span>}
            {noDataCount  > 0 && <span className="text-gray-400">— 이력없음 {noDataCount}</span>}
            <button
              onClick={handleDownload}
              className="ml-auto px-4 py-1.5 bg-green-600 text-white text-xs font-medium rounded hover:bg-green-700 transition"
            >
              📥 Excel 다운로드
            </button>
          </div>

          {/* 결과 테이블 */}
          <div className="overflow-x-auto rounded-lg border">
            <table className="min-w-full text-sm">
              <thead>
                <tr className="bg-[#1F3864] text-white text-xs">
                  <th className="px-3 py-2 text-left whitespace-nowrap">카테고리</th>
                  <th className="px-3 py-2 text-left whitespace-nowrap">품목코드</th>
                  <th className="px-3 py-2 text-left whitespace-nowrap">품목명</th>
                  <th className="px-3 py-2 text-left whitespace-nowrap">규격</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">현재고</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">월간최고</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">커버리지</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">발주필요량</th>
                  <th className="px-3 py-2 text-center whitespace-nowrap">상태</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {rows.map((row, i) => {
                  const s  = rowStatus(row, targetMonths)
                  const bg = STATUS_ROW_BG[s] || (i % 2 === 0 ? 'bg-white' : 'bg-gray-50')
                  return (
                    <tr key={row.code} className={bg}>
                      <td className="px-3 py-1.5 text-xs text-gray-600 whitespace-nowrap">{row.category}</td>
                      <td className="px-3 py-1.5 font-mono text-xs whitespace-nowrap">{row.code}</td>
                      <td className="px-3 py-1.5 whitespace-nowrap">{row.name}</td>
                      <td className="px-3 py-1.5 text-xs text-gray-500 whitespace-nowrap">{row.spec}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{row.stock.toLocaleString()}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums text-gray-600">
                        {row.peakMonthly > 0 ? row.peakMonthly.toLocaleString() : '—'}
                      </td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{coverageText(row)}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums font-medium">
                        {row.orderNeeded > 0 ? row.orderNeeded.toLocaleString() : '—'}
                      </td>
                      <td className="px-3 py-1.5 text-center">
                        <span className={`inline-block text-xs font-medium px-2 py-0.5 rounded-full ${STATUS_BADGE[s]}`}>
                          {STATUS_LABEL[s]}
                        </span>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}
