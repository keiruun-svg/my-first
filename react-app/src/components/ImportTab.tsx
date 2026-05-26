import { useState, useEffect, useMemo } from 'react'
import ExcelJS from 'exceljs'
import FileUploader from './FileUploader'
import ImportOrdersView from './ImportOrdersView'
import { parseStockFile } from '../lib/parse/parseStockFile'
import { loadDetailedSales, saveDetailedSales, loadItemCodes } from '../lib/supabase'
import type { ItemCode } from '../lib/supabase'
import { parseDetailedSalesFile } from '../lib/parse/parseDetailedSales'
import type { DetailedSalesRow } from '../lib/parse/parseDetailedSales'
import { classifyOjc } from '../lib/ojcFilter'
import { downloadXlsx, today } from '../lib/download'

interface YearStat { months: number; total: number }

interface RawRow {
  code:        string
  name:        string
  spec:        string
  category:    string
  stock:       number
  peakMonthly: number
  salesMonths: number
  coverage:    number
  byYear:      Record<string, YearStat>  // key: '23' '24' '25'
}

interface PlanRow extends RawRow {
  orderNeeded: number
}

type RowStatus = 'urgent' | 'warning' | 'ok' | 'intermittent' | 'no_data'

function rowStatus(row: PlanRow, target: number, minMonths: number): RowStatus {
  if (row.peakMonthly === 0)       return 'no_data'
  if (row.salesMonths < minMonths) return 'intermittent'
  if (row.coverage >= target)      return 'ok'
  if (row.coverage >= 1)           return 'warning'
  return 'urgent'
}

const STATUS_LABEL: Record<RowStatus, string> = {
  urgent: '발주필요', warning: '주의', ok: '재고충분', intermittent: '간헐적수요', no_data: '이력없음',
}
const STATUS_BADGE: Record<RowStatus, string> = {
  urgent: 'text-red-700 bg-red-100', warning: 'text-yellow-700 bg-yellow-100',
  ok: 'text-green-700 bg-green-100', intermittent: 'text-purple-700 bg-purple-100',
  no_data: 'text-gray-400 bg-gray-100',
}
const STATUS_ROW_BG: Record<RowStatus, string> = {
  urgent: 'bg-red-50', warning: 'bg-yellow-50', ok: '', intermittent: 'bg-purple-50', no_data: '',
}
const STATUS_SORT: Record<RowStatus, number> = {
  urgent: 0, warning: 1, ok: 2, intermittent: 3, no_data: 4,
}

type SubTab = 'plan' | 'intermittent' | 'orders'

export default function ImportTab() {
  const [stockFile,    setStockFile]  = useState<File | null>(null)
  const [salesFile,    setSalesFile]  = useState<File | null>(null)
  const [salesRows,    setSalesRows]  = useState<DetailedSalesRow[]>([])
  const [itemCodes,    setItemCodes]  = useState<ItemCode[]>([])
  const [targetMonths, setTarget]     = useState(3)
  const [minMonths,    setMinMonths]  = useState(3)
  const [rawRows,      setRawRows]    = useState<RawRow[]>([])
  const [years,        setYears]      = useState<string[]>([])
  const [error,        setError]      = useState('')
  const [loading,      setLoading]    = useState(false)
  const [subTab,       setSubTab]     = useState<SubTab>('plan')

  useEffect(() => {
    loadDetailedSales<DetailedSalesRow>().then(setSalesRows)
    loadItemCodes().then(setItemCodes)
  }, [])

  async function handleSalesFile(file: File) {
    setSalesFile(file)
    try {
      const buf  = await file.arrayBuffer()
      const logs: string[] = []
      const rows = parseDetailedSalesFile(buf, logs)
      setSalesRows(rows)
      saveDetailedSales(rows)
    } catch (e) {
      setError('판매 파일 파싱 오류: ' + String(e))
    }
  }

  const rows = useMemo<PlanRow[]>(() => {
    return rawRows
      .map(r => ({
        ...r,
        orderNeeded: r.peakMonthly > 0 && r.salesMonths >= minMonths
          ? Math.max(0, targetMonths * r.peakMonthly - r.stock)
          : 0,
      }))
      .sort((a, b) => {
        const sa = rowStatus(a, targetMonths, minMonths)
        const sb = rowStatus(b, targetMonths, minMonths)
        if (STATUS_SORT[sa] !== STATUS_SORT[sb]) return STATUS_SORT[sa] - STATUS_SORT[sb]
        if (b.salesMonths !== a.salesMonths)     return b.salesMonths - a.salesMonths
        if (a.category !== b.category)           return a.category.localeCompare(b.category)
        return a.code.localeCompare(b.code)
      })
  }, [rawRows, targetMonths, minMonths])

  const intermittentRows = useMemo(
    () => rows.filter(r => rowStatus(r, targetMonths, minMonths) === 'intermittent')
              .sort((a, b) => b.salesMonths - a.salesMonths || b.peakMonthly - a.peakMonthly),
    [rows, targetMonths, minMonths],
  )

  async function handleRun() {
    if (!stockFile)        { setError('재고 파일을 선택하세요.'); return }
    if (!salesRows.length) { setError('판매 파일을 업로드하거나 판매 현황 분석 탭에서 먼저 로드하세요.'); return }
    setLoading(true); setError('')
    try {
      const stockMap = await parseStockFile(stockFile)
      const itemMap  = new Map<string, ItemCode>(itemCodes.map(i => [i.code, i]))

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

      for (const [code] of stockMap) {
        if (monthlyMap.has(code)) continue
        const item = itemMap.get(code)
        if (!item) continue
        const cat = classifyOjc(item.name)
        if (!cat) continue
        nameMap.set(code, item.name)
        catMap.set(code, cat)
      }

      // 전체 연도 목록 수집
      const yearSet = new Set<string>()
      for (const mMap of monthlyMap.values()) {
        for (const ym of mMap.keys()) yearSet.add(ym.split('-')[0])
      }
      const sortedYears = [...yearSet].sort()
      setYears(sortedYears)

      const allCodes = new Set([
        ...monthlyMap.keys(),
        ...[...stockMap.keys()].filter(c => nameMap.has(c)),
      ])
      const result: RawRow[] = []

      for (const code of allCodes) {
        const item        = itemMap.get(code)
        const name        = item?.name ?? nameMap.get(code) ?? code
        const spec        = item?.spec ?? ''
        const category    = catMap.get(code) ?? ''
        const stock       = stockMap.get(code) ?? 0
        const mMap        = monthlyMap.get(code)
        const salesMonths = mMap?.size ?? 0
        const peakMonthly = mMap && salesMonths > 0 ? Math.max(...mMap.values()) : 0
        const coverage    = peakMonthly > 0 ? stock / peakMonthly : 0

        // 연도별 집계
        const byYear: Record<string, YearStat> = {}
        if (mMap) {
          for (const [ym, qty] of mMap) {
            const yr = ym.split('-')[0]
            if (!byYear[yr]) byYear[yr] = { months: 0, total: 0 }
            byYear[yr].months++
            byYear[yr].total += qty
          }
        }

        result.push({ code, name, spec, category, stock, peakMonthly, salesMonths, coverage, byYear })
      }

      setRawRows(result)
    } catch (e) {
      setError(String(e))
    } finally {
      setLoading(false)
    }
  }

  async function handleDownload() {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('완제품 수입 발주계획')

    ws.addRow([
      '카테고리', '품목코드', '품목명', '규격',
      '현재고', '판매월수', '월간최고', '커버리지(개월)',
      `목표재고(${targetMonths}개월)`, '발주필요량', '상태',
    ])
    ws.getRow(1).eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } }
      cell.font      = { bold: true, size: 9, color: { argb: 'FFFFFFFF' }, name: 'Arial' }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
      cell.border    = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
    })
    ws.getRow(1).height = 20

    for (const row of rows) {
      const s     = rowStatus(row, targetMonths, minMonths)
      const xlRow = ws.addRow([
        row.category, row.code, row.name, row.spec,
        row.stock,
        row.salesMonths || '',
        row.peakMonthly || '',
        row.peakMonthly ? parseFloat(row.coverage.toFixed(1)) : '',
        row.peakMonthly && s !== 'intermittent' ? row.peakMonthly * targetMonths : '',
        row.orderNeeded || '',
        STATUS_LABEL[s],
      ])
      const fillArgb =
        s === 'urgent'       ? 'FFFFCCCC' :
        s === 'warning'      ? 'FFFFFBCC' :
        s === 'intermittent' ? 'FFF3E8FF' : 'FFFFFFFF'
      xlRow.eachCell(cell => {
        cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } }
        cell.font   = { size: 9, name: 'Arial' }
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
      })
    }

    ws.columns = [
      { width: 16 }, { width: 22 }, { width: 38 }, { width: 18 },
      { width: 10 }, { width: 10 }, { width: 12 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 12 },
    ]
    for (let c = 5; c <= 10; c++) ws.getColumn(c).alignment = { horizontal: 'right' }

    // 간헐적 수요 시트
    if (intermittentRows.length > 0) {
      const ws2 = wb.addWorksheet('간헐적 수요')
      const iHeaders = ['카테고리', '품목코드', '품목명', '규격', '현재고', '판매월수(전체)', '최고단월수량',
        ...years.map(y => `20${y}년`)]
      ws2.addRow(iHeaders)
      ws2.getRow(1).eachCell(cell => {
        cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4B2E7E' } }
        cell.font      = { bold: true, size: 9, color: { argb: 'FFFFFFFF' }, name: 'Arial' }
        cell.alignment = { horizontal: 'center', vertical: 'middle' }
        cell.border    = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
      })
      ws2.getRow(1).height = 20

      for (const row of intermittentRows) {
        const yearCells = years.map(yr => {
          const s = row.byYear[yr]
          return s ? `${s.months}회 / ${s.total.toLocaleString()}EA` : '—'
        })
        const xlRow = ws2.addRow([
          row.category, row.code, row.name, row.spec,
          row.stock, row.salesMonths, row.peakMonthly,
          ...yearCells,
        ])
        xlRow.eachCell(cell => {
          cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E8FF' } }
          cell.font   = { size: 9, name: 'Arial' }
          cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
        })
      }
      ws2.columns = [
        { width: 16 }, { width: 22 }, { width: 38 }, { width: 18 },
        { width: 10 }, { width: 12 }, { width: 12 },
        ...years.map(() => ({ width: 16 })),
      ]
    }

    const buf = await wb.xlsx.writeBuffer()
    downloadXlsx(buf as ArrayBuffer, `완제품수입발주계획_${today()}.xlsx`)
  }

  const counts = useMemo(() => ({
    urgent:       rows.filter(r => rowStatus(r, targetMonths, minMonths) === 'urgent').length,
    warning:      rows.filter(r => rowStatus(r, targetMonths, minMonths) === 'warning').length,
    ok:           rows.filter(r => rowStatus(r, targetMonths, minMonths) === 'ok').length,
    intermittent: rows.filter(r => rowStatus(r, targetMonths, minMonths) === 'intermittent').length,
    no_data:      rows.filter(r => rowStatus(r, targetMonths, minMonths) === 'no_data').length,
  }), [rows, targetMonths, minMonths])

  const tabCls = (id: SubTab) =>
    `px-4 py-2 text-sm font-medium border-b-2 transition-colors ${
      subTab === id
        ? 'border-blue-600 text-blue-700'
        : 'border-transparent text-gray-500 hover:text-gray-700'
    }`

  return (
    <div className="p-6 space-y-5 max-w-full">
      <div>
        <h2 className="text-lg font-semibold text-gray-800">🚢 완제품 수입 관리</h2>
        <p className="text-sm text-gray-500 mt-0.5">수입 발주계획 자동 도출 및 발주 현황 관리</p>
      </div>

      {/* 탭 바 — 항상 표시 */}
      <div className="border-b flex gap-1">
        <button className={tabCls('plan')} onClick={() => setSubTab('plan')}>
          📋 발주계획
        </button>
        <button className={tabCls('intermittent')} onClick={() => setSubTab('intermittent')}>
          🔮 간헐적 수요
          {counts.intermittent > 0 && (
            <span className="ml-1.5 text-xs bg-purple-100 text-purple-700 px-1.5 py-0.5 rounded-full">
              {counts.intermittent}
            </span>
          )}
        </button>
        <button className={tabCls('orders')} onClick={() => setSubTab('orders')}>
          📦 발주 현황
        </button>
      </div>

      {/* 발주계획 / 간헐적 수요 탭 */}
      {subTab !== 'orders' && (
        <div className="space-y-5">
          {/* 설정 */}
          <div className="bg-gray-50 border rounded-lg px-4 py-3 flex items-center gap-6 flex-wrap text-sm">
            <div className="flex items-center gap-2">
              <span className="font-medium text-gray-700">목표 재고</span>
              <input type="number" min={1} max={12} value={targetMonths}
                onChange={e => setTarget(Math.max(1, parseInt(e.target.value) || 1))}
                className="w-14 border rounded px-2 py-1 text-center" />
              <span className="text-gray-500">개월</span>
            </div>
            <div className="flex items-center gap-2 border-l pl-6">
              <span className="font-medium text-gray-700">간헐적 수요 기준</span>
              <input type="number" min={1} max={36} value={minMonths}
                onChange={e => setMinMonths(Math.max(1, parseInt(e.target.value) || 1))}
                className="w-14 border rounded px-2 py-1 text-center" />
              <span className="text-gray-500">개월 미만</span>
            </div>
            <span className="text-xs text-gray-400 border-l pl-6">
              판매 발생 월 수가 기준 미만이면 간헐적 수요로 분류 → 발주 대상 제외
            </span>
          </div>

          {/* 파일 업로드 */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <FileUploader
              label="재고 파일 (이카운트 수불부 또는 EMP 재고현황 — 자동 감지)"
              fileName={stockFile?.name ?? ''}
              onFile={setStockFile}
            />
            <div className="space-y-1">
              <FileUploader
                label="판매 현황 파일 (이카운트 판매 현황)"
                fileName={salesFile?.name ?? ''}
                onFile={handleSalesFile}
                optional={salesRows.length > 0}
              />
              {salesRows.length > 0 && (
                <p className="text-xs text-green-600">
                  ✅ {salesRows.length.toLocaleString()}행 로드됨
                  {!salesFile && ' (판매 현황 분석 탭 데이터)'}
                </p>
              )}
            </div>
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
                {counts.urgent       > 0 && <span className="text-red-600 font-medium">🔴 발주필요 {counts.urgent}</span>}
                {counts.warning      > 0 && <span className="text-yellow-600 font-medium">⚠️ 주의 {counts.warning}</span>}
                {counts.ok           > 0 && <span className="text-green-600">✅ 재고충분 {counts.ok}</span>}
                {counts.intermittent > 0 && <span className="text-purple-600">🔮 간헐적수요 {counts.intermittent}</span>}
                {counts.no_data      > 0 && <span className="text-gray-400">— 이력없음 {counts.no_data}</span>}
                <button
                  onClick={handleDownload}
                  className="ml-auto px-4 py-1.5 bg-green-600 text-white text-xs font-medium rounded hover:bg-green-700 transition"
                >
                  📥 Excel 다운로드
                </button>
              </div>

              {/* 발주계획 테이블 */}
              {subTab === 'plan' && (
                <div className="overflow-x-auto rounded-lg border">
                  <table className="min-w-full text-sm">
                    <thead>
                      <tr className="bg-[#1F3864] text-white text-xs">
                        <th className="px-3 py-2 text-left whitespace-nowrap">카테고리</th>
                        <th className="px-3 py-2 text-left whitespace-nowrap">품목코드</th>
                        <th className="px-3 py-2 text-left whitespace-nowrap">품목명</th>
                        <th className="px-3 py-2 text-left whitespace-nowrap">규격</th>
                        <th className="px-3 py-2 text-right whitespace-nowrap">현재고</th>
                        <th className="px-3 py-2 text-right whitespace-nowrap" title="전체 기간 중 판매 발생 월 수">판매월수</th>
                        <th className="px-3 py-2 text-right whitespace-nowrap">월간최고</th>
                        <th className="px-3 py-2 text-right whitespace-nowrap">커버리지</th>
                        <th className="px-3 py-2 text-right whitespace-nowrap">발주필요량</th>
                        <th className="px-3 py-2 text-center whitespace-nowrap">상태</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {rows.map((row, i) => {
                        const s  = rowStatus(row, targetMonths, minMonths)
                        const bg = STATUS_ROW_BG[s] || (i % 2 === 0 ? 'bg-white' : 'bg-gray-50')
                        return (
                          <tr key={row.code} className={bg}>
                            <td className="px-3 py-1.5 text-xs text-gray-600 whitespace-nowrap">{row.category}</td>
                            <td className="px-3 py-1.5 font-mono text-xs whitespace-nowrap">{row.code}</td>
                            <td className="px-3 py-1.5 whitespace-nowrap">{row.name}</td>
                            <td className="px-3 py-1.5 text-xs text-gray-500 whitespace-nowrap">{row.spec}</td>
                            <td className="px-3 py-1.5 text-right tabular-nums">{row.stock.toLocaleString()}</td>
                            <td className="px-3 py-1.5 text-right tabular-nums text-gray-500">
                              {row.salesMonths > 0 ? row.salesMonths : '—'}
                            </td>
                            <td className="px-3 py-1.5 text-right tabular-nums text-gray-600">
                              {row.peakMonthly > 0 ? row.peakMonthly.toLocaleString() : '—'}
                            </td>
                            <td className="px-3 py-1.5 text-right tabular-nums">
                              {row.peakMonthly > 0 ? row.coverage.toFixed(1) + '개월' : '—'}
                            </td>
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
              )}

              {/* 간헐적 수요 테이블 */}
              {subTab === 'intermittent' && (
                <IntermittentView rows={intermittentRows} years={years} minMonths={minMonths} />
              )}
            </div>
          )}
        </div>
      )}

      {/* 발주 현황 탭 — 항상 독립 접근 가능 */}
      {subTab === 'orders' && <ImportOrdersView />}
    </div>
  )
}

function IntermittentView({
  rows, years, minMonths,
}: {
  rows:      PlanRow[]
  years:     string[]
  minMonths: number
}) {
  if (rows.length === 0) {
    return (
      <div className="text-center py-10 text-gray-400 text-sm">
        간헐적 수요 품목이 없습니다. (기준: 판매월수 {minMonths}개월 미만)
      </div>
    )
  }

  return (
    <div className="space-y-2">
      <p className="text-xs text-gray-500">
        판매월수 {minMonths}개월 미만 품목 — 정기 발주 대상 제외, 별도 검토 필요
      </p>
      <div className="overflow-x-auto rounded-lg border">
        <table className="min-w-full text-sm">
          <thead>
            <tr className="bg-purple-900 text-white text-xs">
              <th className="px-3 py-2 text-left whitespace-nowrap">카테고리</th>
              <th className="px-3 py-2 text-left whitespace-nowrap">품목코드</th>
              <th className="px-3 py-2 text-left whitespace-nowrap">품목명</th>
              <th className="px-3 py-2 text-left whitespace-nowrap">규격</th>
              <th className="px-3 py-2 text-right whitespace-nowrap">현재고</th>
              <th className="px-3 py-2 text-right whitespace-nowrap">판매월수</th>
              <th className="px-3 py-2 text-right whitespace-nowrap">최고단월</th>
              {years.map(yr => (
                <th key={yr} className="px-3 py-2 text-center whitespace-nowrap">20{yr}년</th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {rows.map((row, i) => (
              <tr key={row.code} className={i % 2 === 0 ? 'bg-white' : 'bg-purple-50'}>
                <td className="px-3 py-1.5 text-xs text-gray-600 whitespace-nowrap">{row.category}</td>
                <td className="px-3 py-1.5 font-mono text-xs whitespace-nowrap">{row.code}</td>
                <td className="px-3 py-1.5 whitespace-nowrap">{row.name}</td>
                <td className="px-3 py-1.5 text-xs text-gray-500 whitespace-nowrap">{row.spec}</td>
                <td className="px-3 py-1.5 text-right tabular-nums">{row.stock.toLocaleString()}</td>
                <td className="px-3 py-1.5 text-right tabular-nums font-medium text-purple-700">
                  {row.salesMonths}
                </td>
                <td className="px-3 py-1.5 text-right tabular-nums font-medium">
                  {row.peakMonthly.toLocaleString()}
                </td>
                {years.map(yr => {
                  const s = row.byYear[yr]
                  return (
                    <td key={yr} className="px-3 py-1.5 text-center text-xs whitespace-nowrap">
                      {s ? (
                        <span>
                          <span className="text-gray-700 font-medium">{s.months}회</span>
                          <span className="text-gray-400 mx-1">/</span>
                          <span className="tabular-nums">{s.total.toLocaleString()}EA</span>
                        </span>
                      ) : (
                        <span className="text-gray-300">—</span>
                      )}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}
