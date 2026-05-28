import { useState, useMemo } from 'react'
import ExcelJS from 'exceljs'
import FileUploader from './FileUploader'
import {
  parseMaeksanCost, parseFlcCost, parseContractItems,
} from '../lib/parse/parseCostFile'
import type { CostEntry } from '../lib/parse/parseCostFile'
import { classifyOjc } from '../lib/ojcFilter'
import type { DetailedSalesRow } from '../lib/parse/parseDetailedSales'
import { downloadXlsx, today } from '../lib/download'
import {
  getReco, getSignalMargin, marginSignal, applyCost, 적용원가명,
} from '../lib/profitabilityEngine'
import type { ProfitRow } from '../lib/profitabilityEngine'

function pct(v: number | null): string {
  if (v === null) return '—'
  return v.toFixed(1) + '%'
}

function won(v: number | null): string {
  if (v === null) return '—'
  return Math.round(v).toLocaleString()
}

export default function ProfitabilityAnalysis({ salesRows }: { salesRows: DetailedSalesRow[] }) {
  const [maeksanFile,       setMaeksanFile]       = useState<File | null>(null)
  const [flcFile,           setFlcFile]           = useState<File | null>(null)
  const [contractFile,      setContractFile]      = useState<File | null>(null)
  const [rows,              setRows]              = useState<ProfitRow[]>([])
  const [loading,           setLoading]           = useState(false)
  const [error,             setError]             = useState('')
  const [hasBothCost,       setHasBothCost]       = useState(false)
  const [showLegend,        setShowLegend]        = useState(false)
  // 맥산 생산 가중치: 0 = 없음, N > 0 = 맥산이 FLC보다 N%까지 비싸도 생산 유지
  const [bufferPct,         setBufferPct]         = useState(0)
  // 최소 마진율: 이 미만이면 철수검토
  const [minMarginPct,      setMinMarginPct]      = useState(0)

  const importThreshold = bufferPct > 0 ? 100 + bufferPct : 100

  // code → { totalPrice, totalQty, yearSet }
  const salesMap = useMemo(() => {
    const m = new Map<string, { totalPrice: number; totalQty: number; years: Set<string> }>()
    for (const r of salesRows) {
      if (!r.code || r.qty <= 0) continue
      const e = m.get(r.code) ?? { totalPrice: 0, totalQty: 0, years: new Set<string>() }
      e.totalPrice += r.price
      e.totalQty   += r.qty
      e.years.add(r.year)
      m.set(r.code, e)
    }
    return m
  }, [salesRows])

  async function handleRun() {
    if (!maeksanFile && !flcFile) {
      setError('맥산 원가 파일 또는 FLC 원가 파일 중 하나 이상을 업로드하세요.')
      return
    }
    setLoading(true); setError('')
    try {
      const [maeksanMap, flcMap, contractMap] = await Promise.all([
        maeksanFile ? maeksanFile.arrayBuffer().then(parseMaeksanCost) : Promise.resolve(new Map<string, CostEntry>()),
        flcFile     ? flcFile.arrayBuffer().then(parseFlcCost)         : Promise.resolve(new Map<string, CostEntry>()),
        contractFile? contractFile.arrayBuffer().then(parseContractItems) : Promise.resolve(new Map<string, { code: string; name: string; spec: string; 단가: number }>()),
      ])
      setHasBothCost(maeksanMap.size > 0 && flcMap.size > 0)

      const allCodes = new Set([...maeksanMap.keys(), ...flcMap.keys()])
      const result: ProfitRow[] = []

      for (const code of allCodes) {
        const m = maeksanMap.get(code)
        const f = flcMap.get(code)
        const name     = m?.name ?? f?.name ?? code
        const category = classifyOjc(name) ?? '기타'

        const maeksanApplied = m ? applyCost(m, category) : null
        const flcApplied     = f ? applyCost(f, category) : null

        // 판매가: 계약 우선 → 평균
        let 판매가: number | null = null
        let isContract = false
        const contract = contractMap.get(code)
        if (contract && contract.단가 > 0) {
          판매가 = contract.단가; isContract = true
        } else {
          const s = salesMap.get(code)
          if (s && s.totalQty > 0 && s.totalPrice > 0)
            판매가 = s.totalPrice / s.totalQty
        }

        const s = salesMap.get(code)
        const yearCount  = s?.years.size ?? 0
        const avgAnnualQty = s && yearCount > 0 ? s.totalQty / yearCount : 0

        const 마진율맥산 = (판매가 !== null && maeksanApplied !== null && 판매가 > 0)
          ? (판매가 - maeksanApplied) / 판매가 * 100 : null
        const 마진율FLC  = (판매가 !== null && flcApplied !== null && 판매가 > 0)
          ? (판매가 - flcApplied) / 판매가 * 100 : null

        const 연간이익맥산 = (판매가 !== null && maeksanApplied !== null)
          ? (판매가 - maeksanApplied) * avgAnnualQty : null
        const 연간이익FLC  = (판매가 !== null && flcApplied !== null)
          ? (판매가 - flcApplied) * avgAnnualQty : null

        result.push({
          code, name, category,
          maeksanApplied, flcApplied, 판매가, isContract, avgAnnualQty,
          마진율맥산, 마진율FLC, 연간이익맥산, 연간이익FLC,
        })
      }

      result.sort((a, b) => {
        if (a.category !== b.category) return a.category.localeCompare(b.category)
        const ia = a.연간이익맥산 ?? a.연간이익FLC ?? -Infinity
        const ib = b.연간이익맥산 ?? b.연간이익FLC ?? -Infinity
        return ib - ia
      })

      setRows(result)
    } catch (e) {
      setError(String(e))
    } finally {
      setLoading(false)
    }
  }

  async function handleDownload() {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('수익성 분석')

    const headers = [
      '카테고리', '품목코드', '품목명',
      '맥산 원가', 'FLC 원가', '판매가', '판매가 구분',
      '맥산 마진율', 'FLC 마진율',
      '연간판매량', '맥산 연간이익', 'FLC 연간이익', '권고',
    ]
    ws.addRow(headers)
    ws.getRow(1).eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } }
      cell.font      = { bold: true, size: 9, color: { argb: 'FFFFFFFF' }, name: 'Arial' }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
      cell.border    = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
    })
    ws.getRow(1).height = 20

    for (const row of rows) {
      const reco = getReco(row, importThreshold, minMarginPct)
      ws.addRow([
        row.category, row.code, row.name,
        row.maeksanApplied ?? '',
        row.flcApplied ?? '',
        row.판매가 ?? '',
        row.isContract ? '계약' : '평균',
        row.마진율맥산 !== null ? parseFloat(row.마진율맥산.toFixed(1)) : '',
        row.마진율FLC  !== null ? parseFloat(row.마진율FLC.toFixed(1))  : '',
        row.avgAnnualQty ? Math.round(row.avgAnnualQty) : '',
        row.연간이익맥산 !== null ? Math.round(row.연간이익맥산) : '',
        row.연간이익FLC  !== null ? Math.round(row.연간이익FLC)  : '',
        `${marginSignal(getSignalMargin(row, reco.label))} ${reco.label}`.trim(),
      ]).eachCell(cell => {
        cell.font   = { size: 9, name: 'Arial' }
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
      })
    }

    ws.columns = [
      { width: 16 }, { width: 22 }, { width: 36 },
      { width: 12 }, { width: 12 }, { width: 12 }, { width: 10 },
      { width: 12 }, { width: 12 },
      { width: 12 }, { width: 14 }, { width: 14 }, { width: 14 },
    ]
    for (let c = 4; c <= 13; c++) ws.getColumn(c).alignment = { horizontal: 'right' }

    downloadXlsx(await wb.xlsx.writeBuffer() as ArrayBuffer, `수익성분석_${today()}.xlsx`)
  }

  const counts = useMemo(() => {
    let 맥산생산권고 = 0, 맥산생산가능 = 0, 수입권고 = 0, 철수검토 = 0, 검토필요 = 0
    for (const r of rows) {
      const { label } = getReco(r, importThreshold, minMarginPct)
      if (label.includes('✅ 맥산 생산')) 맥산생산권고++
      else if (label.includes('⚖ 맥산 생산')) 맥산생산가능++
      else if (label.includes('수입 권고')) 수입권고++
      else if (label.includes('철수검토')) 철수검토++
      else if (label.includes('검토필요')) 검토필요++
    }
    return { 맥산생산권고, 맥산생산가능, 수입권고, 철수검토, 검토필요 }
  }, [rows, importThreshold, minMarginPct])

  return (
    <div className="space-y-5">
      {/* 설명 */}
      <div className="text-sm text-gray-500">
        맥산(국내 생산) · FLC(수입) 원가 파일과 계약 단가 파일을 업로드하면 품목별 마진율 및 생산 방식 권고를 제공합니다.
        <br />
        <span className="text-xs text-gray-400">LG향 → 생산원가 적용 · KT향 → 표준원가 적용 · 판매가: 계약 품목은 계약 단가, 나머지는 판매 이력 평균</span>
      </div>

      {/* 파일 업로드 */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <FileUploader
          label="맥산 생산원가 파일"
          fileName={maeksanFile?.name ?? ''}
          onFile={setMaeksanFile}
          optional
        />
        <FileUploader
          label="FLC 수입원가 파일"
          fileName={flcFile?.name ?? ''}
          onFile={setFlcFile}
          optional
        />
        <FileUploader
          label="계약 단가 파일 (선택)"
          fileName={contractFile?.name ?? ''}
          onFile={setContractFile}
          optional
        />
      </div>

      {salesRows.length === 0 && (
        <p className="text-xs text-amber-600">
          ⚠ 판매 데이터 없음 — 판매 현황 분석 탭에서 먼저 데이터를 로드하면 비계약 품목의 평균 판매가를 자동 계산합니다.
        </p>
      )}

      {/* 최소 마진율 설정 */}
      <div className="bg-gray-50 border rounded-lg px-4 py-3 flex items-center gap-4 flex-wrap text-sm">
        <span className="font-medium text-gray-700">📊 마진율 기준</span>
        <div className="flex items-center gap-1.5">
          <span className="text-gray-500">최소</span>
          <input
            type="number"
            min={0}
            max={50}
            step={1}
            value={minMarginPct}
            onChange={e => setMinMarginPct(Math.min(50, Math.max(0, parseInt(e.target.value) || 0)))}
            className="w-14 border rounded px-2 py-1 text-center font-mono"
          />
          <span className="text-gray-500">%</span>
        </div>
        <div className="flex items-center gap-2 text-xs border-l pl-4">
          <span className="px-2 py-0.5 rounded-full bg-red-100 text-red-700">
            ⚠ 철수검토 &lt; {minMarginPct}%
          </span>
          <span className="text-gray-400">→</span>
          <span className="px-2 py-0.5 rounded-full bg-yellow-100 text-yellow-700">
            🔍 검토필요 {minMarginPct}~{Math.max(20, minMarginPct * 3)}%
          </span>
          <span className="text-gray-400">→</span>
          <span className="px-2 py-0.5 rounded-full bg-green-100 text-green-700">
            ✅ 맥산 생산 권고 ≥ {Math.max(20, minMarginPct * 3)}%
          </span>
        </div>
      </div>

      {/* 맥산 생산 가중치 설정 */}
      <div className="bg-gray-50 border rounded-lg px-4 py-3 flex items-center gap-4 flex-wrap text-sm">
        <span className="font-medium text-gray-700">📦 맥산 생산 가중치</span>
        <div className="flex items-center gap-1.5">
          <input
            type="number"
            min={0}
            max={30}
            step={1}
            value={bufferPct}
            onChange={e => setBufferPct(Math.min(30, Math.max(0, parseInt(e.target.value) || 0)))}
            className="w-14 border rounded px-2 py-1 text-center font-mono"
          />
          <span className="text-gray-500">%</span>
        </div>
        <span className="text-xs text-gray-500 border-l pl-4">
          {bufferPct > 0
            ? `맥산이 FLC보다 ${bufferPct}%까지 비싸도 국내생산 유지`
            : '0% — FLC가 1원이라도 저렴하면 수입 권고'}
        </span>
      </div>

      <div className="flex items-center gap-4 flex-wrap">
        <button
          onClick={handleRun}
          disabled={loading || (!maeksanFile && !flcFile)}
          className="px-6 py-2 bg-blue-600 text-white text-sm font-semibold rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition"
        >
          {loading ? '분석 중…' : '📊 수익성 분석 실행'}
        </button>
        <button
          onClick={() => setShowLegend(v => !v)}
          className="text-xs text-gray-500 hover:text-gray-800 underline underline-offset-2 transition"
        >
          {showLegend ? '▲ 권고 기준 닫기' : '▽ 권고 기준 안내'}
        </button>
      </div>

      {showLegend && (
        <div className="text-left bg-blue-50 border border-blue-200 rounded-lg p-5 text-sm space-y-5">
          <p className="font-semibold text-[#1F3864] text-base">📋 권고 기준 안내</p>

          {/* 신호등 */}
          <div className="space-y-2">
            <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide">신호등 — 마진율 수준</p>
            <div className="grid grid-cols-1 gap-1 text-xs text-gray-700">
              <div className="flex items-center gap-2">
                <span className="text-base leading-none">🟢</span>
                <span><b>20% 이상</b> — 수익성 양호</span>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-base leading-none">🟡</span>
                <span><b>0 ~ 20%</b> — 수익성 낮음, 영업 부서 확인 필요</span>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-base leading-none">🔴</span>
                <span><b>0% 미만</b> — 원가가 판매가 초과</span>
              </div>
            </div>
            <p className="text-xs text-gray-400 pl-1">
              수입 권고 → FLC 마진율 기준 &nbsp;·&nbsp;
              맥산 생산 권고·검토필요 → 맥산 마진율 기준 &nbsp;·&nbsp;
              철수검토 → 양쪽 중 더 나은 값
            </p>
          </div>

          {/* 권고 유형 */}
          <div className="space-y-2">
            <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide">권고 유형</p>
            <div className="space-y-2.5 text-xs text-gray-700">
              {[
                { bg: 'bg-green-100',  text: 'text-green-700',  label: '✅ 맥산 생산 권고', desc: '맥산 마진 ≥ 생산유지기준이고, 맥산 원가가 FLC보다 저렴하거나 동등한 경우' },
                { bg: 'bg-orange-100', text: 'text-orange-700', label: '⚖ 맥산 생산 가능',  desc: '맥산 마진은 기준 이상이나 원가가 FLC보다 비쌈 — 가중치 버퍼 덕분에 수입 권고를 피한 케이스. 가중치 0%면 수입 권고로 전환' },
                { bg: 'bg-blue-100',   text: 'text-blue-700',   label: '📦 수입 권고',       desc: 'FLC 원가가 유리하거나 맥산 마진이 최소 마진율 미달인 경우' },
                { bg: 'bg-yellow-100', text: 'text-yellow-700', label: '🔍 검토필요',         desc: '마진이 최소 기준을 넘지만 생산유지기준 미만 — 영업 부서 협의 권장' },
                { bg: 'bg-red-100',    text: 'text-red-700',    label: '⚠ 철수검토',         desc: '모든 원가 옵션의 마진이 최소 마진율 미달 — 단가 재협상 또는 판매 중단 건의' },
              ].map(({ bg, text, label, desc }) => (
                <div key={label} className="flex items-start gap-3">
                  <span className={`px-2 py-0.5 rounded-full whitespace-nowrap shrink-0 ${bg} ${text}`}>{label}</span>
                  <span className="text-gray-600 leading-relaxed">{desc}</span>
                </div>
              ))}
            </div>
          </div>

          {/* 설정값 */}
          <div className="space-y-2 border-t border-blue-200 pt-4">
            <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide">설정값 & 공식</p>
            <dl className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-1.5 text-xs">
              <dt className="font-semibold text-gray-700 whitespace-nowrap">최소 마진율</dt>
              <dd className="text-gray-600">이 값 미만이면 무조건 철수검토. 기본 0% (모든 양의 마진 허용)</dd>
              <dt className="font-semibold text-gray-700 whitespace-nowrap">생산유지기준</dt>
              <dd className="text-gray-600">max(20%, 최소마진×3) — 자동 계산. 이 기준 이상이어야 생산/수입 권고</dd>
              <dt className="font-semibold text-gray-700 whitespace-nowrap">맥산 생산 가중치</dt>
              <dd className="text-gray-600">0% = 없음. N% 입력 시 맥산이 FLC보다 N% 이내로 비싸도 생산 유지</dd>
              <dt className="font-semibold text-gray-700 whitespace-nowrap">판매가</dt>
              <dd className="text-gray-600">계약 단가 파일 우선 → 없으면 판매 이력 평균. KT향 → 표준원가, LG향·기타 → 생산원가</dd>
              <dt className="font-semibold text-gray-700 whitespace-nowrap">마진율 공식</dt>
              <dd className="text-gray-600"><code>(판매가 − 원가) / 판매가 × 100</code></dd>
            </dl>
          </div>
        </div>
      )}

      {error && <p className="text-sm text-red-600">{error}</p>}

      {rows.length > 0 && (
        <div className="space-y-3">
          {/* 요약 */}
          <div className="flex items-center gap-4 bg-gray-50 border rounded-lg px-4 py-2.5 flex-wrap text-sm">
            <span className="font-medium text-gray-700">총 {rows.length}개 품목</span>
            {counts.맥산생산권고 > 0 && <span className="text-green-700">✅ 맥산 생산 권고 {counts.맥산생산권고}</span>}
            {counts.맥산생산가능 > 0 && <span className="text-orange-700">⚖ 맥산 생산 가능 {counts.맥산생산가능}</span>}
            {counts.수입권고 > 0 && <span className="text-blue-700">📦 수입 권고 {counts.수입권고}</span>}
            {counts.검토필요 > 0 && <span className="text-yellow-700">🔍 검토필요 {counts.검토필요}</span>}
            {counts.철수검토 > 0 && <span className="text-red-700 font-medium">⚠ 철수검토 {counts.철수검토}</span>}
            {hasBothCost && (
              <span className="text-xs text-gray-400 border-l pl-4">
                {bufferPct > 0 ? `가중치 ${bufferPct}% 적용 중` : '가중치 미적용'}
              </span>
            )}
            <button
              onClick={handleDownload}
              className="ml-auto px-4 py-1.5 bg-green-600 text-white text-xs font-medium rounded hover:bg-green-700 transition"
            >
              📥 Excel 다운로드
            </button>
          </div>

          {/* 테이블 */}
          <div className="overflow-x-auto rounded-lg border">
            <table className="min-w-full text-sm">
              <thead>
                <tr className="bg-[#1F3864] text-white text-xs">
                  <th className="px-3 py-2 text-left whitespace-nowrap">카테고리</th>
                  <th className="px-3 py-2 text-left whitespace-nowrap">품목코드</th>
                  <th className="px-3 py-2 text-left whitespace-nowrap">품목명</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap" title="LG향=생산원가, KT향=표준원가">
                    맥산 원가
                  </th>
                  <th className="px-3 py-2 text-right whitespace-nowrap" title="LG향=생산원가, KT향=표준원가">
                    FLC 원가
                  </th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">판매가</th>
                  <th className="px-3 py-2 text-center whitespace-nowrap">구분</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">맥산 마진율</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">FLC 마진율</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">연간판매량</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">맥산 연간이익</th>
                  <th className="px-3 py-2 text-right whitespace-nowrap">FLC 연간이익</th>
                  <th className="px-3 py-2 text-center whitespace-nowrap">권고</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {rows.map((row, i) => {
                  const reco = getReco(row, importThreshold, minMarginPct)
                  const sig  = marginSignal(getSignalMargin(row, reco.label))
                  const rowBg = i % 2 === 0 ? 'bg-white' : 'bg-gray-50'
                  const catLabel = 적용원가명(row.category)
                  return (
                    <tr key={row.code} className={rowBg}>
                      <td className="px-3 py-1.5 text-xs text-gray-600 whitespace-nowrap">{row.category}</td>
                      <td className="px-3 py-1.5 font-mono text-xs whitespace-nowrap">{row.code}</td>
                      <td className="px-3 py-1.5 whitespace-nowrap">{row.name}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">
                        {row.maeksanApplied !== null ? (
                          <span title={catLabel}>{won(row.maeksanApplied)}</span>
                        ) : (
                          <span className="text-gray-300">—</span>
                        )}
                      </td>
                      <td className="px-3 py-1.5 text-right tabular-nums">
                        {row.flcApplied !== null ? (
                          <span title={catLabel}>{won(row.flcApplied)}</span>
                        ) : (
                          <span className="text-gray-300">—</span>
                        )}
                      </td>
                      <td className="px-3 py-1.5 text-right tabular-nums">
                        {row.판매가 !== null ? won(row.판매가) : <span className="text-gray-300">—</span>}
                      </td>
                      <td className="px-3 py-1.5 text-center">
                        {row.판매가 !== null && (
                          <span className={`text-xs px-1.5 py-0.5 rounded-full ${
                            row.isContract ? 'bg-blue-100 text-blue-700' : 'bg-gray-100 text-gray-600'
                          }`}>
                            {row.isContract ? '계약' : '평균'}
                          </span>
                        )}
                      </td>
                      <td className="px-3 py-1.5 text-right tabular-nums">
                        <MarginCell v={row.마진율맥산} />
                      </td>
                      <td className="px-3 py-1.5 text-right tabular-nums">
                        <MarginCell v={row.마진율FLC} />
                      </td>
                      <td className="px-3 py-1.5 text-right tabular-nums text-gray-600">
                        {row.avgAnnualQty > 0 ? Math.round(row.avgAnnualQty).toLocaleString() : <span className="text-gray-300">—</span>}
                      </td>
                      <td className="px-3 py-1.5 text-right tabular-nums">
                        {row.연간이익맥산 !== null
                          ? <span className={row.연간이익맥산 < 0 ? 'text-red-600' : ''}>{won(row.연간이익맥산)}</span>
                          : <span className="text-gray-300">—</span>}
                      </td>
                      <td className="px-3 py-1.5 text-right tabular-nums">
                        {row.연간이익FLC !== null
                          ? <span className={row.연간이익FLC < 0 ? 'text-red-600' : ''}>{won(row.연간이익FLC)}</span>
                          : <span className="text-gray-300">—</span>}
                      </td>
                      <td className="px-3 py-1.5 text-center">
                        <div className="flex items-center gap-1 justify-center">
                          {sig && <span className="text-sm leading-none">{sig}</span>}
                          <span className={`text-xs px-2 py-0.5 rounded-full whitespace-nowrap ${reco.bg} ${reco.text}`}>
                            {reco.label}
                          </span>
                        </div>
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

function MarginCell({ v }: { v: number | null }) {
  if (v === null) return <span className="text-gray-300">—</span>
  const color = v < 0 ? 'text-red-600 font-medium' : v < 10 ? 'text-yellow-700' : v >= 20 ? 'text-green-700 font-medium' : ''
  return <span className={color}>{pct(v)}</span>
}
