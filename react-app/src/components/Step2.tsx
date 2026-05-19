import { useState, Fragment } from 'react'
import { parseSalesFile, parseProductionFile } from '../lib/parse/parseSales'
import { aggregateSales, forecastProduction } from '../lib/aggregate/salesAgg'
import { saveSalesAgg } from '../lib/supabase'
import { writeSalesAgg } from '../lib/output/writeSalesAgg'
import { downloadXlsx, today } from '../lib/download'
import type { SalesAggResult, SalesProductEntry } from '../lib/aggregate/salesAgg'
import FileUploader from './FileUploader'

interface Props {
  salesAgg:    SalesAggResult | null
  setSalesAgg: (s: SalesAggResult) => void
}

const KIND_LABEL: Record<string, string> = {
  'a1': 'A1 (SM)', 'a1-청': 'A1 청색', 'a1-녹': 'A1 녹색', 'a1-적': 'A1 적색', 'a1-자': 'A1 자색',
  'b3': 'B3 (SM)', 'om1': 'OM1 (MM)', 'om3': 'OM3 (MM)',
  'drop': 'DROP', 'pigtail': 'PIGTAIL', 'om1-pigtail': 'PIGTAIL (MM)',
  'a2': 'Optical Cable',
}

const kindLabel = (k: string) => KIND_LABEL[k] ?? k.toUpperCase()
const pct = (v: number) => `${Math.round(v * 100)}%`
const cagr = (v: number) => {
  const s = Math.round(v * 100)
  return s === 0 ? '0%' : s > 0 ? `+${s}%` : `${s}%`
}
const num = (v: number) => v === 0 ? '—' : v.toLocaleString()

export default function Step2({ salesAgg, setSalesAgg }: Props) {
  const [logs, setLogs]                     = useState<string[]>([])
  const [running, setRunning]               = useState(false)
  const [done, setDone]                     = useState(false)
  const [salesFile, setSalesFile]           = useState<File | null>(null)
  const [purchaseFile, setPurchaseFile]     = useState<File | null>(null)
  const [view, setView]                     = useState<'type' | 'product'>('type')

  const run = async () => {
    if (!salesFile) return alert('전체 판매량 파일을 선택해주세요.')
    setRunning(true); setDone(false); setLogs([])
    try {
      const newLogs: string[] = []
      const salesBuf = await salesFile.arrayBuffer()
      const salesRows = parseSalesFile(salesBuf, newLogs)
      const prodRows = purchaseFile
        ? parseProductionFile(await purchaseFile.arrayBuffer(), newLogs)
        : []
      const result = aggregateSales(salesRows, prodRows, newLogs)
      saveSalesAgg(result)
      setSalesAgg(result)
      setLogs(newLogs)
      setDone(true)
    } catch (e) {
      setLogs([`❌ 오류: ${e}`])
    } finally {
      setRunning(false)
    }
  }

  const downloadSalesAgg = async () => {
    if (!salesAgg) return
    downloadXlsx(await writeSalesAgg(salesAgg), `판매분석_${today()}.xlsx`)
  }

  const years      = salesAgg?.years ?? []
  const latestYr   = years[years.length - 1]
  const nextYrLabel = latestYr ? `${Number(latestYr) + 1}년` : ''
  const forecast   = salesAgg && latestYr
    ? forecastProduction(salesAgg, String(Number(latestYr) + 1), latestYr)
    : {}

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>STEP 2 — 판매 분석</b>: OJC 완제품의 <b>전체 판매량</b>과 <b>맥산 생산량</b>을 비교합니다.
        수입 완제품 = 전체 판매 − 맥산 생산. 판매 CAGR × 생산비중 = <b>B안 발주 예측</b>으로 STEP 3에 자동 반영됩니다.
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <FileUploader
          label="① 전체 판매량.xlsx 또는 판매현황 분석 파일 (자동 감지)"
          fileName={salesFile?.name ?? ''}
          onFile={setSalesFile}
        />
        <FileUploader
          label="② 구매관리(맥산).xlsx — 컬럼: 품목코드, 품목명, 입고일자, 수량"
          fileName={purchaseFile?.name ?? ''}
          onFile={setPurchaseFile}
          optional
        />
      </div>

      <hr className="border-gray-200" />

      <div className="flex gap-2 flex-wrap">
        <button
          onClick={run}
          disabled={running || !salesFile}
          className="flex-1 bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold py-2 px-4 rounded transition text-sm"
        >
          {running ? '⏳ 분석 중...' : '▶ STEP 2 실행 — 판매 분석'}
        </button>
        {salesAgg && (
          <button
            onClick={downloadSalesAgg}
            className="px-4 py-2 text-sm bg-[#2E75B6] hover:bg-[#1a5a9e] text-white font-semibold rounded transition"
          >
            📥 판매분석 Excel 저장
          </button>
        )}
        {(done || logs.length > 0) && (
          <button
            onClick={() => { setDone(false); setLogs([]); setSalesFile(null); setPurchaseFile(null) }}
            className="px-4 py-2 text-sm border border-gray-300 rounded text-gray-600 hover:bg-gray-100 transition"
          >
            🗑 초기화
          </button>
        )}
      </div>

      {done && (
        <div className="bg-[#D6F0D8] px-4 py-3 rounded-md text-sm font-semibold text-[#1a6a2a]">
          ✅ 판매 분석 완료 — STEP 3에 B안(판매CAGR × 생산비중) 예측이 자동 반영됩니다.
        </div>
      )}

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs leading-5 max-h-48 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {salesAgg && years.length > 0 && (
        <div className="space-y-3">
          <div className="flex gap-1 text-sm">
            {(['type', 'product'] as const).map(v => (
              <button
                key={v}
                onClick={() => setView(v)}
                className={`px-3 py-1.5 rounded font-medium border transition ${
                  view === v
                    ? 'bg-[#2E75B6] text-white border-[#2E75B6]'
                    : 'bg-white text-gray-600 border-gray-300 hover:border-[#2E75B6]'
                }`}
              >
                {v === 'type' ? '① 타입별 요약' : '② 품목별 상세'}
              </button>
            ))}
          </div>

          {view === 'type' && (
            <TypeSummary
              byType={salesAgg.byType}
              salesCagr={salesAgg.salesCagr}
              years={years}
              forecast={forecast}
              nextYrLabel={nextYrLabel}
            />
          )}
          {view === 'product' && (
            <ProductDetail byProduct={salesAgg.byProduct} years={years} />
          )}
        </div>
      )}
    </div>
  )
}

// ── 타입별 요약 테이블 ──────────────────────────────────────────
function TypeSummary({
  byType, salesCagr, years, forecast, nextYrLabel,
}: {
  byType:       SalesAggResult['byType']
  salesCagr:    SalesAggResult['salesCagr']
  years:        string[]
  forecast:     Record<string, number>
  nextYrLabel:  string
}) {
  const kinds = Object.keys(byType).sort()

  return (
    <div className="overflow-x-auto border border-gray-200 rounded">
      <table className="text-xs w-full whitespace-nowrap">
        <thead className="bg-gray-50">
          <tr>
            <th className="px-3 py-2 text-left border-b font-semibold text-gray-600" rowSpan={2}>타입</th>
            {years.map(yr => (
              <th key={yr} className="px-2 py-1 text-center border-b font-semibold text-gray-600" colSpan={3}>
                {yr}년
              </th>
            ))}
            <th className="px-3 py-2 text-center border-b font-semibold text-blue-700" rowSpan={2}>판매<br/>CAGR</th>
            {nextYrLabel && (
              <th className="px-3 py-2 text-center border-b font-semibold text-orange-700" rowSpan={2}>
                B안 예측<br/>{nextYrLabel}
              </th>
            )}
          </tr>
          <tr>
            {years.map(yr => (
              <Fragment key={yr}>
                <th className="px-2 py-1 text-right border-b text-gray-500 font-normal">판매</th>
                <th className="px-2 py-1 text-right border-b text-gray-500 font-normal">생산</th>
                <th className="px-2 py-1 text-right border-b text-gray-500 font-normal">생산%</th>
              </Fragment>
            ))}
          </tr>
        </thead>
        <tbody>
          {kinds.map((kind, i) => {
            const entry = byType[kind]
            const c = salesCagr[kind] ?? 0
            const f = forecast[kind]
            return (
              <tr key={kind} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                <td className="px-3 py-1.5 font-medium text-gray-800">{kindLabel(kind)}</td>
                {years.map(yr => {
                  const e = entry[yr]
                  return (
                    <Fragment key={yr}>
                      <td className="px-2 py-1.5 text-right">{num(e?.sales ?? 0)}</td>
                      <td className="px-2 py-1.5 text-right text-blue-700">{num(e?.production ?? 0)}</td>
                      <td className="px-2 py-1.5 text-right text-gray-500">
                        {e?.sales ? pct(e.ratio) : '—'}
                      </td>
                    </Fragment>
                  )
                })}
                <td className={`px-3 py-1.5 text-center font-semibold ${c >= 0 ? 'text-blue-700' : 'text-red-600'}`}>
                  {cagr(c)}
                </td>
                {nextYrLabel && (
                  <td className="px-3 py-1.5 text-right font-semibold text-orange-700">
                    {f ? f.toLocaleString() : '—'}
                  </td>
                )}
              </tr>
            )
          })}
        </tbody>
      </table>
      <div className="px-3 py-2 text-xs text-gray-400 border-t">
        ※ B안 예측 = 최근년 판매 × (1 + 판매CAGR) × 최근년 생산비중 | 단위: 완제품 수량(EA) | STEP 3에서 자재량으로 변환
      </div>
    </div>
  )
}

// ── 품목별 상세 테이블 ──────────────────────────────────────────
function ProductDetail({ byProduct, years }: { byProduct: SalesProductEntry[]; years: string[] }) {
  return (
    <div className="overflow-x-auto border border-gray-200 rounded">
      <table className="text-xs w-full">
        <thead className="bg-gray-50 sticky top-0">
          <tr>
            <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">품목코드</th>
            <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">품목명</th>
            <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">타입</th>
            {years.map(yr => (
              <th key={yr} className="px-3 py-2 text-right border-b font-semibold text-gray-600">{yr}년 판매</th>
            ))}
            {years.map(yr => (
              <th key={`p-${yr}`} className="px-3 py-2 text-right border-b font-semibold text-blue-600">{yr}년 생산</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {byProduct.map((p, i) => (
            <tr key={p.code || p.name} className={i % 2 === 0 ? 'bg-white hover:bg-gray-50' : 'bg-gray-50 hover:bg-gray-100'}>
              <td className="px-3 py-1.5 font-mono text-gray-500 text-xs">{p.code || '—'}</td>
              <td className="px-3 py-1.5 max-w-xs truncate" title={p.name}>{p.name}</td>
              <td className="px-3 py-1.5 text-gray-500">{kindLabel(p.kind)}</td>
              {years.map(yr => (
                <td key={yr} className="px-3 py-1.5 text-right">{num(p.byYear[yr]?.sales ?? 0)}</td>
              ))}
              {years.map(yr => (
                <td key={`p-${yr}`} className="px-3 py-1.5 text-right text-blue-700">{num(p.byYear[yr]?.production ?? 0)}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
