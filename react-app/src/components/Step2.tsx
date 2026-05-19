import { useState } from 'react'
import { parseSalesFile, parseProductionFile } from '../lib/parse/parseSales'
import { aggregateSales } from '../lib/aggregate/salesAgg'
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
const num = (v: number) => v === 0 ? '—' : v.toLocaleString()

export default function Step2({ salesAgg, setSalesAgg }: Props) {
  const [logs, setLogs]                     = useState<string[]>([])
  const [running, setRunning]               = useState(false)
  const [done, setDone]                     = useState(false)
  const [salesFile, setSalesFile]           = useState<File | null>(null)
  const [purchaseFile, setPurchaseFile]     = useState<File | null>(null)

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

  const years = salesAgg?.years ?? []

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
        <ProductDetail byProduct={salesAgg.byProduct} years={years} />
      )}
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
