import { useState } from 'react'
import { buildStep3Plan, parseSalesAggExcel } from '../lib/step3Core'
import type { CodeCableEntry, Step3Row } from '../lib/step3Core'
import type { Metadata, Inventory, AppSettings } from '../lib/types'
import type { SalesAggResult } from '../lib/aggregate/salesAgg'
import { buildStep3Workbook } from '../lib/output/writeStep3Excel'
import { downloadXlsx, today } from '../lib/download'
import FileUploader from './FileUploader'

interface Props {
  metadata:  Metadata
  inventory: Inventory
  settings:  AppSettings
  salesAgg?: SalesAggResult | null
}

export default function Step3({ metadata, inventory, settings, salesAgg }: Props) {
  const [logs, setLogs]                     = useState<string[]>([])
  const [running, setRunning]               = useState(false)
  const [done, setDone]                     = useState(false)
const [gaongFile, setGaongFile]           = useState<File | null>(null)
  const [salesAggFile, setSalesAggFile]     = useState<File | null>(null)
  const [rows, setRows]                     = useState<Step3Row[]>([])
  const [years, setYears]                   = useState<string[]>([])
  const [codeToCable, setCodeToCable]       = useState<Record<string, CodeCableEntry>>({})

  const run = async () => {
    if (!gaongFile) return alert('가공파일을 선택해주세요.')
    setRunning(true); setDone(false); setLogs([]); setRows([])
    try {
      const buf = await gaongFile.arrayBuffer()
      let kindCagr: Record<string, number> | undefined
      if (salesAggFile) {
        kindCagr = parseSalesAggExcel(await salesAggFile.arrayBuffer())
      }
      const result = buildStep3Plan(buf, metadata, inventory, settings.lead_time_default, settings.safety_stock_k ?? 1.5, kindCagr)
      setLogs(result.logs)
      setRows(result.rows)
      setYears(result.years)
      setCodeToCable(result.codeToCable)
      setDone(true)
    } catch (e) {
      setLogs([`❌ 오류: ${e}`])
    } finally {
      setRunning(false)
    }
  }

  const exportExcel = async () => {
    if (!rows.length) return
    const wb  = buildStep3Workbook(rows, years, settings.colors.main_header, settings.safety_stock_k ?? 1.5, metadata, inventory, salesAgg ?? undefined, codeToCable)
    const buf = await wb.xlsx.writeBuffer()
    downloadXlsx(buf as ArrayBuffer, `발주계획_${today()}.xlsx`)
  }

  const missingPn   = rows.filter(r => !r.품번).length

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>STEP 3 — 발주계획 생성</b>: STEP 1 가공파일을 업로드하면 자재별 <b>안전재고</b>와
        <b>발주 필요량</b>을 한눈에 파악하고 Excel로 다운로드합니다.
        <span className="text-gray-500 ml-1">안전재고 = 월평균 × (LT/30) × k &nbsp;(k={settings.safety_stock_k ?? 1.5}, ⚙️ 설정에서 변경)</span>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-start">
        <div className="md:col-span-2 space-y-2">
          <FileUploader
            label="① 가공파일.xlsx (STEP 1 결과)"
            fileName={gaongFile?.name ?? ''}
            onFile={setGaongFile}
          />
          <FileUploader
            label="② 판매분석.xlsx (STEP 2 결과, 선택) — 케이블 발주량에 판매CAGR 반영"
            fileName={salesAggFile?.name ?? ''}
            onFile={setSalesAggFile}
            optional
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
            onClick={() => { setDone(false); setLogs([]); setGaongFile(null); setSalesAggFile(null); setRows([]); setYears([]); setCodeToCable({}) }}
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

    </div>
  )
}


