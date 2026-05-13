import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import { saveSalesAnalysis } from '../lib/supabase'
import { classifyOjc } from '../lib/ojcFilter'
import type { SalesAnalysis, SalesItem } from '../lib/types'

interface Props {
  sales: SalesAnalysis
  setSales: (s: SalesAnalysis) => void
}

export default function Step2({ sales, setSales }: Props) {
  const [logs, setLogs] = useState<string[]>([])
  const [running, setRunning] = useState(false)
  const [done, setDone] = useState(false)
  const salesRef = useRef<HTMLInputElement>(null)
  const purchaseRef = useRef<HTMLInputElement>(null)
  const [salesName, setSalesName] = useState('')
  const [purchaseName, setPurchaseName] = useState('')

  const nSales = Object.values(sales).filter(v =>
    ['23','24','25'].some(yr => (v[yr as '23'|'24'|'25']?.sales ?? 0) > 0)
  ).length

  const run = async () => {
    const salesFile = salesRef.current?.files?.[0]
    if (!salesFile) return alert('판매량 파일을 선택해주세요.')
    setRunning(true); setDone(false); setLogs([])
    try {
      const newLogs: string[] = []
      const salesBuf = await salesFile.arrayBuffer()
      const wb = XLSX.read(salesBuf, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })
      newLogs.push(`판매량 파일 로드 — ${rows.length.toLocaleString()}행`)

      const salesBy: Record<string, { 품목명: string; [yr: string]: number | string }> = {}
      for (const row of rows) {
        const name = String(row['품목명'] ?? '')
        const ojcType = classifyOjc(name)
        if (!ojcType) continue
        const code = String(row['품목코드'] ?? '').trim()
        if (!code) continue
        let qty = 0
        try { qty = parseInt(String(row['수량'])) || 0 } catch {}
        if (qty <= 0) continue
        let yr = ''
        const yrRaw = row['년']
        if (yrRaw) {
          const yrS = String(parseInt(String(yrRaw)))
          yr = yrS.length === 4 ? yrS.slice(2) : yrS.slice(-2)
        }
        if (!['23','24','25'].includes(yr)) continue
        if (!salesBy[code]) salesBy[code] = { 품목명: name }
        salesBy[code][yr] = (Number(salesBy[code][yr] ?? 0)) + qty
      }
      newLogs.push(`OJC 분류 완료 — ${Object.keys(salesBy).length.toLocaleString()}개 품목`)

      const purchaseFile = purchaseRef.current?.files?.[0]
      const prodBy: Record<string, { [yr: string]: number }> = {}
      if (purchaseFile) {
        const purchBuf = await purchaseFile.arrayBuffer()
        const pwb = XLSX.read(purchBuf, { type: 'array' })
        const pws = pwb.Sheets[pwb.SheetNames[0]]
        const prows = XLSX.utils.sheet_to_json<Record<string, unknown>>(pws, { defval: null })
        for (const row of prows) {
          const code = String(row['품목코드'] ?? '').trim()
          if (!code) continue
          let qty = 0
          try { qty = parseInt(String(row['수량'])) || 0 } catch {}
          if (qty <= 0) continue
          let yr = ''
          const dateRaw = String(row['입고일자'] ?? '')
          const m = dateRaw.replace(/\s*-\d+\s*$/, '').match(/^(\d{2,4})/)
          if (m) yr = m[1].length === 4 ? m[1].slice(2) : m[1]
          if (!['23','24','25'].includes(yr)) continue
          if (!prodBy[code]) prodBy[code] = {}
          prodBy[code][yr] = (prodBy[code][yr] ?? 0) + qty
        }
        newLogs.push(`구매관리(맥산) 파일 로드 완료`)
      }

      const analysis: SalesAnalysis = {}
      for (const [code, sd] of Object.entries(salesBy)) {
        const item: SalesItem = { 품목명: sd['품목명'] as string, '23': {sales:0,production:0,ratio:0}, '24': {sales:0,production:0,ratio:0}, '25': {sales:0,production:0,ratio:0} }
        for (const yr of ['23','24','25'] as const) {
          const s = Number(sd[yr] ?? 0)
          const p = Number(prodBy[code]?.[yr] ?? 0)
          item[yr] = { sales: s, production: p, ratio: p > 0 ? s / p : 0 }
        }
        analysis[code] = item
      }

      saveSalesAnalysis(analysis)
      setSales(analysis)
      newLogs.push(`✅ 판매 분석 완료 — ${Object.keys(analysis).length.toLocaleString()}개 품목`)
      setLogs(newLogs)
      setDone(true)
    } catch (e) {
      setLogs([`❌ 오류: ${e}`])
    } finally {
      setRunning(false)
    }
  }

  return (
    <div className="space-y-4">
      <div className="bg-[#E0F2FE] border-l-4 border-[#0EA5E9] px-4 py-3 rounded">
        <b>STEP 2 — 판매 분석</b>: 전체 판매량 파일과 구매관리(맥산) 파일을 분석하여
        품목별 판매/생산 비중을 계산합니다. STEP 3 수요 기반 분석에 자동 반영됩니다.
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="space-y-1.5">
          <label className="block text-sm font-semibold text-gray-700">판매량 파일 <span className="text-red-500">*필수</span></label>
          <div className="bg-white border border-gray-300 rounded-lg px-4 py-3 flex items-center gap-3">
            <label className="flex items-center gap-1.5 bg-white border border-gray-400 rounded px-3 py-1.5 cursor-pointer hover:bg-gray-50 transition text-sm font-medium text-gray-700 shrink-0">
              <span>↑</span> Upload
              <input ref={salesRef} type="file" accept=".xlsx,.xls" className="hidden"
                onChange={e => setSalesName(e.target.files?.[0]?.name ?? '')} />
            </label>
            <span className="text-sm text-gray-500 truncate">
              {salesName ? <span className="text-green-600 font-medium">✓ {salesName}</span> : '200MB per file • XLSX'}
            </span>
          </div>
        </div>
        <div className="space-y-1.5">
          <label className="block text-sm font-semibold text-gray-700">구매관리(맥산) 파일 <span className="text-gray-400 font-normal">선택</span></label>
          <div className="bg-white border border-gray-300 rounded-lg px-4 py-3 flex items-center gap-3">
            <label className="flex items-center gap-1.5 bg-white border border-gray-400 rounded px-3 py-1.5 cursor-pointer hover:bg-gray-50 transition text-sm font-medium text-gray-700 shrink-0">
              <span>↑</span> Upload
              <input ref={purchaseRef} type="file" accept=".xlsx,.xls" className="hidden"
                onChange={e => setPurchaseName(e.target.files?.[0]?.name ?? '')} />
            </label>
            <span className="text-sm text-gray-500 truncate">
              {purchaseName ? <span className="text-green-600 font-medium">✓ {purchaseName}</span> : '200MB per file • XLSX'}
            </span>
          </div>
        </div>
      </div>

      <button
        onClick={run}
        disabled={running || !salesName}
        className="w-full bg-[#0EA5E9] hover:bg-[#0284C7] disabled:bg-gray-300 text-white font-bold py-2.5 rounded transition"
      >
        {running ? '⏳ 분석 중...' : '▶ STEP 2 실행'}
      </button>

      {done && nSales > 0 && (
        <div className="bg-[#e8f5e9] border-l-4 border-[#1a7a3c] px-4 py-2 rounded text-sm font-semibold text-[#1a7a3c]">
          ✅ 판매 분석 완료 — <b>{nSales.toLocaleString()}</b>개 품목 저장됨. STEP 3에서 수요 기반 분석이 자동 포함됩니다.
        </div>
      )}

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs space-y-0.5 max-h-48 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {nSales > 0 && !done && (
        <div className="bg-green-50 border border-green-300 rounded-lg overflow-hidden">
          <div className="bg-[#0369A1] text-white px-4 py-2 font-bold text-sm">
            📈 STEP 2 판매 분석 데이터 <b>{nSales.toLocaleString()}</b>개 품목 — 수요 기반 분석 자동 포함
          </div>
          <div className="overflow-x-auto max-h-72">
            <table className="text-xs w-full">
              <thead className="bg-gray-50 sticky top-0">
                <tr>
                  {['품목코드','품목명','23년 판매','24년 판매','25년 판매','25년 생산비중'].map(h => (
                    <th key={h} className="px-3 py-2 text-left border-b font-semibold text-gray-600">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Object.entries(sales).slice(0, 100).map(([code, item]) => (
                  <tr key={code} className="hover:bg-gray-50 border-b">
                    <td className="px-3 py-1 font-mono text-gray-500">{code}</td>
                    <td className="px-3 py-1 max-w-xs truncate">{item.품목명}</td>
                    <td className="px-3 py-1 text-right">{item['23'].sales.toLocaleString()}</td>
                    <td className="px-3 py-1 text-right">{item['24'].sales.toLocaleString()}</td>
                    <td className="px-3 py-1 text-right">{item['25'].sales.toLocaleString()}</td>
                    <td className="px-3 py-1 text-right">{(item['25'].ratio * 100).toFixed(1)}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}
