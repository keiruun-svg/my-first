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
  const salesRef = useRef<HTMLInputElement>(null)
  const purchaseRef = useRef<HTMLInputElement>(null)

  const run = async () => {
    const salesFile = salesRef.current?.files?.[0]
    if (!salesFile) return alert('판매량 파일을 선택해주세요.')
    setRunning(true); setLogs([])
    try {
      const newLogs: string[] = []
      const salesBuf = await salesFile.arrayBuffer()
      const wb = XLSX.read(salesBuf, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })
      newLogs.push(`판매량 파일 로드 — ${rows.length.toLocaleString()}행`)

      // Collect OJC rows
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

      // Build production data from purchase file if available
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
        newLogs.push('구매관리 파일 로드 완료')
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
    } catch (e) {
      setLogs([`❌ 오류: ${e}`])
    } finally {
      setRunning(false)
    }
  }

  const hasSales = Object.keys(sales).length > 0

  return (
    <div className="space-y-4">
      <div className="bg-green-50 border border-green-200 rounded p-4">
        <h3 className="font-bold text-green-800 mb-2">STEP 2 — 판매 분석</h3>
        <p className="text-sm text-green-700">전체 판매량 파일과 구매관리(맥산) 파일을 분석하여 품목별 판매/생산 비중을 계산합니다.</p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-semibold text-gray-700 mb-1">판매량 파일 (필수)</label>
          <input ref={salesRef} type="file" accept=".xlsx,.xls"
            className="block w-full text-sm border rounded p-2 bg-white" />
        </div>
        <div>
          <label className="block text-sm font-semibold text-gray-700 mb-1">구매관리(맥산) 파일 (선택)</label>
          <input ref={purchaseRef} type="file" accept=".xlsx,.xls"
            className="block w-full text-sm border rounded p-2 bg-white" />
        </div>
      </div>

      <button
        onClick={run}
        disabled={running}
        className="w-full bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white font-bold py-3 rounded-lg transition"
      >
        {running ? '⏳ 분석 중...' : '▶ STEP 2 실행'}
      </button>

      {logs.length > 0 && (
        <div className="bg-gray-900 text-green-300 rounded p-4 font-mono text-sm space-y-1 max-h-60 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {hasSales && (
        <div className="bg-white border rounded-lg overflow-hidden">
          <div className="bg-green-700 text-white px-4 py-2 font-bold text-sm">
            저장된 판매 분석 — {Object.keys(sales).length.toLocaleString()}개 품목
          </div>
          <div className="overflow-x-auto max-h-80">
            <table className="text-xs w-full">
              <thead className="bg-gray-50 sticky top-0">
                <tr>
                  {['품목코드','품목명','23년 판매','24년 판매','25년 판매','25년 생산비중'].map(h => (
                    <th key={h} className="px-3 py-2 text-left border-b">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Object.entries(sales).slice(0, 100).map(([code, item]) => (
                  <tr key={code} className="hover:bg-gray-50 border-b">
                    <td className="px-3 py-1 font-mono">{code}</td>
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
