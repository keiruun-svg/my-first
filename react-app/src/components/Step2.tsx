import { useState } from 'react'
import * as XLSX from 'xlsx'
import { saveSalesAnalysis } from '../lib/supabase'
import { classifyOjc } from '../lib/ojcFilter'
import type { SalesAnalysis, SalesItem } from '../lib/types'
import FileUploader from './FileUploader'

interface Props {
  sales: SalesAnalysis
  setSales: (s: SalesAnalysis) => void
}

export default function Step2({ sales, setSales }: Props) {
  const [logs, setLogs] = useState<string[]>([])
  const [running, setRunning] = useState(false)
  const [done, setDone] = useState(false)
  const [salesFile, setSalesFile] = useState<File | null>(null)
  const [purchaseFile, setPurchaseFile] = useState<File | null>(null)

  const salesYears = Array.from(new Set(
    Object.values(sales).flatMap(v => Object.keys(v).filter(k => /^\d{2}$/.test(k)))
  )).sort()

  const nSales = Object.values(sales).filter(v =>
    salesYears.some(yr => ((v[yr] as { sales: number } | undefined)?.sales ?? 0) > 0)
  ).length

  const run = async () => {
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
        if (!/^\d{2}$/.test(yr)) continue
        if (!salesBy[code]) salesBy[code] = { 품목명: name }
        salesBy[code][yr] = (Number(salesBy[code][yr] as number ?? 0)) + qty
      }
      newLogs.push(`OJC 분류 완료 — ${Object.keys(salesBy).length.toLocaleString()}개 품목`)

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
          if (!/^\d{2}$/.test(yr)) continue
          if (!prodBy[code]) prodBy[code] = {}
          prodBy[code][yr] = (prodBy[code][yr] ?? 0) + qty
        }
        newLogs.push(`구매관리(맥산) 파일 로드 완료`)
      }

      const detectedYears = Array.from(new Set(
        Object.values(salesBy).flatMap(sd => Object.keys(sd).filter(k => /^\d{2}$/.test(k)))
      )).sort()

      const analysis: SalesAnalysis = {}
      for (const [code, sd] of Object.entries(salesBy)) {
        const item: SalesItem = { 품목명: sd['품목명'] as string }
        for (const yr of detectedYears) {
          const s = Number(sd[yr] as number ?? 0)
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
      {/* step-box */}
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>STEP 2 — 판매 분석</b>: 전체 판매량 파일과 구매관리(맥산) 파일을 분석하여
        품목별 판매/생산 비중을 계산합니다. STEP 3 수요 기반 분석에 자동 반영됩니다.
      </div>

      {/* columns [1, 1] */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <FileUploader
          label="① 전체 판매량.xlsx (ERP 원본)"
          fileName={salesFile?.name ?? ''}
          onFile={setSalesFile}
        />
        <FileUploader
          label="② 구매관리(맥산).xlsx (ERP 원본)"
          fileName={purchaseFile?.name ?? ''}
          onFile={setPurchaseFile}
          optional
        />
      </div>

      <hr className="border-gray-200" />

      {/* buttons */}
      <div className="flex gap-2">
        <button
          onClick={run}
          disabled={running || !salesFile}
          className="flex-1 bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold py-2 px-4 rounded transition text-sm"
        >
          {running ? '⏳ 분석 중...' : '▶ STEP 2 실행 — 판매 분석'}
        </button>
        {(done || logs.length > 0) && (
          <button
            onClick={() => { setDone(false); setLogs([]); setSalesFile(null); setPurchaseFile(null) }}
            className="px-4 py-2 text-sm border border-gray-300 rounded text-gray-600 hover:bg-gray-100 transition"
          >
            🗑 초기화
          </button>
        )}
      </div>

      {done && nSales > 0 && (
        <div className="bg-[#D6F0D8] px-4 py-3 rounded-md text-sm font-semibold text-[#1a6a2a]">
          ✅ 판매 분석 완료 — <b>{nSales.toLocaleString()}</b>개 품목 저장됨. STEP 3에서 수요 기반 분석이 자동 포함됩니다.
        </div>
      )}

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs space-y-0.5 max-h-48 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {nSales > 0 && !done && (
        <div>
          <div className="text-sm font-semibold text-gray-700 mb-2">
            저장된 분석 데이터 — <b>{nSales.toLocaleString()}</b>개 품목
          </div>
          <div className="overflow-x-auto border border-gray-200 rounded">
            <table className="text-xs w-full">
              <thead className="bg-gray-50 sticky top-0">
                <tr>
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">품목코드</th>
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">품목명</th>
                  {salesYears.map(yr => (
                    <th key={yr} className="px-3 py-2 text-left border-b font-semibold text-gray-600">{yr}년 판매</th>
                  ))}
                  {salesYears.length > 0 && (
                    <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">{salesYears[salesYears.length - 1]}년 생산비중</th>
                  )}
                </tr>
              </thead>
              <tbody>
                {Object.entries(sales).slice(0, 100).map(([code, item]) => (
                  <tr key={code} className="hover:bg-gray-50 border-b">
                    <td className="px-3 py-1.5 font-mono text-gray-500">{code}</td>
                    <td className="px-3 py-1.5 max-w-xs truncate">{item.품목명}</td>
                    {salesYears.map(yr => (
                      <td key={yr} className="px-3 py-1.5 text-right">
                        {((item[yr] as { sales: number } | undefined)?.sales ?? 0).toLocaleString()}
                      </td>
                    ))}
                    {salesYears.length > 0 && (
                      <td className="px-3 py-1.5 text-right">
                        {(((item[salesYears[salesYears.length - 1]] as { ratio: number } | undefined)?.ratio ?? 0) * 100).toFixed(1)}%
                      </td>
                    )}
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
