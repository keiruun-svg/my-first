import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import { saveInventory } from '../lib/supabase'
import type { Inventory as Inv, Metadata } from '../lib/types'

interface Props {
  inventory: Inv
  setInventory: (i: Inv) => void
  metadata: Metadata
}

interface ImportRow {
  code: string
  name: string
  total: number
  byWh: Record<string, number>
  matchType: 'cable' | 'housing' | null
  metaKey: string | null
}

const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'
const thL    = `${thBase} text-left`
const thR    = `${thBase} text-right`
const tdL    = 'px-3 py-1.5 text-sm text-gray-900'
const tdMono = 'px-3 py-1.5 text-xs font-mono text-gray-500'

export default function Inventory({ inventory, setInventory, metadata }: Props) {
  const [tab, setTab]       = useState<'cable' | 'housing' | 'import'>('cable')
  const [saved, setSaved]   = useState(false)

  // ── ERP 업로드 상태 ─────────────────────────────────────────
  const [importRows,    setImportRows]    = useState<ImportRow[] | null>(null)
  const [warehouses,    setWarehouses]    = useState<string[]>([])
  const [selectedWh,    setSelectedWh]    = useState('')
  const [importSubTab,  setImportSubTab]  = useState<'total' | 'warehouse'>('total')
  const [importApplied, setImportApplied] = useState(false)
  const [importFileName, setImportFileName] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)

  const updateCable = (key: string, field: '현재고' | '기발주', value: string) => {
    const prev = inventory.cable[key] ?? { 현재고: 0, 기발주: 0 }
    setInventory({ ...inventory, cable: { ...inventory.cable, [key]: { ...prev, [field]: parseInt(value) || 0 } } })
  }

  const updateHousing = (key: string, field: '현재고' | '기발주', value: string) => {
    const cur  = inventory.housing[key]
    const prev = Array.isArray(cur) ? cur[0] : (cur ?? { 현재고: 0, 기발주: 0 })
    setInventory({ ...inventory, housing: { ...inventory.housing, [key]: { ...prev, [field]: parseInt(value) || 0 } } })
  }

  const save = () => {
    saveInventory(inventory)
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  // ── ESZ018R 파싱 ────────────────────────────────────────────
  function handleImportFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    setImportFileName(file.name)
    const reader = new FileReader()
    reader.onload = ev => {
      const wb = XLSX.read(ev.target?.result, { type: 'array' })
      const ws = wb.Sheets['재고현황']
      if (!ws) { alert('재고현황 시트를 찾을 수 없습니다.'); return }

      const range = XLSX.utils.decode_range(ws['!ref']!)

      // 2행(index 1) = 헤더
      const headerRow: string[] = []
      for (let c = 0; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: 1, c })]
        headerRow.push(cell?.v?.toString().trim() ?? '')
      }
      // E열(index 4)~끝 = 창고명
      const whCols = headerRow.slice(4).filter(h => h)
      setWarehouses(whCols)
      setSelectedWh(whCols[0] ?? '')

      // 품번 역방향 맵 구성
      const cableMap: Record<string, string> = {}
      const housingMap: Record<string, string> = {}
      Object.entries(metadata.cable).forEach(([key, m]) => {
        if (m.품번) cableMap[m.품번] = key
      })
      Object.entries(metadata.housing).forEach(([key, mRaw]) => {
        const m = Array.isArray(mRaw) ? mRaw[0] : mRaw
        if (m?.품번) housingMap[m.품번] = key
      })

      // 3행(index 2)~끝 = 데이터
      const rows: ImportRow[] = []
      for (let r = 2; r <= range.e.r; r++) {
        const code  = ws[XLSX.utils.encode_cell({ r, c: 0 })]?.v?.toString().trim() ?? ''
        const name  = ws[XLSX.utils.encode_cell({ r, c: 1 })]?.v?.toString().trim() ?? ''
        const total = Number(ws[XLSX.utils.encode_cell({ r, c: 3 })]?.v ?? 0)
        if (!code) continue

        const byWh: Record<string, number> = {}
        whCols.forEach((wh, wi) => {
          byWh[wh] = Number(ws[XLSX.utils.encode_cell({ r, c: 4 + wi })]?.v ?? 0)
        })

        let matchType: 'cable' | 'housing' | null = null
        let metaKey: string | null = null
        if (cableMap[code])   { matchType = 'cable';   metaKey = cableMap[code] }
        else if (housingMap[code]) { matchType = 'housing'; metaKey = housingMap[code] }

        rows.push({ code, name, total, byWh, matchType, metaKey })
      }
      setImportRows(rows)
      setImportApplied(false)
    }
    reader.readAsArrayBuffer(file)
    // input 초기화 (같은 파일 재업로드 허용)
    e.target.value = ''
  }

  function applyImport(useTotal: boolean) {
    if (!importRows) return
    const newCable   = { ...inventory.cable }
    const newHousing = { ...inventory.housing }

    importRows.forEach(row => {
      if (!row.matchType || !row.metaKey) return
      const qty = useTotal ? row.total : (row.byWh[selectedWh] ?? 0)
      if (row.matchType === 'cable') {
        const prev = newCable[row.metaKey] ?? { 현재고: 0, 기발주: 0 }
        newCable[row.metaKey] = { ...prev, 현재고: qty }
      } else {
        const prevRaw = newHousing[row.metaKey]
        const prev    = (Array.isArray(prevRaw) ? prevRaw[0] : prevRaw) ?? { 현재고: 0, 기발주: 0 }
        newHousing[row.metaKey] = { ...prev, 현재고: qty }
      }
    })

    setInventory({ cable: newCable, housing: newHousing })
    setImportApplied(true)
  }

  const cableKeys   = Object.keys(metadata.cable).sort()
  const housingKeys = Object.keys(metadata.housing).sort()

  // 업로드 탭 필터
  const matched   = importRows?.filter(r => r.matchType) ?? []
  const unmatched = importRows?.filter(r => !r.matchType) ?? []

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>재고 현황</b>: 현재고 및 기발주 수량을 입력하거나 ERP 재고파일(ESZ018R)을 업로드하세요.
      </div>

      {/* 메인 탭 */}
      <div className="flex gap-2">
        {([
          ['cable',   `케이블 (${cableKeys.length})`],
          ['housing', `하우징 (${housingKeys.length})`],
          ['import',  '📥 ERP 재고 업로드'],
        ] as [typeof tab, string][]).map(([id, lbl]) => (
          <button
            key={id}
            onClick={() => setTab(id)}
            className={`px-5 py-2 rounded font-semibold text-sm border transition ${
              tab === id
                ? 'bg-yellow-400 border-yellow-500 text-black'
                : 'bg-white border-gray-300 text-gray-600 hover:bg-gray-50'
            }`}
          >
            {lbl}
          </button>
        ))}
      </div>

      {/* ── 케이블 탭 ─────────────────────────────────────────── */}
      {tab === 'cable' && (
        <div className="overflow-x-auto border border-gray-200 rounded-lg">
          <table className="text-sm w-full border-collapse bg-white">
            <thead>
              <tr>
                <th className={thL}>파이</th>
                <th className={thL}>케이블 종류</th>
                <th className={thL}>품번</th>
                <th className={thL}>품명</th>
                <th className={thR}>현재고 (m)</th>
                <th className={thR}>기발주 (m)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {cableKeys.map((key, i) => {
                const [pai, ct] = key.split('|')
                const m   = metadata.cable[key]
                const inv = inventory.cable[key] ?? { 현재고: 0, 기발주: 0 }
                return (
                  <tr key={key} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className={tdMono}>{pai}</td>
                    <td className={`${tdL} font-semibold`}>{ct}</td>
                    <td className={tdMono}>{m?.품번 ?? ''}</td>
                    <td className={`${tdL} max-w-xs truncate`} title={m?.품명 ?? ''}>{m?.품명 ?? ''}</td>
                    <td className="px-2 py-1 text-right">
                      <input type="number" min="0" value={inv.현재고}
                        onChange={e => updateCable(key, '현재고', e.target.value)}
                        className="w-28 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400" />
                    </td>
                    <td className="px-2 py-1 text-right">
                      <input type="number" min="0" value={inv.기발주 ?? 0}
                        onChange={e => updateCable(key, '기발주', e.target.value)}
                        className="w-28 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400" />
                    </td>
                  </tr>
                )
              })}
              {cableKeys.length === 0 && (
                <tr><td colSpan={6} className="text-center text-gray-400 py-10">STEP 1 실행 후 품번 관리 탭에서 케이블을 등록하세요.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* ── 하우징 탭 ─────────────────────────────────────────── */}
      {tab === 'housing' && (
        <div className="overflow-x-auto border border-gray-200 rounded-lg">
          <table className="text-sm w-full border-collapse bg-white">
            <thead>
              <tr>
                <th className={thL}>파이</th>
                <th className={thL}>하우징 타입</th>
                <th className={thL}>품번</th>
                <th className={thL}>품명</th>
                <th className={thR}>현재고 (EA)</th>
                <th className={thR}>기발주 (EA)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {housingKeys.map((key, i) => {
                const [pai, ht] = key.split('|')
                const mRaw   = metadata.housing[key]
                const m      = Array.isArray(mRaw) ? mRaw[0] : mRaw
                const invRaw = inventory.housing[key]
                const inv    = (Array.isArray(invRaw) ? invRaw[0] : invRaw) ?? { 현재고: 0, 기발주: 0 }
                return (
                  <tr key={key} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className={tdMono}>{pai}</td>
                    <td className={`${tdL} font-semibold`}>{ht}</td>
                    <td className={tdMono}>{m?.품번 ?? ''}</td>
                    <td className={`${tdL} max-w-xs truncate`} title={m?.품명 ?? ''}>{m?.품명 ?? ''}</td>
                    <td className="px-2 py-1 text-right">
                      <input type="number" min="0" value={inv.현재고}
                        onChange={e => updateHousing(key, '현재고', e.target.value)}
                        className="w-28 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400" />
                    </td>
                    <td className="px-2 py-1 text-right">
                      <input type="number" min="0" value={inv.기발주}
                        onChange={e => updateHousing(key, '기발주', e.target.value)}
                        className="w-28 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400" />
                    </td>
                  </tr>
                )
              })}
              {housingKeys.length === 0 && (
                <tr><td colSpan={6} className="text-center text-gray-400 py-10">STEP 1 실행 후 품번 관리 탭에서 하우징을 등록하세요.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* ── ERP 업로드 탭 ─────────────────────────────────────── */}
      {tab === 'import' && (
        <div className="space-y-4">
          {/* 파일 선택 영역 */}
          <div className="bg-white border-2 border-dashed border-gray-300 rounded-lg px-6 py-5 flex items-center gap-4">
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImportFile} className="hidden" />
            <button
              onClick={() => fileRef.current?.click()}
              className="px-4 py-2 bg-[#2E75B6] text-white text-sm font-semibold rounded hover:bg-[#1f5a9a] transition"
            >
              📂 파일 선택
            </button>
            <span className="text-sm text-gray-500">
              {importFileName ? importFileName : 'ESZ018R 형식의 xlsx 파일을 선택하세요 (재고현황 시트)'}
            </span>
            {importRows && (
              <span className="ml-auto text-xs text-gray-400">
                총 {importRows.length}행 | 매칭 <span className="text-green-600 font-bold">{matched.length}</span>건 / 미매칭 {unmatched.length}건
              </span>
            )}
          </div>

          {/* 파일 로드 후 */}
          {importRows && (
            <>
              {/* 서브탭 */}
              <div className="flex border-b border-gray-200">
                {([
                  ['total',     '총수량 (D열)'],
                  ['warehouse', '창고별'],
                ] as ['total' | 'warehouse', string][]).map(([id, lbl]) => (
                  <button
                    key={id}
                    onClick={() => { setImportSubTab(id); setImportApplied(false) }}
                    className={`px-5 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${
                      importSubTab === id
                        ? 'border-[#2E75B6] text-[#2E75B6]'
                        : 'border-transparent text-gray-500 hover:text-gray-700'
                    }`}
                  >
                    {lbl}
                  </button>
                ))}
              </div>

              {/* 창고 선택 (창고별 탭일 때) */}
              {importSubTab === 'warehouse' && (
                <div className="flex items-center gap-3 bg-sky-50 border border-sky-200 rounded px-4 py-2.5">
                  <span className="text-sm font-medium text-sky-800">창고 선택</span>
                  <select
                    value={selectedWh}
                    onChange={e => { setSelectedWh(e.target.value); setImportApplied(false) }}
                    className="border border-sky-300 rounded px-2 py-1 text-sm bg-white focus:outline-none focus:border-blue-400"
                  >
                    {warehouses.map(wh => <option key={wh} value={wh}>{wh}</option>)}
                  </select>
                </div>
              )}

              {/* 매칭 결과 미리보기 */}
              <div className="overflow-x-auto border border-gray-200 rounded-lg">
                <table className="text-sm w-full border-collapse bg-white">
                  <thead>
                    <tr>
                      <th className={thL}>품목코드</th>
                      <th className={thL}>품목명</th>
                      <th className={thL}>구분</th>
                      <th className={thR}>{importSubTab === 'total' ? '총수량' : selectedWh}</th>
                      <th className={thL}>매칭</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {/* 매칭된 항목 먼저 */}
                    {matched.map((row, i) => {
                      const qty = importSubTab === 'total' ? row.total : (row.byWh[selectedWh] ?? 0)
                      return (
                        <tr key={i} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                          <td className={tdMono}>{row.code}</td>
                          <td className={`${tdL} max-w-xs truncate`} title={row.name}>{row.name}</td>
                          <td className="px-3 py-1.5">
                            <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${
                              row.matchType === 'cable'
                                ? 'bg-blue-50 text-blue-700'
                                : 'bg-purple-50 text-purple-700'
                            }`}>
                              {row.matchType === 'cable' ? '케이블' : '하우징'}
                            </span>
                          </td>
                          <td className="px-3 py-1.5 text-right font-mono text-sm font-semibold">
                            {qty.toLocaleString()}
                          </td>
                          <td className="px-3 py-1.5 text-xs text-green-600">✓ {row.metaKey}</td>
                        </tr>
                      )
                    })}
                    {/* 미매칭 항목 */}
                    {unmatched.map((row, i) => {
                      const qty = importSubTab === 'total' ? row.total : (row.byWh[selectedWh] ?? 0)
                      return (
                        <tr key={`u-${i}`} className="bg-gray-50 opacity-50">
                          <td className={tdMono}>{row.code}</td>
                          <td className={`${tdL} max-w-xs truncate`} title={row.name}>{row.name}</td>
                          <td className="px-3 py-1.5 text-xs text-gray-400">—</td>
                          <td className="px-3 py-1.5 text-right font-mono text-sm text-gray-400">
                            {qty.toLocaleString()}
                          </td>
                          <td className="px-3 py-1.5 text-xs text-gray-400">미매칭</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>

              {/* 적용 버튼 */}
              <div className="flex items-center gap-3 pt-1">
                <button
                  onClick={() => applyImport(importSubTab === 'total')}
                  className="px-5 py-2 bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold text-sm rounded-lg transition"
                >
                  {importSubTab === 'total'
                    ? `총수량으로 현재고 적용 (${matched.length}건)`
                    : `${selectedWh} 수량으로 현재고 적용 (${matched.length}건)`}
                </button>
                {importApplied && (
                  <span className="text-green-600 text-sm font-semibold">
                    ✅ 적용됐습니다. 저장 버튼으로 확정하세요.
                  </span>
                )}
              </div>
            </>
          )}
        </div>
      )}

      {/* 저장 버튼 (케이블/하우징 탭에서만 표시) */}
      {tab !== 'import' && (
        <div className="flex items-center gap-3">
          <button onClick={save}
            className="bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold px-6 py-2 rounded-lg transition">
            💾 저장
          </button>
          {saved && <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>}
        </div>
      )}

      {/* ERP 업로드 탭 저장 버튼 */}
      {tab === 'import' && importApplied && (
        <div className="flex items-center gap-3">
          <button onClick={save}
            className="bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold px-6 py-2 rounded-lg transition">
            💾 저장 (Supabase 반영)
          </button>
          {saved && <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>}
        </div>
      )}
    </div>
  )
}
