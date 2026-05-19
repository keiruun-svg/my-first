import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import { saveMetadata, saveInventory, CAN_WRITE } from '../lib/supabase'
import type { Metadata, Inventory, CableMeta, HousingComp } from '../lib/types'

interface Props {
  metadata:     Metadata
  setMetadata:  (m: Metadata) => void
  inventory:    Inventory
  setInventory: (i: Inventory) => void
}

interface PreviewRow {
  key: string; pai: string; type: string
  품번: string; 품명: string; 구매처: string; 리드타임: string
  isNew: boolean
}

interface ImportRow {
  code: string; name: string; total: number; byWh: Record<string, number>
  matchType: 'cable' | 'housing' | null; metaKey: string | null
}

const thCls = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-200 bg-gray-50 whitespace-nowrap text-left'
const thR   = `${thCls} text-right`

export default function MaterialManager({ metadata, setMetadata, inventory, setInventory }: Props) {
  const [tab,   setTab]   = useState<'cable' | 'housing'>('cable')
  const [saved, setSaved] = useState<'' | 'both' | 'meta' | 'inv'>('')

  // 품번 Excel 업로드 상태
  const [pnPanel,    setPnPanel]    = useState(false)
  const [preview,    setPreview]    = useState<{ kind: 'cable' | 'housing'; rows: PreviewRow[] } | null>(null)
  const [previewErr, setPreviewErr] = useState('')
  const pnFileRef = useRef<HTMLInputElement>(null)

  // ERP 재고 업로드 상태
  const [erpPanel,       setErpPanel]       = useState(false)
  const [importRows,     setImportRows]     = useState<ImportRow[] | null>(null)
  const [warehouses,     setWarehouses]     = useState<string[]>([])
  const [selectedWh,     setSelectedWh]     = useState('')
  const [importSubTab,   setImportSubTab]   = useState<'total' | 'warehouse'>('total')
  const [importApplied,  setImportApplied]  = useState(false)
  const [importFileName, setImportFileName] = useState('')
  const erpFileRef = useRef<HTMLInputElement>(null)

  // ── 메타 수정 ──────────────────────────────────────────────────
  const updateCableMeta = (key: string, field: keyof CableMeta, value: string) => {
    setMetadata({ ...metadata, cable: {
      ...metadata.cable,
      [key]: { ...(metadata.cable[key] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }), [field]: value },
    }})
  }
  const updateHousingMeta = (key: string, field: keyof HousingComp, value: string) => {
    const cur  = metadata.housing[key]
    const base = (Array.isArray(cur) ? cur[0] : cur) ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }
    setMetadata({ ...metadata, housing: { ...metadata.housing, [key]: { ...base, [field]: value } } })
  }

  // ── 재고 수정 ──────────────────────────────────────────────────
  const updateCableInv = (key: string, field: '현재고' | '기발주', value: string) => {
    const prev = inventory.cable[key] ?? { 현재고: 0, 기발주: 0 }
    setInventory({ ...inventory, cable: { ...inventory.cable, [key]: { ...prev, [field]: parseInt(value) || 0 } } })
  }
  const updateHousingInv = (key: string, field: '현재고' | '기발주', value: string) => {
    const cur  = inventory.housing[key]
    const prev = (Array.isArray(cur) ? cur[0] : cur) ?? { 현재고: 0, 기발주: 0 }
    setInventory({ ...inventory, housing: { ...inventory.housing, [key]: { ...prev, [field]: parseInt(value) || 0 } } })
  }

  // ── 저장 ───────────────────────────────────────────────────────
  const flash = (type: typeof saved) => { setSaved(type); setTimeout(() => setSaved(''), 2000) }
  const saveAll  = () => { saveMetadata(metadata); saveInventory(inventory); flash('both') }

  // ── 품번 Excel 템플릿 다운로드 ────────────────────────────────
  function downloadTemplate() {
    const isCable = tab === 'cable'
    const keys    = isCable ? Object.keys(metadata.cable).sort() : Object.keys(metadata.housing).sort()
    const label   = isCable ? '케이블종류' : '하우징타입'
    const getCm   = (k: string): CableMeta => metadata.cable[k] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }
    const getHm   = (k: string): HousingComp => { const c = metadata.housing[k]; return (Array.isArray(c) ? c[0] : c) ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null } }
    const header  = ['파이', label, '품번', '품명', '구매처', '리드타임(일)']
    const rows    = keys.length ? keys.map(k => {
      const [pai, type] = k.split('|')
      const m = isCable ? getCm(k) : getHm(k)
      return [pai, type, m.품번 ?? '', m.품명 ?? '', m.구매처 ?? '', m.리드타임 ?? '']
    }) : [['A1', '예시종류', '', '', '', '60']]
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows])
    ws['!cols'] = [10, 20, 16, 30, 16, 12].map(w => ({ wch: w }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, isCable ? '케이블' : '하우징')
    XLSX.writeFile(wb, `품번관리_${isCable ? '케이블' : '하우징'}_템플릿.xlsx`)
  }

  // ── 품번 Excel 업로드 파싱 ────────────────────────────────────
  function handlePnUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]; if (!file) return
    e.target.value = ''; setPreviewErr('')
    const reader = new FileReader()
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target?.result, { type: 'array' })
        const sheetName = tab === 'cable' ? '케이블' : '하우징'
        const ws = wb.Sheets[sheetName] ?? wb.Sheets[wb.SheetNames[0]]
        if (!ws) { setPreviewErr('시트를 찾을 수 없습니다.'); return }
        const raw = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' }) as string[][]
        if (raw.length < 2) { setPreviewErr('데이터 행이 없습니다.'); return }
        const parsed: PreviewRow[] = []
        for (let i = 1; i < raw.length; i++) {
          const r = raw[i]
          const pai  = String(r[0] ?? '').trim()
          const type = String(r[1] ?? '').trim()
          if (!pai || !type) continue
          const key    = `${pai}|${type}`
          const exists = tab === 'cable' ? !!metadata.cable[key] : !!metadata.housing[key]
          parsed.push({ key, pai, type,
            품번: String(r[2] ?? '').trim(), 품명: String(r[3] ?? '').trim(),
            구매처: String(r[4] ?? '').trim(), 리드타임: String(r[5] ?? '').trim(),
            isNew: !exists })
        }
        if (!parsed.length) { setPreviewErr('유효한 데이터가 없습니다.'); return }
        setPreview({ kind: tab, rows: parsed })
      } catch { setPreviewErr('파일을 읽는 중 오류가 발생했습니다.') }
    }
    reader.readAsArrayBuffer(file)
  }

  function applyPreview() {
    if (!preview) return
    const next = { ...metadata }
    if (preview.kind === 'cable') {
      const cable = { ...next.cable }
      for (const r of preview.rows)
        cable[r.key] = { 품번: r.품번, 품명: r.품명, 구매처: r.구매처, 리드타임: r.리드타임 || null }
      next.cable = cable
    } else {
      const housing = { ...next.housing }
      for (const r of preview.rows) {
        const cur  = housing[r.key]
        const base = (Array.isArray(cur) ? cur[0] : cur) ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }
        housing[r.key] = { ...base, 품번: r.품번, 품명: r.품명, 구매처: r.구매처, 리드타임: r.리드타임 || null }
      }
      next.housing = housing
    }
    setMetadata(next); saveMetadata(next)
    setPreview(null); flash('meta')
  }

  // ── ERP 재고 파일 파싱 ────────────────────────────────────────
  function handleErpUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]; if (!file) return
    setImportFileName(file.name); e.target.value = ''
    const reader = new FileReader()
    reader.onload = ev => {
      const wb = XLSX.read(ev.target?.result, { type: 'array' })
      const ws = wb.Sheets['재고현황']
      if (!ws) { alert('재고현황 시트를 찾을 수 없습니다.'); return }
      const range = XLSX.utils.decode_range(ws['!ref']!)
      const headerRow: string[] = []
      for (let c = 0; c <= range.e.c; c++)
        headerRow.push(ws[XLSX.utils.encode_cell({ r: 1, c })]?.v?.toString().trim() ?? '')
      const whCols = headerRow.slice(4).filter(h => h)
      setWarehouses(whCols); setSelectedWh(whCols[0] ?? '')
      const cableMap: Record<string, string> = {}
      const housingMap: Record<string, string> = {}
      Object.entries(metadata.cable).forEach(([k, m]) => { if (m.품번) cableMap[m.품번] = k })
      Object.entries(metadata.housing).forEach(([k, mRaw]) => {
        const m = Array.isArray(mRaw) ? mRaw[0] : mRaw
        if (m?.품번) housingMap[m.품번] = k
      })
      const rows: ImportRow[] = []
      for (let r = 2; r <= range.e.r; r++) {
        const code  = ws[XLSX.utils.encode_cell({ r, c: 0 })]?.v?.toString().trim() ?? ''
        const name  = ws[XLSX.utils.encode_cell({ r, c: 1 })]?.v?.toString().trim() ?? ''
        const total = Number(ws[XLSX.utils.encode_cell({ r, c: 3 })]?.v ?? 0)
        if (!code) continue
        const byWh: Record<string, number> = {}
        whCols.forEach((wh, wi) => { byWh[wh] = Number(ws[XLSX.utils.encode_cell({ r, c: 4 + wi })]?.v ?? 0) })
        let matchType: 'cable' | 'housing' | null = null
        let metaKey: string | null = null
        if (cableMap[code])        { matchType = 'cable';   metaKey = cableMap[code] }
        else if (housingMap[code]) { matchType = 'housing'; metaKey = housingMap[code] }
        rows.push({ code, name, total, byWh, matchType, metaKey })
      }
      setImportRows(rows); setImportApplied(false)
    }
    reader.readAsArrayBuffer(file)
  }

  function applyErpImport(useTotal: boolean) {
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

  // ── 공통 헬퍼 ─────────────────────────────────────────────────
  const getCm  = (k: string): CableMeta   => metadata.cable[k] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }
  const getHm  = (k: string): HousingComp => { const c = metadata.housing[k]; return (Array.isArray(c) ? c[0] : c) ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null } }
  const getCiv = (k: string) => inventory.cable[k] ?? { 현재고: 0, 기발주: 0 }
  const getHiv = (k: string) => { const c = inventory.housing[k]; return (Array.isArray(c) ? c[0] : c) ?? { 현재고: 0, 기발주: 0 } }

  const cableKeys   = Object.keys(metadata.cable).sort()
  const housingKeys = Object.keys(metadata.housing).sort()
  const cableMissing   = cableKeys.filter(k => !getCm(k).품번).length
  const housingMissing = housingKeys.filter(k => !getHm(k).품번).length
  const matched   = importRows?.filter(r => r.matchType) ?? []
  const unmatched = importRows?.filter(r => !r.matchType) ?? []

  const inputCls    = 'w-full border border-gray-200 rounded px-2 py-0.5 text-xs focus:outline-none focus:border-blue-400'
  const inputRoCls  = `${inputCls} bg-gray-50 text-gray-400 cursor-default`
  const inputInvCls = 'w-24 border border-gray-200 rounded px-2 py-0.5 text-xs text-right focus:outline-none focus:border-blue-400'

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>자재 관리</b> — 품번·구매처·리드타임과 현재고·기발주를 한 화면에서 관리합니다.
        {!CAN_WRITE && <span className="ml-2 text-gray-500">(품번 정보는 읽기 전용)</span>}
      </div>

      {/* ── 툴바 ─────────────────────────────────────────────── */}
      <div className="flex items-center justify-between flex-wrap gap-2">
        {/* 케이블 / 하우징 토글 */}
        <div className="flex gap-2">
          {([
            ['cable',   '케이블', cableKeys.length,   cableMissing],
            ['housing', '하우징', housingKeys.length,  housingMissing],
          ] as [typeof tab, string, number, number][]).map(([id, lbl, cnt, miss]) => (
            <button key={id} onClick={() => { setTab(id); setPreview(null) }}
              className={`px-4 py-2 rounded font-semibold text-sm flex items-center gap-2 transition ${
                tab === id ? 'bg-[#2E75B6] text-white' : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
              }`}>
              {lbl}
              <span className={`text-xs px-1.5 py-0.5 rounded-full font-bold ${tab === id ? 'bg-white/20' : 'bg-gray-300'}`}>{cnt}</span>
              {miss > 0 && <span className="text-xs bg-red-500 text-white px-1.5 py-0.5 rounded-full">⚠ {miss}</span>}
            </button>
          ))}
        </div>

        {/* 액션 버튼들 */}
        <div className="flex items-center gap-2">
          {CAN_WRITE && (
            <>
              <button onClick={downloadTemplate}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border border-gray-300 rounded bg-white hover:bg-gray-50 transition">
                📤 템플릿 다운로드
              </button>
              <button onClick={() => { setPnPanel(v => !v); setErpPanel(false) }}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border rounded transition ${
                  pnPanel ? 'border-[#2E75B6] bg-blue-50 text-[#2E75B6]' : 'border-[#2E75B6] text-[#2E75B6] bg-white hover:bg-blue-50'
                }`}>
                📥 품번 일괄 업로드
              </button>
            </>
          )}
          <button onClick={() => { setErpPanel(v => !v); setPnPanel(false) }}
            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border rounded transition ${
              erpPanel ? 'border-yellow-500 bg-yellow-50 text-yellow-700' : 'border-yellow-500 text-yellow-700 bg-white hover:bg-yellow-50'
            }`}>
            🏭 ERP 재고 업로드
          </button>
        </div>
      </div>

      {/* ── 품번 Excel 업로드 패널 ──────────────────────────── */}
      {pnPanel && CAN_WRITE && (
        <div className="border border-blue-200 rounded-lg bg-blue-50 p-4 space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-sm font-semibold text-blue-800">
              품번 일괄 업로드 — {tab === 'cable' ? '케이블' : '하우징'}
            </span>
            <button onClick={() => { setPnPanel(false); setPreview(null); setPreviewErr('') }}
              className="text-xs text-gray-400 hover:text-gray-600">✕</button>
          </div>
          {previewErr && <div className="text-sm text-red-600 bg-red-50 border border-red-200 px-3 py-2 rounded">{previewErr}</div>}
          {!preview ? (
            <button onClick={() => pnFileRef.current?.click()}
              className="px-4 py-2 bg-[#2E75B6] text-white text-sm font-semibold rounded hover:bg-[#1f5a9a] transition">
              📂 Excel 파일 선택
            </button>
          ) : (
            <div className="space-y-2">
              <div className="text-xs text-blue-700">
                {preview.rows.length}개 행 —
                기존 <b>{preview.rows.filter(r => !r.isNew).length}</b>개 업데이트 /
                신규 <b>{preview.rows.filter(r => r.isNew).length}</b>개 추가
              </div>
              <div className="overflow-x-auto max-h-52 overflow-y-auto rounded border border-blue-100">
                <table className="text-xs w-full bg-white">
                  <thead className="sticky top-0 bg-blue-100 text-blue-800">
                    <tr>{['파이', tab==='cable'?'케이블종류':'하우징타입','품번','품명','구매처','LT','구분'].map(h=>(
                      <th key={h} className="px-3 py-1.5 text-left font-semibold">{h}</th>
                    ))}</tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {preview.rows.map(r => (
                      <tr key={r.key} className={r.isNew ? 'bg-green-50' : ''}>
                        <td className="px-3 py-1 font-mono text-gray-500">{r.pai}</td>
                        <td className="px-3 py-1 font-semibold">{r.type}</td>
                        <td className="px-3 py-1">{r.품번 || <span className="text-red-400">미입력</span>}</td>
                        <td className="px-3 py-1">{r.품명 || '—'}</td>
                        <td className="px-3 py-1">{r.구매처 || '—'}</td>
                        <td className="px-3 py-1">{r.리드타임 || '—'}</td>
                        <td className="px-3 py-1">{r.isNew ? <span className="text-green-600 font-semibold">신규</span> : <span className="text-blue-600">업데이트</span>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex gap-2">
                <button onClick={applyPreview}
                  className="bg-[#2E75B6] hover:bg-[#1F5A9E] text-white font-bold px-5 py-2 rounded-lg text-sm transition">
                  ✅ 적용 및 저장
                </button>
                <button onClick={() => setPreview(null)} className="px-4 py-2 text-sm text-gray-500 hover:text-gray-700">취소</button>
              </div>
            </div>
          )}
          <input ref={pnFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handlePnUpload} />
        </div>
      )}

      {/* ── ERP 재고 업로드 패널 ────────────────────────────── */}
      {erpPanel && (
        <div className="border border-yellow-200 rounded-lg bg-yellow-50 p-4 space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-sm font-semibold text-yellow-800">ERP 재고 업로드 (ESZ018R)</span>
            <button onClick={() => { setErpPanel(false); setImportRows(null); setImportFileName('') }}
              className="text-xs text-gray-400 hover:text-gray-600">✕</button>
          </div>
          <div className="flex items-center gap-4">
            <button onClick={() => erpFileRef.current?.click()}
              className="px-4 py-2 bg-[#2E75B6] text-white text-sm font-semibold rounded hover:bg-[#1f5a9a] transition">
              📂 파일 선택
            </button>
            <span className="text-sm text-gray-500">{importFileName || 'ESZ018R xlsx (재고현황 시트)'}</span>
            {importRows && (
              <span className="ml-auto text-xs text-gray-500">
                매칭 <span className="text-green-600 font-bold">{matched.length}</span>건 / 미매칭 {unmatched.length}건
              </span>
            )}
          </div>
          {importRows && (
            <>
              <div className="flex border-b border-yellow-200">
                {([['total','총수량(D열)'],['warehouse','창고별']] as ['total'|'warehouse',string][]).map(([id,lbl])=>(
                  <button key={id} onClick={() => { setImportSubTab(id); setImportApplied(false) }}
                    className={`px-4 py-2 text-sm font-medium border-b-2 -mb-px transition ${
                      importSubTab===id ? 'border-[#2E75B6] text-[#2E75B6]' : 'border-transparent text-gray-500 hover:text-gray-700'
                    }`}>{lbl}</button>
                ))}
              </div>
              {importSubTab === 'warehouse' && (
                <div className="flex items-center gap-3">
                  <span className="text-sm font-medium text-yellow-800">창고 선택</span>
                  <select value={selectedWh} onChange={e => { setSelectedWh(e.target.value); setImportApplied(false) }}
                    className="border border-yellow-300 rounded px-2 py-1 text-sm bg-white focus:outline-none">
                    {warehouses.map(wh => <option key={wh} value={wh}>{wh}</option>)}
                  </select>
                </div>
              )}
              <div className="overflow-x-auto max-h-52 overflow-y-auto rounded border border-yellow-100">
                <table className="text-xs w-full bg-white">
                  <thead className="sticky top-0 bg-yellow-100">
                    <tr><th className={thCls}>품목코드</th><th className={thCls}>품목명</th><th className={thCls}>구분</th>
                        <th className={thR}>{importSubTab==='total'?'총수량':selectedWh}</th><th className={thCls}>매칭</th></tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {matched.map((r, i) => {
                      const qty = importSubTab==='total' ? r.total : (r.byWh[selectedWh]??0)
                      return (
                        <tr key={i} className={i%2===0?'bg-white':'bg-gray-50'}>
                          <td className="px-3 py-1 font-mono text-gray-500">{r.code}</td>
                          <td className="px-3 py-1 max-w-xs truncate" title={r.name}>{r.name}</td>
                          <td className="px-3 py-1">
                            <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${r.matchType==='cable'?'bg-blue-50 text-blue-700':'bg-purple-50 text-purple-700'}`}>
                              {r.matchType==='cable'?'케이블':'하우징'}
                            </span>
                          </td>
                          <td className="px-3 py-1 text-right font-mono font-semibold">{qty.toLocaleString()}</td>
                          <td className="px-3 py-1 text-xs text-green-600">✓ {r.metaKey}</td>
                        </tr>
                      )
                    })}
                    {unmatched.map((r, i) => {
                      const qty = importSubTab==='total' ? r.total : (r.byWh[selectedWh]??0)
                      return (
                        <tr key={`u-${i}`} className="opacity-40">
                          <td className="px-3 py-1 font-mono text-gray-500">{r.code}</td>
                          <td className="px-3 py-1 max-w-xs truncate" title={r.name}>{r.name}</td>
                          <td className="px-3 py-1 text-xs text-gray-400">—</td>
                          <td className="px-3 py-1 text-right font-mono text-gray-400">{qty.toLocaleString()}</td>
                          <td className="px-3 py-1 text-xs text-gray-400">미매칭</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
              <div className="flex items-center gap-3">
                <button onClick={() => applyErpImport(importSubTab==='total')}
                  className="px-5 py-2 bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold text-sm rounded-lg transition">
                  {importSubTab==='total' ? `총수량으로 현재고 적용 (${matched.length}건)` : `${selectedWh} 수량으로 현재고 적용 (${matched.length}건)`}
                </button>
                {importApplied && <span className="text-green-600 text-sm font-semibold">✅ 적용됐습니다. 저장 버튼으로 확정하세요.</span>}
              </div>
            </>
          )}
          <input ref={erpFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleErpUpload} />
        </div>
      )}

      {/* ── 통합 테이블 ─────────────────────────────────────── */}
      <div className="overflow-x-auto border border-gray-200 rounded-lg">
        <table className="text-xs w-full border-collapse bg-white">
          <thead>
            <tr>
              <th className={thCls}>파이</th>
              <th className={thCls}>{tab==='cable'?'케이블종류':'하우징타입'}</th>
              <th className={thCls}>품번</th>
              <th className={thCls}>품명</th>
              <th className={thCls}>구매처</th>
              <th className={thCls}>LT(일)</th>
              <th className={thR}>현재고{tab==='cable'?' (m)':' (EA)'}</th>
              <th className={thR}>기발주{tab==='cable'?' (m)':' (EA)'}</th>
              <th className={thCls}>상태</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {(tab==='cable' ? cableKeys : housingKeys).map((key, i) => {
              const [pai, type] = key.split('|')
              const m   = tab==='cable' ? getCm(key) : getHm(key)
              const inv = tab==='cable' ? getCiv(key) : getHiv(key)
              const missing = !m.품번
              return (
                <tr key={key} className={missing ? 'bg-red-50' : i%2===0 ? 'bg-white' : 'bg-gray-50'}>
                  <td className="px-3 py-1 font-mono text-gray-500 whitespace-nowrap">{pai}</td>
                  <td className="px-3 py-1.5 font-semibold whitespace-nowrap">{type}</td>
                  {/* 품번 필드 (품번·품명·구매처·LT) */}
                  {(['품번','품명','구매처','리드타임'] as const).map(f => (
                    <td key={f} className="px-1 py-1">
                      <input
                        value={String(m[f] ?? '')}
                        readOnly={!CAN_WRITE}
                        onChange={e => tab==='cable' ? updateCableMeta(key, f, e.target.value) : updateHousingMeta(key, f, e.target.value)}
                        className={CAN_WRITE
                          ? `${inputCls} ${missing && f==='품번' ? 'border-red-400 bg-red-50' : ''}`
                          : inputRoCls}
                        placeholder={f==='리드타임' ? '숫자(일)' : f}
                      />
                    </td>
                  ))}
                  {/* 재고 필드 */}
                  <td className="px-1 py-1">
                    <input type="number" min="0" value={inv.현재고}
                      onChange={e => tab==='cable' ? updateCableInv(key,'현재고',e.target.value) : updateHousingInv(key,'현재고',e.target.value)}
                      className={inputInvCls} />
                  </td>
                  <td className="px-1 py-1">
                    <input type="number" min="0" value={inv.기발주 ?? 0}
                      onChange={e => tab==='cable' ? updateCableInv(key,'기발주',e.target.value) : updateHousingInv(key,'기발주',e.target.value)}
                      className={inputInvCls} />
                  </td>
                  <td className="px-2 py-1 text-center">
                    {missing ? <span className="text-red-500 font-semibold">⚠</span> : <span className="text-green-600">✓</span>}
                  </td>
                </tr>
              )
            })}
            {(tab==='cable' ? cableKeys : housingKeys).length === 0 && (
              <tr>
                <td colSpan={9} className="text-center text-gray-400 py-10">
                  STEP 1 실행 후 품번을 입력하세요.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* ── 저장 ────────────────────────────────────────────── */}
      <div className="flex items-center gap-3">
        <button onClick={saveAll}
          className="bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold px-6 py-2 rounded-lg transition">
          💾 저장
        </button>
        {saved && (
          <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>
        )}
      </div>
    </div>
  )
}
