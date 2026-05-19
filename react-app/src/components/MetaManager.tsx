import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import { saveMetadata } from '../lib/supabase'
import { today } from '../lib/download'
import type { Metadata, CableMeta, HousingComp } from '../lib/types'

interface Props {
  metadata: Metadata
  setMetadata: (m: Metadata) => void
}

interface PreviewRow {
  key:      string
  pai:      string
  type:     string
  품번:     string
  품명:     string
  구매처:   string
  리드타임: string
  isNew:    boolean
}

export default function MetaManager({ metadata, setMetadata }: Props) {
  const [tab, setTab]         = useState<'cable' | 'housing'>('cable')
  const [saved, setSaved]     = useState(false)
  const [preview, setPreview] = useState<{ kind: 'cable' | 'housing'; rows: PreviewRow[] } | null>(null)
  const [previewErr, setPreviewErr] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)

  const updateCable = (key: string, field: keyof CableMeta, value: string) => {
    setMetadata({
      ...metadata,
      cable: {
        ...metadata.cable,
        [key]: { ...(metadata.cable[key] || { 품번:'', 품명:'', 구매처:'', 리드타임:null }), [field]: value }
      }
    })
  }

  const updateHousing = (key: string, field: keyof HousingComp, value: string) => {
    const cur = metadata.housing[key]
    const comp = Array.isArray(cur) ? cur[0] : (cur || { 품번:'', 품명:'', 구매처:'', 리드타임:null })
    setMetadata({
      ...metadata,
      housing: { ...metadata.housing, [key]: { ...comp, [field]: value } }
    })
  }

  const save = () => {
    saveMetadata(metadata)
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  const getCableMeta   = (key: string): CableMeta   => metadata.cable[key]   || { 품번:'', 품명:'', 구매처:'', 리드타임:null }
  const getHousingMeta = (key: string): HousingComp => {
    const cur = metadata.housing[key]
    return (Array.isArray(cur) ? cur[0] : cur) || { 품번:'', 품명:'', 구매처:'', 리드타임:null }
  }

  const cableKeys   = Object.keys(metadata.cable).sort()
  const housingKeys = Object.keys(metadata.housing).sort()
  const cableMissing   = cableKeys  .filter(k => !getCableMeta(k).품번).length
  const housingMissing = housingKeys.filter(k => !getHousingMeta(k).품번).length

  // ── 템플릿 다운로드 ──────────────────────────────────────────
  function downloadTemplate(kind: 'cable' | 'housing') {
    const isCable = kind === 'cable'
    const keys    = isCable ? cableKeys : housingKeys
    const label   = isCable ? '케이블종류' : '하우징타입'

    const header = ['파이', label, '품번', '품명', '구매처', '리드타임(일)']
    const rows   = keys.map(key => {
      const [pai, type] = key.split('|')
      const m = isCable ? getCableMeta(key) : getHousingMeta(key)
      return [pai, type, m.품번 ?? '', m.품명 ?? '', m.구매처 ?? '', m.리드타임 ?? '']
    })

    // 데이터가 없으면 예시 행 추가
    if (rows.length === 0) rows.push(['A1', '예시종류', '', '', '', '60'])

    const ws = XLSX.utils.aoa_to_sheet([header, ...rows])
    ws['!cols'] = [10, 20, 16, 30, 16, 12].map(w => ({ wch: w }))

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, isCable ? '케이블' : '하우징')
    XLSX.writeFile(wb, `품번관리_${isCable ? '케이블' : '하우징'}_${today()}.xlsx`)
  }

  // ── Excel 파싱 ───────────────────────────────────────────────
  function handleUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    e.target.value = ''
    setPreviewErr('')

    const reader = new FileReader()
    reader.onload = ev => {
      try {
        const wb   = XLSX.read(ev.target?.result, { type: 'array' })
        const isCable = tab === 'cable'
        const sheetName = isCable ? '케이블' : '하우징'
        const ws   = wb.Sheets[sheetName] ?? wb.Sheets[wb.SheetNames[0]]
        if (!ws) { setPreviewErr('시트를 찾을 수 없습니다.'); return }

        const raw: string[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as string[][]
        if (raw.length < 2) { setPreviewErr('데이터 행이 없습니다.'); return }

        // 헤더 행(0번)은 건너뜀, 1번부터 데이터
        const parsed: PreviewRow[] = []
        for (let i = 1; i < raw.length; i++) {
          const r = raw[i]
          const pai  = String(r[0] ?? '').trim()
          const type = String(r[1] ?? '').trim()
          if (!pai || !type) continue
          const key = `${pai}|${type}`
          const exists = isCable ? !!metadata.cable[key] : !!metadata.housing[key]
          parsed.push({
            key,
            pai,
            type,
            품번:     String(r[2] ?? '').trim(),
            품명:     String(r[3] ?? '').trim(),
            구매처:   String(r[4] ?? '').trim(),
            리드타임: String(r[5] ?? '').trim(),
            isNew:    !exists,
          })
        }

        if (parsed.length === 0) { setPreviewErr('유효한 데이터가 없습니다.'); return }
        setPreview({ kind: tab, rows: parsed })
      } catch {
        setPreviewErr('파일을 읽는 중 오류가 발생했습니다.')
      }
    }
    reader.readAsArrayBuffer(file)
  }

  // ── 업로드 적용 ──────────────────────────────────────────────
  function applyPreview() {
    if (!preview) return
    const next = { ...metadata }

    if (preview.kind === 'cable') {
      const cable = { ...next.cable }
      for (const row of preview.rows) {
        cable[row.key] = { 품번: row.품번, 품명: row.품명, 구매처: row.구매처, 리드타임: row.리드타임 || null }
      }
      next.cable = cable
    } else {
      const housing = { ...next.housing }
      for (const row of preview.rows) {
        const cur = housing[row.key]
        const base = Array.isArray(cur) ? cur[0] : (cur || { 품번:'', 품명:'', 구매처:'', 리드타임:null })
        housing[row.key] = { ...base, 품번: row.품번, 품명: row.품명, 구매처: row.구매처, 리드타임: row.리드타임 || null }
      }
      next.housing = housing
    }

    setMetadata(next)
    saveMetadata(next)
    setPreview(null)
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>품번 관리</b>: STEP 1 실행 후 발견된 신규 타입이 자동으로 추가됩니다.
        품번·품명·구매처·리드타임을 직접 입력하거나 <b>Excel 템플릿을 다운로드</b>해서 채운 후 업로드하세요.
      </div>

      {/* 탭 + 업로드 버튼 */}
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="flex gap-2">
          <button onClick={() => { setTab('cable'); setPreview(null) }}
            className={`px-4 py-2 rounded font-semibold text-sm transition flex items-center gap-2 ${tab==='cable' ? 'bg-[#2E75B6] text-white' : 'bg-gray-100 text-gray-700 hover:bg-gray-200'}`}>
            케이블
            <span className={`text-xs px-1.5 py-0.5 rounded-full font-bold ${tab==='cable' ? 'bg-white/20' : 'bg-gray-300'}`}>{cableKeys.length}</span>
            {cableMissing > 0 && <span className="text-xs bg-red-500 text-white px-1.5 py-0.5 rounded-full">⚠ {cableMissing}</span>}
          </button>
          <button onClick={() => { setTab('housing'); setPreview(null) }}
            className={`px-4 py-2 rounded font-semibold text-sm transition flex items-center gap-2 ${tab==='housing' ? 'bg-[#2E75B6] text-white' : 'bg-gray-100 text-gray-700 hover:bg-gray-200'}`}>
            하우징
            <span className={`text-xs px-1.5 py-0.5 rounded-full font-bold ${tab==='housing' ? 'bg-white/20' : 'bg-gray-300'}`}>{housingKeys.length}</span>
            {housingMissing > 0 && <span className="text-xs bg-red-500 text-white px-1.5 py-0.5 rounded-full">⚠ {housingMissing}</span>}
          </button>
        </div>

        {/* Excel 버튼 */}
        <div className="flex gap-2">
          <button
            onClick={() => downloadTemplate(tab)}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border border-gray-300 rounded bg-white hover:bg-gray-50 transition"
          >
            📤 템플릿 다운로드
          </button>
          <button
            onClick={() => fileRef.current?.click()}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border border-[#2E75B6] text-[#2E75B6] rounded bg-white hover:bg-blue-50 transition"
          >
            📥 Excel 업로드
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUpload} />
        </div>
      </div>

      {previewErr && (
        <div className="text-sm text-red-600 bg-red-50 border border-red-200 px-4 py-2 rounded">{previewErr}</div>
      )}

      {/* 업로드 미리보기 */}
      {preview && (
        <div className="border border-blue-200 rounded-lg bg-blue-50 p-4 space-y-3">
          <div className="flex items-center justify-between">
            <div className="text-sm font-semibold text-blue-800">
              업로드 미리보기 — {preview.rows.length}개 행
              <span className="ml-2 text-xs font-normal text-blue-600">
                (기존 {preview.rows.filter(r => !r.isNew).length}개 업데이트 / 신규 {preview.rows.filter(r => r.isNew).length}개 추가)
              </span>
            </div>
            <button onClick={() => setPreview(null)} className="text-xs text-gray-400 hover:text-gray-600">✕ 취소</button>
          </div>
          <div className="overflow-x-auto max-h-60 overflow-y-auto">
            <table className="text-xs w-full bg-white rounded border border-blue-100">
              <thead className="sticky top-0">
                <tr className="bg-blue-100 text-blue-800">
                  {['파이', tab === 'cable' ? '케이블종류' : '하우징타입', '품번', '품명', '구매처', '리드타임', '구분'].map(h => (
                    <th key={h} className="px-3 py-1.5 text-left font-semibold">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {preview.rows.map(row => (
                  <tr key={row.key} className={row.isNew ? 'bg-green-50' : 'bg-white'}>
                    <td className="px-3 py-1 font-mono text-gray-500">{row.pai}</td>
                    <td className="px-3 py-1 font-semibold">{row.type}</td>
                    <td className="px-3 py-1">{row.품번 || <span className="text-red-400">미입력</span>}</td>
                    <td className="px-3 py-1">{row.품명 || <span className="text-gray-300">—</span>}</td>
                    <td className="px-3 py-1">{row.구매처 || <span className="text-gray-300">—</span>}</td>
                    <td className="px-3 py-1">{row.리드타임 || <span className="text-gray-300">—</span>}</td>
                    <td className="px-3 py-1">
                      {row.isNew
                        ? <span className="text-green-600 font-semibold">신규</span>
                        : <span className="text-blue-600">업데이트</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button
            onClick={applyPreview}
            className="bg-[#2E75B6] hover:bg-[#1F5A9E] text-white font-bold px-5 py-2 rounded-lg text-sm transition"
          >
            ✅ 적용 및 저장
          </button>
        </div>
      )}

      {/* 케이블 테이블 */}
      {tab === 'cable' && (
        <div className="overflow-x-auto">
          <table className="text-xs w-full border-collapse">
            <thead>
              <tr className="bg-[#2E75B6] text-white">
                {['파이','케이블종류','품번','품명','구매처','리드타임(일)','상태'].map(h => (
                  <th key={h} className="px-3 py-2 text-left font-semibold border-r border-white/20 last:border-0">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cableKeys.map((key, i) => {
                const [pai, ct] = key.split('|')
                const m = getCableMeta(key)
                const missing = !m.품번
                return (
                  <tr key={key} className={`border-b ${missing ? 'bg-red-50' : i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}`}>
                    <td className="px-3 py-1 font-mono text-gray-500 whitespace-nowrap">{pai}</td>
                    <td className="px-3 py-1 font-semibold whitespace-nowrap">{ct}</td>
                    {(['품번','품명','구매처','리드타임'] as const).map(f => (
                      <td key={f} className="px-1 py-1">
                        <input
                          value={String(m[f] ?? '')}
                          onChange={e => updateCable(key, f, e.target.value)}
                          className={`w-full border rounded px-2 py-0.5 text-xs focus:outline-blue-400 ${missing && f==='품번' ? 'border-red-400 bg-red-50' : ''}`}
                          placeholder={f === '리드타임' ? '숫자(일)' : f}
                        />
                      </td>
                    ))}
                    <td className="px-2 py-1 text-center">
                      {missing ? <span className="text-red-500 font-semibold">⚠ 미입력</span> : <span className="text-green-600">✓</span>}
                    </td>
                  </tr>
                )
              })}
              {cableKeys.length === 0 && (
                <tr><td colSpan={7} className="text-center text-gray-400 py-8">STEP 1을 실행하면 케이블 목록이 나타납니다.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* 하우징 테이블 */}
      {tab === 'housing' && (
        <div className="overflow-x-auto">
          <table className="text-xs w-full border-collapse">
            <thead>
              <tr className="bg-[#2E75B6] text-white">
                {['파이','하우징타입','품번','품명','구매처','리드타임(일)','상태'].map(h => (
                  <th key={h} className="px-3 py-2 text-left font-semibold border-r border-white/20 last:border-0">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {housingKeys.map((key, i) => {
                const [pai, ht] = key.split('|')
                const m = getHousingMeta(key)
                const missing = !m.품번
                return (
                  <tr key={key} className={`border-b ${missing ? 'bg-red-50' : i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}`}>
                    <td className="px-3 py-1 font-mono text-gray-500 whitespace-nowrap">{pai}</td>
                    <td className="px-3 py-1 font-semibold whitespace-nowrap">{ht}</td>
                    {(['품번','품명','구매처','리드타임'] as const).map(f => (
                      <td key={f} className="px-1 py-1">
                        <input
                          value={String(m[f] ?? '')}
                          onChange={e => updateHousing(key, f, e.target.value)}
                          className={`w-full border rounded px-2 py-0.5 text-xs focus:outline-blue-400 ${missing && f==='품번' ? 'border-red-400 bg-red-50' : ''}`}
                          placeholder={f === '리드타임' ? '숫자(일)' : f}
                        />
                      </td>
                    ))}
                    <td className="px-2 py-1 text-center">
                      {missing ? <span className="text-red-500 font-semibold">⚠ 미입력</span> : <span className="text-green-600">✓</span>}
                    </td>
                  </tr>
                )
              })}
              {housingKeys.length === 0 && (
                <tr><td colSpan={7} className="text-center text-gray-400 py-8">STEP 1을 실행하면 하우징 목록이 나타납니다.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      <div className="flex items-center gap-3 pt-1">
        <button onClick={save}
          className="bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold px-6 py-2 rounded-lg transition">
          💾 품번 저장
        </button>
        {saved && <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>}
      </div>
    </div>
  )
}
