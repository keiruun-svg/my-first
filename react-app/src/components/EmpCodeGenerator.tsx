import { useState, useEffect, useRef } from 'react'
import { loadVendors, upsertVendor, deleteVendor, loadItemCodes, saveItemCodes } from '../lib/supabase'
import type { Vendor, ItemCode } from '../lib/supabase'
import { parseItemCodes } from '../lib/parse/parseItemCodes'

function buildEmpCode(itemCode: string, vendorCode: string, yymm: string, orderSeq: string, receiptSeq: string): string {
  if (!itemCode || !vendorCode || !yymm || !orderSeq || !receiptSeq) return ''
  return `${itemCode} (${vendorCode}-${yymm}-${orderSeq}_${receiptSeq})`
}

export default function EmpCodeGenerator() {
  const [vendors, setVendors]           = useState<Vendor[]>([])
  const [itemCodes, setItemCodes]       = useState<ItemCode[]>([])
  const [itemCode, setItemCode]         = useState('')
  const [itemSearch, setItemSearch]     = useState('')
  const [showDropdown, setShowDropdown] = useState(false)
  const [vendorCode, setVendorCode]     = useState('')
  const [yymm, setYymm]                 = useState('')
  const [orderSeq, setOrderSeq]         = useState('01')
  const [receiptSeq, setReceiptSeq]     = useState('01')
  const [copied, setCopied]             = useState(false)
  const [history, setHistory]           = useState<string[]>([])

  const [uploadMsg, setUploadMsg]       = useState<string | null>(null)
  const fileInputRef                    = useRef<HTMLInputElement>(null)

  const [editingCode, setEditingCode]   = useState<string | null>(null)
  const [newCode, setNewCode]           = useState('')
  const [newName, setNewName]           = useState('')
  const [addMode, setAddMode]           = useState(false)
  const [addCode, setAddCode]           = useState('')
  const [addName, setAddName]           = useState('')
  const [vendorMsg, setVendorMsg]       = useState('')

  useEffect(() => {
    loadVendors().then(setVendors)
    loadItemCodes().then(setItemCodes)
    const now = new Date()
    const yy  = String(now.getFullYear()).slice(2)
    const mm  = String(now.getMonth() + 1).padStart(2, '0')
    setYymm(`${yy}${mm}`)
  }, [])

  const filtered = itemSearch.trim()
    ? itemCodes.filter(c =>
        c.code.toLowerCase().includes(itemSearch.toLowerCase()) ||
        c.name.toLowerCase().includes(itemSearch.toLowerCase())
      )
    : itemCodes.slice(0, 50)

  const seqs    = Array.from({ length: 99 }, (_, i) => String(i + 1).padStart(2, '0'))
  const recSeqs = ['01', '02', '03', '04', '05']
  const result  = buildEmpCode(itemCode, vendorCode, yymm, orderSeq, receiptSeq)

  const handleCopy = () => {
    if (!result) return
    navigator.clipboard.writeText(result)
    setCopied(true)
    setHistory(prev => [result, ...prev.filter(h => h !== result)].slice(0, 10))
    setTimeout(() => setCopied(false), 1500)
  }

  const handleItemUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setUploadMsg('⏳ 파싱 중...')
    try {
      const codes = await parseItemCodes(file)
      setItemCodes(codes)
      saveItemCodes(codes)
      setUploadMsg(`✅ ${codes.length}개 품목 등록됨`)
    } catch (err) {
      setUploadMsg(`❌ ${err}`)
    } finally {
      e.target.value = ''
      setTimeout(() => setUploadMsg(null), 3000)
    }
  }

  const handleVendorSave = async (code: string, name: string) => {
    if (!code.trim() || !name.trim()) return
    const v: Vendor = { code: code.trim().toUpperCase(), name: name.trim() }
    await upsertVendor(v)
    setVendors(prev => {
      const next = prev.filter(x => x.code !== v.code)
      return [...next, v].sort((a, b) => a.code.localeCompare(b.code))
    })
    setEditingCode(null); setNewCode(''); setNewName('')
    setAddMode(false); setAddCode(''); setAddName('')
    setVendorMsg(`✅ ${v.code} 저장됨`)
    setTimeout(() => setVendorMsg(''), 2000)
  }

  const handleVendorDelete = async (code: string) => {
    if (!confirm(`${code} 업체코드를 삭제할까요?`)) return
    await deleteVendor(code)
    setVendors(prev => prev.filter(x => x.code !== code))
    if (vendorCode === code) setVendorCode('')
    setVendorMsg(`🗑 ${code} 삭제됨`)
    setTimeout(() => setVendorMsg(''), 2000)
  }

  return (
    <div className="space-y-5">

      {/* 이카운트 품목 리스트 업로드 */}
      <div className="flex items-center gap-4 bg-white border border-gray-200 rounded-lg px-5 py-3">
        <div className="flex-1">
          <div className="text-sm font-medium text-gray-700">이카운트 품목 리스트</div>
          <div className="text-xs text-gray-400 mt-0.5">
            {itemCodes.length > 0
              ? `${itemCodes.length}개 품목 등록됨 — 품번 생성기 중복 확인에도 사용됩니다`
              : '이카운트 품목 전체 내보내기(.xlsx)를 업로드하면 기본코드 드롭다운과 중복 확인이 활성화됩니다'}
          </div>
        </div>
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleItemUpload} />
        <button
          onClick={() => fileInputRef.current?.click()}
          className="flex-shrink-0 px-4 py-2 text-sm font-semibold rounded-lg bg-[#1F3864] hover:bg-[#162a4a] text-white transition"
        >
          📂 파일 업로드
        </button>
        {uploadMsg && (
          <span className={`text-xs font-medium ${uploadMsg.startsWith('✅') ? 'text-green-700' : uploadMsg.startsWith('❌') ? 'text-red-600' : 'text-gray-500'}`}>
            {uploadMsg}
          </span>
        )}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">

        {/* ── EMP 코드 생성기 ── */}
        <div className="border border-gray-200 rounded-lg p-5 space-y-4">
          <div className="text-xs text-gray-500">
            EMP 로트 코드 형식: <code className="bg-gray-100 px-1 rounded">기본코드 (업체코드-YYMM-발주차수_입고회차)</code>
          </div>

          {/* 기본코드 */}
          <div className="relative">
            <label className="block text-xs font-medium text-gray-600 mb-1">① 기본코드 (이카운트 품번)</label>
            <input
              type="text"
              placeholder={itemCodes.length === 0 ? '품목 리스트를 먼저 업로드하세요' : '코드 또는 품명으로 검색...'}
              value={itemSearch}
              onChange={e => { setItemSearch(e.target.value); setItemCode(''); setShowDropdown(true) }}
              onFocus={() => setShowDropdown(true)}
              onBlur={() => setTimeout(() => setShowDropdown(false), 150)}
              className="w-full border border-gray-300 rounded px-3 py-1.5 text-sm focus:outline-none focus:border-sky-500"
              disabled={itemCodes.length === 0}
            />
            {showDropdown && itemSearch.trim() && (
              <div className="absolute z-10 left-0 right-0 border border-gray-200 rounded-b max-h-44 overflow-y-auto bg-white shadow-md">
                {filtered.length === 0 ? (
                  <div className="px-3 py-2 text-xs text-gray-400">검색 결과 없음</div>
                ) : (
                  filtered.map(c => (
                    <button
                      key={c.code}
                      onMouseDown={() => { setItemCode(c.code); setItemSearch(c.code); setShowDropdown(false) }}
                      className={`w-full text-left px-3 py-1.5 text-xs hover:bg-[#f0f4fa] ${itemCode === c.code ? 'bg-[#e8f4fd] font-semibold text-sky-700' : ''}`}
                    >
                      <span className="font-mono font-semibold">{c.code}</span>
                      {c.name && <span className="text-gray-500 ml-2">{c.name}</span>}
                    </button>
                  ))
                )}
              </div>
            )}
            {itemCode && <div className="mt-1 text-xs text-sky-700 font-semibold">선택됨: {itemCode}</div>}
          </div>

          {/* 업체코드 */}
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">② 업체코드</label>
            <div className="flex gap-2 flex-wrap">
              {vendors.map(v => (
                <button
                  key={v.code}
                  onClick={() => setVendorCode(v.code)}
                  className={`px-3 py-1.5 rounded border text-sm font-medium transition ${
                    vendorCode === v.code
                      ? 'bg-sky-600 text-white border-sky-600'
                      : 'bg-white text-gray-700 border-gray-300 hover:border-sky-500'
                  }`}
                  title={v.name}
                >
                  {v.code}
                </button>
              ))}
              {vendors.length === 0 && (
                <span className="text-xs text-gray-400">등록된 업체코드가 없습니다. 우측에서 추가하세요.</span>
              )}
            </div>
          </div>

          {/* 발주연월 / 발주차수 / 입고회차 */}
          <div className="grid grid-cols-3 gap-3">
            <div>
              <label className="block text-xs font-medium text-gray-600 mb-1">③ 발주연월 (YYMM)</label>
              <input
                type="text" maxLength={4} placeholder="2601" value={yymm}
                onChange={e => setYymm(e.target.value.replace(/\D/g, '').slice(0, 4))}
                className="w-full border border-gray-300 rounded px-3 py-1.5 text-sm text-center font-mono focus:outline-none focus:border-sky-500"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-600 mb-1">④ 발주차수</label>
              <select value={orderSeq} onChange={e => setOrderSeq(e.target.value)}
                className="w-full border border-gray-300 rounded px-2 py-1.5 text-sm text-center focus:outline-none focus:border-sky-500">
                {seqs.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-600 mb-1">⑤ 입고회차</label>
              <select value={receiptSeq} onChange={e => setReceiptSeq(e.target.value)}
                className="w-full border border-gray-300 rounded px-2 py-1.5 text-sm text-center focus:outline-none focus:border-sky-500">
                {recSeqs.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>

          {/* 결과 미리보기 */}
          <div className={`rounded-lg border-2 p-4 transition-colors ${result ? 'border-sky-500 bg-sky-50' : 'border-dashed border-gray-300 bg-gray-50'}`}>
            <div className="text-xs text-gray-500 mb-1">생성 결과</div>
            <div className={`font-mono text-base font-bold ${result ? 'text-sky-900' : 'text-gray-400'}`}>
              {result || '항목을 모두 선택하면 코드가 생성됩니다'}
            </div>
          </div>

          <button
            onClick={handleCopy} disabled={!result}
            className="w-full py-2 rounded bg-sky-600 hover:bg-sky-700 disabled:bg-gray-300 text-white text-sm font-semibold transition"
          >
            {copied ? '✅ 복사됨!' : '📋 복사'}
          </button>

          {history.length > 0 && (
            <div>
              <div className="text-xs font-medium text-gray-500 mb-1">최근 생성 이력 (세션)</div>
              <div className="space-y-1 max-h-32 overflow-y-auto">
                {history.map((h, i) => (
                  <button key={i}
                    onClick={() => { navigator.clipboard.writeText(h); setCopied(true); setTimeout(() => setCopied(false), 1500) }}
                    className="w-full text-left px-2 py-1 text-xs font-mono bg-gray-50 hover:bg-sky-50 border border-gray-200 rounded"
                  >
                    {h}
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* ── 업체코드 관리 ── */}
        <div className="border border-gray-200 rounded-lg p-5 space-y-3">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-semibold text-[#1F3864]">🏢 업체코드 관리</h2>
            {vendorMsg && <span className="text-xs text-green-700">{vendorMsg}</span>}
          </div>

          <table className="w-full text-sm border-collapse">
            <thead>
              <tr className="bg-[#1F3864] text-white">
                <th className="px-3 py-2 text-left text-xs font-medium w-24">코드</th>
                <th className="px-3 py-2 text-left text-xs font-medium">업체명</th>
                <th className="px-3 py-2 w-24"></th>
              </tr>
            </thead>
            <tbody>
              {vendors.map((v, i) => (
                editingCode === v.code ? (
                  <tr key={v.code} className="bg-yellow-50">
                    <td className="px-2 py-1">
                      <input value={newCode} onChange={e => setNewCode(e.target.value.toUpperCase())}
                        className="w-full border border-gray-300 rounded px-2 py-1 text-xs font-mono focus:outline-none" />
                    </td>
                    <td className="px-2 py-1">
                      <input value={newName} onChange={e => setNewName(e.target.value)}
                        className="w-full border border-gray-300 rounded px-2 py-1 text-xs focus:outline-none"
                        onKeyDown={e => { if (e.key === 'Enter') handleVendorSave(newCode, newName) }} />
                    </td>
                    <td className="px-2 py-1">
                      <div className="flex gap-1">
                        <button onClick={() => handleVendorSave(newCode, newName)}
                          className="px-2 py-1 text-xs bg-green-700 text-white rounded hover:bg-green-800">저장</button>
                        <button onClick={() => setEditingCode(null)}
                          className="px-2 py-1 text-xs border border-gray-300 rounded hover:bg-gray-100">취소</button>
                      </div>
                    </td>
                  </tr>
                ) : (
                  <tr key={v.code} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="px-3 py-2 font-mono font-bold text-sky-700 text-xs">{v.code}</td>
                    <td className="px-3 py-2 text-gray-700 text-xs">{v.name}</td>
                    <td className="px-3 py-2">
                      <div className="flex gap-1">
                        <button onClick={() => { setEditingCode(v.code); setNewCode(v.code); setNewName(v.name) }}
                          className="px-2 py-1 text-xs border border-gray-300 rounded hover:bg-gray-100">수정</button>
                        <button onClick={() => handleVendorDelete(v.code)}
                          className="px-2 py-1 text-xs border border-red-200 text-red-600 rounded hover:bg-red-50">삭제</button>
                      </div>
                    </td>
                  </tr>
                )
              ))}
              {addMode && (
                <tr className="bg-sky-50">
                  <td className="px-2 py-1">
                    <input autoFocus placeholder="FLC" value={addCode}
                      onChange={e => setAddCode(e.target.value.toUpperCase())}
                      className="w-full border border-gray-300 rounded px-2 py-1 text-xs font-mono focus:outline-none" />
                  </td>
                  <td className="px-2 py-1">
                    <input placeholder="업체명" value={addName} onChange={e => setAddName(e.target.value)}
                      className="w-full border border-gray-300 rounded px-2 py-1 text-xs focus:outline-none"
                      onKeyDown={e => { if (e.key === 'Enter') handleVendorSave(addCode, addName) }} />
                  </td>
                  <td className="px-2 py-1">
                    <div className="flex gap-1">
                      <button onClick={() => handleVendorSave(addCode, addName)}
                        className="px-2 py-1 text-xs bg-sky-600 text-white rounded hover:bg-sky-700">추가</button>
                      <button onClick={() => { setAddMode(false); setAddCode(''); setAddName('') }}
                        className="px-2 py-1 text-xs border border-gray-300 rounded hover:bg-gray-100">취소</button>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>

          {!addMode && (
            <button onClick={() => { setAddMode(true); setEditingCode(null) }}
              className="text-sm text-sky-600 hover:underline">
              + 업체코드 추가
            </button>
          )}

          <div className="bg-gray-50 rounded p-3 text-xs text-gray-500 space-y-0.5 mt-2">
            <div className="font-semibold text-gray-600 mb-1">예시</div>
            <div><code className="bg-white px-1 rounded border">14-K-107 (FLC-2601-01_02)</code></div>
            <div className="text-gray-400 mt-1">FLC = 업체, 2601 = 26년 1월, 01 = 첫 번째 발주, 02 = 두 번째 입고</div>
          </div>
        </div>
      </div>
    </div>
  )
}
