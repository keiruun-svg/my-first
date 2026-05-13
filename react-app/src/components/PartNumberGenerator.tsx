import { useState, useEffect } from 'react'
import type { OjcRules } from '../lib/types'
import { DEFAULT_OJC_RULES } from '../lib/types'
import { loadOjcRules, saveOjcRules } from '../lib/supabase'
import type { DetectedFields } from '../lib/ojcAutoDetect'
import { analyzeProduct } from '../lib/ojcAutoDetect'

type SubTab = 'auto' | 'manual' | 'rules'
type Section = keyof OjcRules

const SECTION_LABELS: Record<Section, string> = {
  ojcType:      '① OJC 종류',
  coreType:     '② 코어 종류',
  coreCount:    '③ 심선수',
  material:     '④ 재질',
  connector:    '⑤⑥ 커넥터',
  marking:      '⑧ 마킹',
  lengthPreset: '⑦ 길이 코드표',
}

function lengthToCodeAlgo(m: number): string | null {
  if ([100, 150, 200, 250, 300].includes(m)) return String(m)
  const dec = (d: number) => String.fromCharCode(64 + d)
  if (m < 1) {
    const d = Math.round(m * 10)
    return (d >= 1 && d <= 9) ? `00${dec(d)}` : null
  }
  if (m < 10) {
    const i = Math.floor(m), d = Math.round((m - i) * 10)
    return d === 0 ? `00${i}` : `0${i}${dec(d)}`
  }
  if (m < 100) {
    const i = Math.floor(m), d = Math.round((m - i) * 10)
    return d === 0 ? `${String(i).padStart(2, '0')}J` : `${String(i).padStart(2, '0')}${dec(d)}`
  }
  return null
}

const inputCls = 'w-full border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:border-sky-500'
const selectCls = `${inputCls} bg-white`

// 각 필드에 표시할 라벨
const FIELD_LABELS: Record<keyof DetectedFields, string> = {
  ojcType:    '① OJC 종류',
  coreType:   '② 코어 종류',
  coreCount:  '③ 심선수',
  material:   '④ 재질',
  connectorA: '⑤ 커넥터 A端',
  connectorB: '⑥ 커넥터 B端',
  lengthCode: '⑦ 길이',
  marking:    '⑧ 마킹',
}

export default function PartNumberGenerator() {
  const [subTab, setSubTab]         = useState<SubTab>('auto')
  const [rules, setRules]           = useState<OjcRules>(DEFAULT_OJC_RULES)
  const [editRules, setEditRules]   = useState<OjcRules>(DEFAULT_OJC_RULES)
  const [hasChanges, setHasChanges] = useState(false)
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle')
  const [copied, setCopied]         = useState(false)
  const [openSections, setOpenSections] = useState<Partial<Record<Section, boolean>>>({ ojcType: true })

  // ── 자동 분석 탭 상태 ─────────────────────────────────────
  const [autoName,    setAutoName]    = useState('')
  const [autoSpec,    setAutoSpec]    = useState('')
  const [autoDiam,    setAutoDiam]    = useState('')
  const [detected,    setDetected]    = useState<DetectedFields | null>(null)
  const [overrides,   setOverrides]   = useState<Partial<DetectedFields>>({})
  const [needsRerun,  setNeedsRerun]  = useState(false)

  // ── 수동 선택 탭 상태 ─────────────────────────────────────
  const [ojcType,    setOjcType]    = useState('')
  const [coreType,   setCoreType]   = useState('')
  const [coreCount,  setCoreCount]  = useState('')
  const [material,   setMaterial]   = useState('')
  const [connectorA, setConnectorA] = useState('')
  const [connectorB, setConnectorB] = useState('')
  const [connBSame,  setConnBSame]  = useState(false)
  const [lenIdx,     setLenIdx]     = useState('')
  const [lenCustom,  setLenCustom]  = useState('')
  const [marking,    setMarking]    = useState('N')

  useEffect(() => {
    loadOjcRules().then(r => { setRules(r); setEditRules(r) })
  }, [])

  // ── 자동 분석 ──────────────────────────────────────────────
  function handleAnalyze() {
    if (!autoName.trim()) return
    const diamMm = autoDiam ? parseFloat(autoDiam) : null
    setDetected(analyzeProduct(autoName, autoSpec, diamMm, rules))
    setOverrides({})
    setNeedsRerun(false)
  }

  function effField(key: keyof DetectedFields): string {
    return (overrides[key] as string | null | undefined) ?? detected?.[key] ?? ''
  }

  function setOverride(key: keyof DetectedFields, val: string) {
    setOverrides(o => ({ ...o, [key]: val || undefined }))
  }

  // 자동 탭 품번
  const autoPartNumber = (() => {
    if (!detected) return null
    const ot = effField('ojcType'),  ct = effField('coreType')
    const cc = effField('coreCount'), mt = effField('material')
    const ca = effField('connectorA'), cb = effField('connectorB')
    const lc = effField('lengthCode'), mk = effField('marking')
    if (!ot || !ct || !cc || !mt || !ca || !lc || !mk) return null
    return `A14-${ot}-${ct}${cc}${mt}-${ca}${cb}-${lc}${mk}`
  })()

  const autoMissing = !detected ? [] : (
    Object.keys(FIELD_LABELS) as (keyof DetectedFields)[]
  ).filter(k => k !== 'connectorB' && !effField(k))

  // 감지 결과 표시 텍스트
  function detectedLabel(key: keyof DetectedFields): string {
    const raw = detected?.[key]
    if (!raw) return ''
    if (key === 'ojcType')    return rules.ojcType.find(e => e.code === raw)?.label ?? raw
    if (key === 'coreType')   return rules.coreType.find(e => e.code === raw)?.label ?? raw
    if (key === 'coreCount')  return `${rules.coreCount.find(e => e.code === raw)?.count ?? '?'}심`
    if (key === 'material')   return rules.material.find(e => e.code === raw)?.label ?? raw
    if (key === 'connectorA' || key === 'connectorB') return rules.connector.find(e => e.code === raw)?.label ?? raw
    if (key === 'lengthCode') {
      const preset = rules.lengthPreset.find(p => p.code === raw)
      return preset ? `${preset.m}M` : raw
    }
    if (key === 'marking') return rules.marking.find(e => e.code === raw)?.label ?? raw
    return raw
  }

  // ── 수동 탭 길이 ───────────────────────────────────────────
  const manualLengthCode = (() => {
    if (!lenIdx) return null
    if (lenIdx === 'custom') {
      const m = parseFloat(lenCustom)
      if (isNaN(m) || m <= 0) return null
      return rules.lengthPreset.find(p => p.m === m)?.code ?? lengthToCodeAlgo(m)
    }
    return rules.lengthPreset[parseInt(lenIdx)]?.code ?? null
  })()

  const manualConnB = connBSame ? connectorA : connectorB

  const manualPartNumber = (() => {
    if (!ojcType || !coreType || !coreCount || !material || !connectorA || !manualLengthCode || !marking) return null
    return `A14-${ojcType}-${coreType}${coreCount}${material}-${connectorA}${manualConnB}-${manualLengthCode}${marking}`
  })()

  const manualMissing = [
    !ojcType          && '① OJC 종류',
    !coreType         && '② 코어 종류',
    !coreCount        && '③ 심선수',
    !material         && '④ 재질',
    !connectorA       && '⑤ 커넥터 A端',
    !manualLengthCode && '⑦ 길이',
    !marking          && '⑧ 마킹',
  ].filter(Boolean) as string[]

  function handleCopy(pn: string) {
    navigator.clipboard.writeText(pn)
    setCopied(true)
    setTimeout(() => setCopied(false), 2000)
  }

  // ── 규칙 편집 헬퍼 ────────────────────────────────────────
  type AnyEntry = Record<string, string | number>

  function updateRule(section: Section, idx: number, key: string, value: string | number) {
    setEditRules(r => {
      const arr = [...(r[section] as unknown as AnyEntry[])]
      arr[idx] = { ...arr[idx], [key]: value }
      return { ...r, [section]: arr }
    })
    setHasChanges(true)
  }

  function addRow(section: Section) {
    setEditRules(r => {
      let row: AnyEntry
      if (section === 'coreCount')     row = { code: '', count: 0 }
      else if (section === 'lengthPreset') row = { m: 0, code: '' }
      else                             row = { code: '', label: '' }
      return { ...r, [section]: [...(r[section] as unknown as AnyEntry[]), row] }
    })
    setHasChanges(true)
  }

  function deleteRow(section: Section, idx: number) {
    setEditRules(r => ({
      ...r,
      [section]: (r[section] as unknown as AnyEntry[]).filter((_, i) => i !== idx),
    }))
    setHasChanges(true)
  }

  async function handleSave() {
    setSaveStatus('saving')
    try {
      await saveOjcRules(editRules)
      setRules(editRules)
      setHasChanges(false)
      setSaveStatus('saved')
      setTimeout(() => setSaveStatus('idle'), 2000)
    } catch {
      setSaveStatus('error')
      setTimeout(() => setSaveStatus('idle'), 3000)
    }
  }

  function handleReset() {
    if (!confirm('모든 코드표를 기본값으로 초기화하시겠습니까?')) return
    setEditRules(DEFAULT_OJC_RULES)
    setHasChanges(true)
  }

  // ── 렌더 ──────────────────────────────────────────────────
  return (
    <div>
      {/* 서브탭 */}
      <div className="flex border-b mb-6">
        {([['auto', '자동 분석'], ['manual', '수동 선택'], ['rules', '규칙 편집']] as [SubTab, string][]).map(([id, lbl]) => (
          <button
            key={id}
            onClick={() => setSubTab(id)}
            className={`px-5 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${
              subTab === id
                ? 'border-sky-500 text-sky-600'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            {lbl}{id === 'rules' && hasChanges && <span className="ml-1 text-orange-500">●</span>}
          </button>
        ))}
      </div>

      {/* ── 자동 분석 탭 ──────────────────────────────────── */}
      {subTab === 'auto' && (
        <div className="max-w-2xl space-y-5">
          {/* 입력 영역 */}
          <div className="bg-white rounded-lg border border-gray-200 p-5">
            <h2 className="text-sm font-semibold text-gray-700 mb-4">
              품목명과 규격을 입력하면 품번을 자동으로 분석합니다.
            </h2>
            <div className="space-y-3">
              <div className="flex items-center gap-3">
                <span className="text-sm font-medium text-gray-700 w-16 flex-shrink-0">품목명</span>
                <input
                  value={autoName}
                  onChange={e => { setAutoName(e.target.value); setNeedsRerun(!!detected) }}
                  onKeyDown={e => e.key === 'Enter' && handleAnalyze()}
                  placeholder="예: DOJC-SM-LC/PC-LC/PC"
                  className={inputCls}
                />
              </div>
              <div className="flex items-center gap-3">
                <span className="text-sm font-medium text-gray-700 w-16 flex-shrink-0">규격</span>
                <input
                  value={autoSpec}
                  onChange={e => { setAutoSpec(e.target.value); setNeedsRerun(!!detected) }}
                  onKeyDown={e => e.key === 'Enter' && handleAnalyze()}
                  placeholder="예: 17M, B3, 2.0mm"
                  className={inputCls}
                />
              </div>
              <div className="flex items-center gap-3">
                <span className="text-sm font-medium text-gray-700 w-16 flex-shrink-0">직경(mm)</span>
                <input
                  type="number" step="0.1" min="0"
                  value={autoDiam}
                  onChange={e => { setAutoDiam(e.target.value); setNeedsRerun(!!detected) }}
                  placeholder="예: 2.0  (선택)"
                  className="w-36 border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:border-sky-500"
                />
                <button
                  onClick={handleAnalyze}
                  disabled={!autoName.trim()}
                  className={`ml-auto px-5 py-1.5 text-sm font-medium rounded transition-colors ${
                    autoName.trim()
                      ? needsRerun
                        ? 'bg-orange-500 text-white hover:bg-orange-600'
                        : 'bg-sky-600 text-white hover:bg-sky-700'
                      : 'bg-gray-100 text-gray-400 cursor-not-allowed'
                  }`}
                >
                  {needsRerun ? '재분석' : '분석'}
                </button>
              </div>
            </div>
          </div>

          {/* 분석 결과 */}
          {detected && (
            <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
              <div className="px-5 py-3 border-b border-gray-100 bg-gray-50 text-xs text-gray-500">
                감지된 항목은 자동으로 선택됩니다. 잘못된 항목은 드롭다운으로 수정하세요.
              </div>
              <div className="divide-y divide-gray-50">
                {(Object.keys(FIELD_LABELS) as (keyof DetectedFields)[]).map(key => {
                  const raw = detected[key]
                  const effectiveVal = effField(key)
                  const isOverridden = !!overrides[key]
                  return (
                    <div key={key} className="flex items-center px-4 py-2.5 gap-3">
                      {/* 항목명 */}
                      <span className="text-xs font-medium text-gray-500 w-28 flex-shrink-0">
                        {FIELD_LABELS[key]}
                      </span>
                      {/* 감지 결과 뱃지 */}
                      <div className="w-48 flex-shrink-0">
                        {raw ? (
                          <span className={`inline-flex items-center gap-1 text-xs px-2 py-0.5 rounded-full ${
                            isOverridden
                              ? 'bg-blue-50 text-blue-700'
                              : 'bg-green-50 text-green-700'
                          }`}>
                            <span className="font-mono font-bold">{raw}</span>
                            <span className="text-gray-500">— {detectedLabel(key)}</span>
                            {isOverridden && <span className="text-blue-400">(수정됨)</span>}
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 text-xs px-2 py-0.5 rounded-full bg-orange-50 text-orange-600">
                            ⚠ 감지 안됨
                          </span>
                        )}
                      </div>
                      {/* 수동 변경 드롭다운 */}
                      <div className="flex-1">
                        {key === 'lengthCode' ? (
                          <select
                            value={effectiveVal}
                            onChange={e => setOverride(key, e.target.value)}
                            className={`${selectCls} text-xs`}
                          >
                            <option value="">— 변경 —</option>
                            {rules.lengthPreset.map(p => (
                              <option key={p.code} value={p.code}>{p.m}M ({p.code})</option>
                            ))}
                          </select>
                        ) : key === 'coreCount' ? (
                          <select
                            value={effectiveVal}
                            onChange={e => setOverride(key, e.target.value)}
                            className={`${selectCls} text-xs`}
                          >
                            <option value="">— 변경 —</option>
                            {rules.coreCount.map(e => (
                              <option key={e.code} value={e.code}>{e.code} — {e.count}심</option>
                            ))}
                          </select>
                        ) : key === 'connectorA' || key === 'connectorB' ? (
                          <select
                            value={effectiveVal}
                            onChange={e => setOverride(key, e.target.value)}
                            className={`${selectCls} text-xs`}
                          >
                            <option value="">— 변경 —</option>
                            {rules.connector.map(e => (
                              <option key={e.code} value={e.code}>{e.code} — {e.label}</option>
                            ))}
                          </select>
                        ) : key === 'ojcType' ? (
                          <select value={effectiveVal} onChange={e => setOverride(key, e.target.value)} className={`${selectCls} text-xs`}>
                            <option value="">— 변경 —</option>
                            {rules.ojcType.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                          </select>
                        ) : key === 'coreType' ? (
                          <select value={effectiveVal} onChange={e => setOverride(key, e.target.value)} className={`${selectCls} text-xs`}>
                            <option value="">— 변경 —</option>
                            {rules.coreType.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                          </select>
                        ) : key === 'material' ? (
                          <select value={effectiveVal} onChange={e => setOverride(key, e.target.value)} className={`${selectCls} text-xs`}>
                            <option value="">— 변경 —</option>
                            {rules.material.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                          </select>
                        ) : key === 'marking' ? (
                          <select value={effectiveVal} onChange={e => setOverride(key, e.target.value)} className={`${selectCls} text-xs`}>
                            <option value="">— 변경 —</option>
                            {rules.marking.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                          </select>
                        ) : null}
                      </div>
                    </div>
                  )
                })}
              </div>

              {/* 품번 결과 */}
              <div className="px-5 py-4 border-t border-gray-200 bg-gray-50">
                {autoPartNumber ? (
                  <div className="flex items-center gap-3">
                    <span className="font-mono text-xl font-bold text-sky-700 tracking-widest flex-1 break-all">
                      {autoPartNumber}
                    </span>
                    <button
                      onClick={() => handleCopy(autoPartNumber)}
                      className="px-4 py-1.5 text-sm bg-sky-600 text-white rounded hover:bg-sky-700 transition-colors whitespace-nowrap"
                    >
                      {copied ? '✓ 복사됨' : '복사'}
                    </button>
                  </div>
                ) : (
                  <p className="text-sm text-orange-500">
                    미확인 항목: {autoMissing.join(' / ')}
                  </p>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── 수동 선택 탭 ──────────────────────────────────── */}
      {subTab === 'manual' && (
        <div className="max-w-2xl">
          <div className="bg-white rounded-lg border border-gray-200 p-6">
            <h2 className="text-base font-semibold text-gray-800 mb-5">
              수동 선택
              <span className="ml-2 text-xs font-normal text-gray-400">
                A14-[①]-[②③④]-[⑤⑥]-[⑦⑧]
              </span>
            </h2>
            <div className="space-y-3.5">
              <Row label="① OJC 종류">
                <select value={ojcType} onChange={e => setOjcType(e.target.value)} className={selectCls}>
                  <option value="">— 선택 —</option>
                  {rules.ojcType.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                </select>
              </Row>
              <Row label="② 코어 종류">
                <select value={coreType} onChange={e => setCoreType(e.target.value)} className={selectCls}>
                  <option value="">— 선택 —</option>
                  {rules.coreType.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                </select>
              </Row>
              <Row label="③ 심선수">
                <select value={coreCount} onChange={e => setCoreCount(e.target.value)} className={selectCls}>
                  <option value="">— 선택 —</option>
                  {rules.coreCount.map(e => <option key={e.code} value={e.code}>{e.code} — {e.count}심</option>)}
                </select>
              </Row>
              <Row label="④ 재질">
                <select value={material} onChange={e => setMaterial(e.target.value)} className={selectCls}>
                  <option value="">— 선택 —</option>
                  {rules.material.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                </select>
              </Row>
              <Row label="⑤ 커넥터 A端">
                <select value={connectorA} onChange={e => setConnectorA(e.target.value)} className={selectCls}>
                  <option value="">— 선택 —</option>
                  {rules.connector.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                </select>
              </Row>
              <Row label="⑥ 커넥터 B端">
                <div className="flex items-center gap-2">
                  <select
                    value={connectorB} onChange={e => setConnectorB(e.target.value)}
                    disabled={connBSame} className={`${selectCls} flex-1 ${connBSame ? 'opacity-40' : ''}`}
                  >
                    <option value="">— 선택 —</option>
                    {rules.connector.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                  </select>
                  <label className="flex items-center gap-1 text-xs text-gray-600 whitespace-nowrap cursor-pointer select-none">
                    <input type="checkbox" checked={connBSame} onChange={e => setConnBSame(e.target.checked)} className="w-3.5 h-3.5" />
                    A端과 동일
                  </label>
                </div>
              </Row>
              <Row label="⑦ 길이 (M)">
                <div className="flex items-center gap-2">
                  <select value={lenIdx} onChange={e => setLenIdx(e.target.value)} className={selectCls}>
                    <option value="">— 선택 —</option>
                    {rules.lengthPreset.map((e, i) => <option key={i} value={String(i)}>{e.m}M ({e.code})</option>)}
                    <option value="custom">직접 입력...</option>
                  </select>
                  {lenIdx === 'custom' && (
                    <input
                      type="number" step="0.1" min="0.1" value={lenCustom}
                      onChange={e => setLenCustom(e.target.value)} placeholder="예: 7.5"
                      className="w-24 border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:border-sky-500"
                    />
                  )}
                  {manualLengthCode && (
                    <span className="text-xs text-gray-400 font-mono whitespace-nowrap">→ {manualLengthCode}</span>
                  )}
                </div>
              </Row>
              <Row label="⑧ 마킹">
                <select value={marking} onChange={e => setMarking(e.target.value)} className={selectCls}>
                  <option value="">— 선택 —</option>
                  {rules.marking.map(e => <option key={e.code} value={e.code}>{e.code} — {e.label}</option>)}
                </select>
              </Row>
            </div>
            <div className="mt-6 pt-5 border-t border-gray-200">
              {manualPartNumber ? (
                <div className="flex items-center gap-3">
                  <span className="font-mono text-2xl font-bold text-sky-700 tracking-widest flex-1 break-all">
                    {manualPartNumber}
                  </span>
                  <button
                    onClick={() => handleCopy(manualPartNumber)}
                    className="px-4 py-1.5 text-sm bg-sky-600 text-white rounded hover:bg-sky-700 transition-colors"
                  >
                    {copied ? '✓ 복사됨' : '복사'}
                  </button>
                </div>
              ) : (
                <p className="text-sm text-gray-400">
                  {manualMissing.length > 0 ? `미선택: ${manualMissing.join(' / ')}` : '—'}
                </p>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ── 규칙 편집 탭 ──────────────────────────────────── */}
      {subTab === 'rules' && (
        <div>
          <div className="flex items-center justify-between mb-5">
            <p className="text-sm text-gray-500">
              각 코드표를 직접 수정합니다. 변경사항은 <strong>저장</strong> 버튼을 눌러야 반영됩니다.
            </p>
            <div className="flex gap-2">
              <button onClick={handleReset} className="px-3 py-1.5 text-sm border border-gray-300 rounded hover:bg-gray-50 transition-colors">
                기본값 초기화
              </button>
              <button
                onClick={handleSave}
                disabled={!hasChanges || saveStatus === 'saving'}
                className={`px-4 py-1.5 text-sm rounded font-medium transition-colors ${
                  hasChanges ? 'bg-sky-600 text-white hover:bg-sky-700' : 'bg-gray-100 text-gray-400 cursor-not-allowed'
                }`}
              >
                {saveStatus === 'saving' ? '저장 중...' : saveStatus === 'saved' ? '✓ 저장됨' : saveStatus === 'error' ? '⚠ 오류' : '저장'}
              </button>
            </div>
          </div>
          <div className="space-y-3">
            {(Object.keys(SECTION_LABELS) as Section[]).map(section => (
              <div key={section} className="bg-white border border-gray-200 rounded-lg overflow-hidden">
                <button
                  onClick={() => setOpenSections(o => ({ ...o, [section]: !o[section] }))}
                  className="w-full flex items-center justify-between px-4 py-3 text-sm font-medium text-gray-700 hover:bg-gray-50 transition-colors"
                >
                  <span>{SECTION_LABELS[section]}</span>
                  <span className="text-gray-400 text-xs">
                    {(editRules[section] as unknown as AnyEntry[]).length}개 항목 {openSections[section] ? '▲' : '▼'}
                  </span>
                </button>
                {openSections[section] && (
                  <div className="border-t border-gray-100 p-4">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="text-xs text-gray-400 border-b border-gray-100">
                          {section === 'lengthPreset' ? (
                            <><th className="text-left pb-2 w-32">길이 (M)</th><th className="text-left pb-2">코드 (3자)</th></>
                          ) : (
                            <><th className="text-left pb-2 w-24">코드</th><th className="text-left pb-2">{section === 'coreCount' ? '심선수 (수량)' : '설명'}</th></>
                          )}
                          <th className="w-8"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {(editRules[section] as unknown as AnyEntry[]).map((entry, idx) => (
                          <tr key={idx} className="border-b border-gray-50 last:border-0">
                            {section === 'lengthPreset' ? (
                              <>
                                <td className="py-1 pr-2">
                                  <input type="number" step="0.1" min="0" value={entry.m as number}
                                    onChange={e => updateRule(section, idx, 'm', parseFloat(e.target.value) || 0)} className={inputCls} />
                                </td>
                                <td className="py-1 pr-2">
                                  <input value={String(entry.code)}
                                    onChange={e => updateRule(section, idx, 'code', e.target.value.toUpperCase().slice(0, 3))}
                                    className={inputCls} maxLength={3} />
                                </td>
                              </>
                            ) : (
                              <>
                                <td className="py-1 pr-2">
                                  <input value={String(entry.code)}
                                    onChange={e => updateRule(section, idx, 'code', e.target.value.toUpperCase().slice(0, 2))}
                                    className={inputCls} maxLength={2} />
                                </td>
                                <td className="py-1 pr-2">
                                  {section === 'coreCount' ? (
                                    <input type="number" min="1" value={entry.count as number}
                                      onChange={e => updateRule(section, idx, 'count', parseInt(e.target.value) || 0)} className={inputCls} />
                                  ) : (
                                    <input value={String(entry.label)}
                                      onChange={e => updateRule(section, idx, 'label', e.target.value)} className={inputCls} />
                                  )}
                                </td>
                              </>
                            )}
                            <td className="py-1 text-center">
                              <button onClick={() => deleteRow(section, idx)} className="text-red-300 hover:text-red-500 px-1 text-sm" title="삭제">✕</button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <button onClick={() => addRow(section)} className="mt-2 text-xs text-sky-600 hover:text-sky-800 font-medium">
                      + 행 추가
                    </button>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}

function Row({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex items-center gap-3">
      <span className="text-sm font-medium text-gray-700 w-32 flex-shrink-0">{label}</span>
      <div className="flex-1">{children}</div>
    </div>
  )
}
