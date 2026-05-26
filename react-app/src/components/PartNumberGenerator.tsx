import { useState, useEffect } from 'react'
import type { OjcRules } from '../lib/types'
import { DEFAULT_OJC_RULES } from '../lib/types'
import { loadOjcRules, loadItemCodes } from '../lib/supabase'
import type { OjcProduct, ItemCode } from '../lib/supabase'
import type { DetectedFields } from '../lib/ojcAutoDetect'
import { analyzeProduct } from '../lib/ojcAutoDetect'
import EmpCodeGenerator from './EmpCodeGenerator'

type SubTab = 'auto' | 'manual' | 'emp'

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

interface Props {
  ojcProducts:    OjcProduct[]
  setOjcProducts: (p: OjcProduct[]) => void
}

export default function PartNumberGenerator({ ojcProducts, setOjcProducts }: Props) {
  const [subTab, setSubTab]         = useState<SubTab>('auto')
  const [rules, setRules]           = useState<OjcRules>(DEFAULT_OJC_RULES)
  const [copied, setCopied]         = useState(false)
  const [savedPN, setSavedPN]       = useState(false)
  const [itemCodes, setItemCodes] = useState<ItemCode[]>([])

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
    loadOjcRules().then(setRules)
    loadItemCodes().then(setItemCodes)
  }, [])

  // ── 기존 품번 조회 ─────────────────────────────────────────
  const existingProduct = ojcProducts.find(
    p => p.name.trim() === autoName.trim() && p.spec.trim() === autoSpec.trim()
  )

  function ecountCheck(pn: string | null): ItemCode | null {
    if (!pn) return null
    return itemCodes.find(ic => ic.code === pn) ?? null
  }

  function saveCurrentPartNumber(pn: string) {
    if (!autoName.trim()) return
    const next: OjcProduct = {
      name:       autoName.trim(),
      spec:       autoSpec.trim(),
      partNumber: pn,
      createdAt:  new Date().toISOString().slice(0, 10),
    }
    const filtered = ojcProducts.filter(
      p => !(p.name === next.name && p.spec === next.spec)
    )
    setOjcProducts([...filtered, next])
    setSavedPN(true)
    setTimeout(() => setSavedPN(false), 2000)
  }

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

  // ── 렌더 ──────────────────────────────────────────────────
  return (
    <div>
      {/* 서브탭 */}
      <div className="flex border-b mb-6">
        {([
          ['auto',   'OJC 품번 — 자동 분석'],
          ['manual', 'OJC 품번 — 수동 선택'],
          ['emp',    'EMP 코드 생성기'],
        ] as [SubTab, string][]).map(([id, lbl]) => (
          <button
            key={id}
            onClick={() => setSubTab(id)}
            className={`px-5 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${
              subTab === id
                ? 'border-sky-500 text-sky-600'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            {lbl}
          </button>
        ))}
      </div>

      {/* ── 자동 분석 탭 ──────────────────────────────────── */}
      {subTab === 'auto' && (
        <div className="max-w-2xl space-y-5">
          {/* 입력 영역 */}
          <div className="bg-white rounded-lg border border-gray-200 p-5">
            <h2 className="text-sm font-semibold text-gray-700 mb-4">
              품목명과 규격을 입력하면 품번이 자동 생성됩니다.
              동일한 품목명·규격으로 이미 저장된 품번이 있으면 먼저 표시됩니다.
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

          {/* 기존 품번 조회 결과 */}
          {existingProduct && autoName.trim() && (
            <div className="bg-sky-50 border border-sky-200 rounded-lg px-5 py-3 flex items-center justify-between gap-3">
              <div>
                <div className="text-xs text-sky-600 font-semibold mb-0.5">기존 등록 품번</div>
                <div className="font-mono text-lg font-bold text-sky-800 tracking-widest">
                  {existingProduct.partNumber}
                </div>
                <div className="text-xs text-sky-500 mt-0.5">
                  {existingProduct.name} / {existingProduct.spec} — {existingProduct.createdAt} 저장
                </div>
              </div>
              <button
                onClick={() => handleCopy(existingProduct.partNumber)}
                className="flex-shrink-0 px-4 py-1.5 text-sm bg-sky-600 text-white rounded hover:bg-sky-700 transition-colors"
              >
                {copied ? '✓ 복사됨' : '복사'}
              </button>
            </div>
          )}

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
                  <div className="space-y-2">
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
                    <EcountBadge match={ecountCheck(autoPartNumber)} hasItemCodes={itemCodes.length > 0} />
                    {autoName.trim() && (
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => saveCurrentPartNumber(autoPartNumber)}
                          className="text-xs px-3 py-1 border border-sky-300 text-sky-600 rounded hover:bg-sky-50 transition-colors"
                        >
                          {savedPN ? '✓ 저장됨' : '💾 품번 저장 (조회용)'}
                        </button>
                        {existingProduct?.partNumber === autoPartNumber && (
                          <span className="text-xs text-gray-400">이미 저장된 품번과 동일</span>
                        )}
                      </div>
                    )}
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
            <div className="mt-6 pt-5 border-t border-gray-200 space-y-2">
              {manualPartNumber ? (
                <>
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
                  <EcountBadge match={ecountCheck(manualPartNumber)} hasItemCodes={itemCodes.length > 0} />
                </>
              ) : (
                <p className="text-sm text-gray-400">
                  {manualMissing.length > 0 ? `미선택: ${manualMissing.join(' / ')}` : '—'}
                </p>
              )}
            </div>
          </div>
        </div>
      )}

      {subTab === 'emp' && <EmpCodeGenerator />}
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

function EcountBadge({ match, hasItemCodes }: { match: ItemCode | null; hasItemCodes: boolean }) {
  if (!hasItemCodes) return (
    <span className="inline-flex items-center text-xs text-gray-400 gap-1">
      ⚪ 이카운트 품목 리스트 미등록 — 수입 관리 탭에서 파일을 업로드하면 중복 확인이 활성화됩니다
    </span>
  )
  if (match) return (
    <span className="inline-flex items-center gap-1.5 text-xs px-2 py-0.5 rounded-full bg-red-50 text-red-700 font-medium">
      ⚠ 이카운트 등록됨 — {match.name || match.code} (신규 등록 전 확인 필요)
    </span>
  )
  return (
    <span className="inline-flex items-center gap-1.5 text-xs px-2 py-0.5 rounded-full bg-green-50 text-green-700 font-medium">
      ✅ 이카운트 미등록 — 신규 품번
    </span>
  )
}
