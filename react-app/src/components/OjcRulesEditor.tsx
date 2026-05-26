import { useState, useEffect } from 'react'
import type { OjcRules } from '../lib/types'
import { DEFAULT_OJC_RULES } from '../lib/types'
import { loadOjcRules, saveOjcRules } from '../lib/supabase'

type Section = keyof OjcRules
type AnyEntry = Record<string, string | number>

const SECTION_LABELS: Record<Section, string> = {
  ojcType:      '① OJC 종류',
  coreType:     '② 코어 종류',
  coreCount:    '③ 심선수',
  material:     '④ 재질',
  connector:    '⑤⑥ 커넥터',
  marking:      '⑧ 마킹',
  lengthPreset: '⑦ 길이 코드표',
}

const inputCls = 'w-full border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:border-sky-500'

export default function OjcRulesEditor() {
  const [editRules,    setEditRules]    = useState<OjcRules>(DEFAULT_OJC_RULES)
  const [hasChanges,   setHasChanges]   = useState(false)
  const [saveStatus,   setSaveStatus]   = useState<'idle' | 'saving' | 'saved' | 'error'>('idle')
  const [openSections, setOpenSections] = useState<Partial<Record<Section, boolean>>>({ ojcType: true })

  useEffect(() => { loadOjcRules().then(setEditRules) }, [])

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
      if (section === 'coreCount')         row = { code: '', count: 0 }
      else if (section === 'lengthPreset') row = { m: 0, code: '' }
      else                                 row = { code: '', label: '' }
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

  return (
    <div>
      <div className="flex items-center justify-between mb-5">
        <p className="text-sm text-gray-500">
          각 코드표를 직접 수정합니다. 변경사항은 <strong>저장</strong> 버튼을 눌러야 반영됩니다.
        </p>
        <div className="flex gap-2">
          <button
            onClick={handleReset}
            className="px-3 py-1.5 text-sm border border-gray-300 rounded hover:bg-gray-50 transition-colors"
          >
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
                                onChange={e => updateRule(section, idx, 'm', parseFloat(e.target.value) || 0)}
                                className={inputCls} />
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
                                  onChange={e => updateRule(section, idx, 'count', parseInt(e.target.value) || 0)}
                                  className={inputCls} />
                              ) : (
                                <input value={String(entry.label)}
                                  onChange={e => updateRule(section, idx, 'label', e.target.value)}
                                  className={inputCls} />
                              )}
                            </td>
                          </>
                        )}
                        <td className="py-1 text-center">
                          <button
                            onClick={() => deleteRow(section, idx)}
                            className="text-red-300 hover:text-red-500 px-1 text-sm"
                            title="삭제"
                          >✕</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <button
                  onClick={() => addRow(section)}
                  className="mt-2 text-xs text-sky-600 hover:text-sky-800 font-medium"
                >
                  + 행 추가
                </button>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  )
}
