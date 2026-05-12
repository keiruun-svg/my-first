import { useState } from 'react'
import { saveMetadata } from '../lib/supabase'
import type { Metadata, CableMeta, HousingComp } from '../lib/types'

interface Props {
  metadata: Metadata
  setMetadata: (m: Metadata) => void
}

export default function MetaManager({ metadata, setMetadata }: Props) {
  const [tab, setTab] = useState<'cable' | 'housing'>('cable')
  const [saved, setSaved] = useState(false)

  const updateCable = (key: string, field: keyof CableMeta, value: string) => {
    const newMeta: Metadata = {
      ...metadata,
      cable: {
        ...metadata.cable,
        [key]: { ...(metadata.cable[key] || { 품번:'', 품명:'', 구매처:'', 리드타임:null }), [field]: value }
      }
    }
    setMetadata(newMeta)
  }

  const updateHousing = (key: string, field: keyof HousingComp, value: string) => {
    const cur = metadata.housing[key]
    const comp = Array.isArray(cur) ? cur[0] : (cur || { 품번:'', 품명:'', 구매처:'', 리드타임:null })
    const updated = { ...comp, [field]: value }
    const newMeta: Metadata = {
      ...metadata,
      housing: { ...metadata.housing, [key]: updated }
    }
    setMetadata(newMeta)
  }

  const save = () => {
    saveMetadata(metadata)
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  const cableKeys = Object.keys(metadata.cable).sort()
  const housingKeys = Object.keys(metadata.housing).sort()

  const getCableMeta = (key: string): CableMeta =>
    metadata.cable[key] || { 품번:'', 품명:'', 구매처:'', 리드타임:null }

  const getHousingMeta = (key: string): HousingComp => {
    const cur = metadata.housing[key]
    return (Array.isArray(cur) ? cur[0] : cur) || { 품번:'', 품명:'', 구매처:'', 리드타임:null }
  }

  return (
    <div className="space-y-4">
      <div className="bg-orange-50 border border-orange-200 rounded p-4">
        <h3 className="font-bold text-orange-800 mb-1">품번 관리</h3>
        <p className="text-sm text-orange-700">STEP 1 실행 후 감지된 케이블·하우징 타입의 품번·품명·구매처·리드타임을 입력하세요.</p>
      </div>

      <div className="flex gap-2">
        <button onClick={() => setTab('cable')}
          className={`px-4 py-2 rounded font-semibold text-sm transition ${tab==='cable' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>
          케이블 ({cableKeys.length})
        </button>
        <button onClick={() => setTab('housing')}
          className={`px-4 py-2 rounded font-semibold text-sm transition ${tab==='housing' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>
          하우징 ({housingKeys.length})
        </button>
      </div>

      {tab === 'cable' && (
        <div className="overflow-x-auto">
          <table className="text-xs w-full border-collapse">
            <thead className="bg-gray-700 text-white">
              <tr>
                {['파이','케이블종류','품번','품명','구매처','리드타임(일)'].map(h => (
                  <th key={h} className="px-3 py-2 text-left font-semibold">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cableKeys.map((key, i) => {
                const [pai, ct] = key.split('|')
                const m = getCableMeta(key)
                return (
                  <tr key={key} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="px-3 py-1 font-mono text-gray-600">{pai}</td>
                    <td className="px-3 py-1 font-semibold">{ct}</td>
                    {(['품번','품명','구매처','리드타임'] as const).map(f => (
                      <td key={f} className="px-1 py-1">
                        <input
                          value={String(m[f] ?? '')}
                          onChange={e => updateCable(key, f, e.target.value)}
                          className="w-full border rounded px-2 py-0.5 text-xs focus:outline-blue-400"
                          placeholder={f}
                        />
                      </td>
                    ))}
                  </tr>
                )
              })}
            </tbody>
          </table>
          {cableKeys.length === 0 && (
            <div className="text-center text-gray-400 py-8">STEP 1을 실행하면 케이블 목록이 나타납니다.</div>
          )}
        </div>
      )}

      {tab === 'housing' && (
        <div className="overflow-x-auto">
          <table className="text-xs w-full border-collapse">
            <thead className="bg-gray-700 text-white">
              <tr>
                {['파이','하우징타입','품번','품명','구매처','리드타임(일)'].map(h => (
                  <th key={h} className="px-3 py-2 text-left font-semibold">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {housingKeys.map((key, i) => {
                const [pai, ht] = key.split('|')
                const m = getHousingMeta(key)
                return (
                  <tr key={key} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="px-3 py-1 font-mono text-gray-600">{pai}</td>
                    <td className="px-3 py-1 font-semibold">{ht}</td>
                    {(['품번','품명','구매처','리드타임'] as const).map(f => (
                      <td key={f} className="px-1 py-1">
                        <input
                          value={String(m[f] ?? '')}
                          onChange={e => updateHousing(key, f, e.target.value)}
                          className="w-full border rounded px-2 py-0.5 text-xs focus:outline-blue-400"
                          placeholder={f}
                        />
                      </td>
                    ))}
                  </tr>
                )
              })}
            </tbody>
          </table>
          {housingKeys.length === 0 && (
            <div className="text-center text-gray-400 py-8">STEP 1을 실행하면 하우징 목록이 나타납니다.</div>
          )}
        </div>
      )}

      <div className="flex items-center gap-3">
        <button onClick={save}
          className="bg-blue-600 hover:bg-blue-700 text-white font-bold px-6 py-2 rounded-lg transition">
          💾 저장
        </button>
        {saved && <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>}
      </div>
    </div>
  )
}
