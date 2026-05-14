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

  const getCableMeta = (key: string): CableMeta =>
    metadata.cable[key] || { 품번:'', 품명:'', 구매처:'', 리드타임:null }

  const getHousingMeta = (key: string): HousingComp => {
    const cur = metadata.housing[key]
    return (Array.isArray(cur) ? cur[0] : cur) || { 품번:'', 품명:'', 구매처:'', 리드타임:null }
  }

  const cableKeys = Object.keys(metadata.cable).sort()
  const housingKeys = Object.keys(metadata.housing).sort()
  const cableMissing = cableKeys.filter(k => !getCableMeta(k).품번).length
  const housingMissing = housingKeys.filter(k => !getHousingMeta(k).품번).length

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>품번 관리</b>: STEP 1 실행 후 발견된 신규 타입이 자동으로 추가됩니다.
        품번·품명·구매처·리드타임을 직접 입력하고 <b>저장</b> 버튼을 누르세요.
        다음 STEP 1/3 실행 시 자동으로 적용됩니다.
      </div>

      <div className="flex gap-2">
        <button onClick={() => setTab('cable')}
          className={`px-4 py-2 rounded font-semibold text-sm transition flex items-center gap-2 ${tab==='cable' ? 'bg-[#2E75B6] text-white' : 'bg-gray-100 text-gray-700 hover:bg-gray-200'}`}>
          케이블
          <span className={`text-xs px-1.5 py-0.5 rounded-full font-bold ${tab==='cable' ? 'bg-white/20' : 'bg-gray-300'}`}>{cableKeys.length}</span>
          {cableMissing > 0 && <span className="text-xs bg-red-500 text-white px-1.5 py-0.5 rounded-full">⚠ {cableMissing}</span>}
        </button>
        <button onClick={() => setTab('housing')}
          className={`px-4 py-2 rounded font-semibold text-sm transition flex items-center gap-2 ${tab==='housing' ? 'bg-[#2E75B6] text-white' : 'bg-gray-100 text-gray-700 hover:bg-gray-200'}`}>
          하우징
          <span className={`text-xs px-1.5 py-0.5 rounded-full font-bold ${tab==='housing' ? 'bg-white/20' : 'bg-gray-300'}`}>{housingKeys.length}</span>
          {housingMissing > 0 && <span className="text-xs bg-red-500 text-white px-1.5 py-0.5 rounded-full">⚠ {housingMissing}</span>}
        </button>
      </div>

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
