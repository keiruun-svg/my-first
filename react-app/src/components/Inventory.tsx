import { useState } from 'react'
import { saveInventory } from '../lib/supabase'
import type { Inventory as Inv, Metadata } from '../lib/types'

interface Props {
  inventory: Inv
  setInventory: (i: Inv) => void
  metadata: Metadata
}

export default function Inventory({ inventory, setInventory, metadata }: Props) {
  const [tab, setTab] = useState<'cable' | 'housing'>('cable')
  const [saved, setSaved] = useState(false)

  const updateCable = (key: string, value: string) => {
    setInventory({
      ...inventory,
      cable: { ...inventory.cable, [key]: { 현재고: parseInt(value) || 0 } }
    })
  }

  const updateHousing = (key: string, field: '현재고' | '기발주', value: string) => {
    const cur = inventory.housing[key]
    const prev = Array.isArray(cur)
      ? cur[0]
      : (cur ?? { 현재고: 0, 기발주: 0 })
    setInventory({
      ...inventory,
      housing: { ...inventory.housing, [key]: { ...prev, [field]: parseInt(value) || 0 } }
    })
  }

  const save = () => {
    saveInventory(inventory)
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  const cableKeys = Object.keys(metadata.cable).sort()
  const housingKeys = Object.keys(metadata.housing).sort()

  return (
    <div className="space-y-4">
      <div className="bg-teal-50 border border-teal-200 rounded p-4">
        <h3 className="font-bold text-teal-800 mb-1">재고 현황</h3>
        <p className="text-sm text-teal-700">현재고 및 기발주 수량을 입력하세요. STEP 3 발주계획 생성 시 자동 반영됩니다.</p>
      </div>

      <div className="flex gap-2">
        <button onClick={() => setTab('cable')}
          className={`px-4 py-2 rounded font-semibold text-sm transition ${tab==='cable' ? 'bg-teal-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>
          케이블 ({cableKeys.length})
        </button>
        <button onClick={() => setTab('housing')}
          className={`px-4 py-2 rounded font-semibold text-sm transition ${tab==='housing' ? 'bg-teal-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>
          하우징 ({housingKeys.length})
        </button>
      </div>

      {tab === 'cable' && (
        <div className="overflow-x-auto">
          <table className="text-sm w-full border-collapse">
            <thead className="bg-teal-700 text-white">
              <tr>
                {['파이','케이블종류','품번','품명','현재고(m)'].map(h => (
                  <th key={h} className="px-3 py-2 text-left font-semibold">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cableKeys.map((key, i) => {
                const [pai, ct] = key.split('|')
                const m = metadata.cable[key]
                const inv = inventory.cable[key] ?? { 현재고: 0 }
                return (
                  <tr key={key} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="px-3 py-1 text-xs font-mono text-gray-500">{pai}</td>
                    <td className="px-3 py-1 font-semibold text-sm">{ct}</td>
                    <td className="px-3 py-1 text-xs text-gray-500">{m?.품번 ?? ''}</td>
                    <td className="px-3 py-1 text-xs max-w-xs truncate">{m?.품명 ?? ''}</td>
                    <td className="px-2 py-1">
                      <input type="number" min="0"
                        value={inv.현재고}
                        onChange={e => updateCable(key, e.target.value)}
                        className="w-28 border rounded px-2 py-0.5 text-right text-sm focus:outline-teal-400"
                      />
                    </td>
                  </tr>
                )
              })}
              {cableKeys.length === 0 && (
                <tr><td colSpan={5} className="text-center text-gray-400 py-8">품번 관리 탭에서 케이블을 먼저 등록하세요.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {tab === 'housing' && (
        <div className="overflow-x-auto">
          <table className="text-sm w-full border-collapse">
            <thead className="bg-teal-700 text-white">
              <tr>
                {['파이','하우징타입','품번','품명','현재고(EA)','기발주(EA)'].map(h => (
                  <th key={h} className="px-3 py-2 text-left font-semibold">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {housingKeys.map((key, i) => {
                const [pai, ht] = key.split('|')
                const mRaw = metadata.housing[key]
                const m = Array.isArray(mRaw) ? mRaw[0] : mRaw
                const invRaw = inventory.housing[key]
                const inv = (Array.isArray(invRaw) ? invRaw[0] : invRaw) ?? { 현재고: 0, 기발주: 0 }
                return (
                  <tr key={key} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="px-3 py-1 text-xs font-mono text-gray-500">{pai}</td>
                    <td className="px-3 py-1 font-semibold text-sm">{ht}</td>
                    <td className="px-3 py-1 text-xs text-gray-500">{m?.품번 ?? ''}</td>
                    <td className="px-3 py-1 text-xs max-w-xs truncate">{m?.품명 ?? ''}</td>
                    <td className="px-2 py-1">
                      <input type="number" min="0"
                        value={inv.현재고}
                        onChange={e => updateHousing(key, '현재고', e.target.value)}
                        className="w-24 border rounded px-2 py-0.5 text-right text-sm focus:outline-teal-400"
                      />
                    </td>
                    <td className="px-2 py-1">
                      <input type="number" min="0"
                        value={inv.기발주}
                        onChange={e => updateHousing(key, '기발주', e.target.value)}
                        className="w-24 border rounded px-2 py-0.5 text-right text-sm focus:outline-teal-400"
                      />
                    </td>
                  </tr>
                )
              })}
              {housingKeys.length === 0 && (
                <tr><td colSpan={6} className="text-center text-gray-400 py-8">품번 관리 탭에서 하우징을 먼저 등록하세요.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      <div className="flex items-center gap-3">
        <button onClick={save}
          className="bg-teal-600 hover:bg-teal-700 text-white font-bold px-6 py-2 rounded-lg transition">
          💾 저장
        </button>
        {saved && <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>}
      </div>
    </div>
  )
}
