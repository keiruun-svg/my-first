import { useState } from 'react'
import { saveInventory } from '../lib/supabase'
import type { Inventory as Inv, Metadata } from '../lib/types'

interface Props {
  inventory: Inv
  setInventory: (i: Inv) => void
  metadata: Metadata
}

const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'
const thL    = `${thBase} text-left`
const thR    = `${thBase} text-right`
const tdL    = 'px-3 py-1.5 text-sm text-gray-900'
const tdMono = 'px-3 py-1.5 text-xs font-mono text-gray-500'

export default function Inventory({ inventory, setInventory, metadata }: Props) {
  const [tab, setTab] = useState<'cable' | 'housing'>('cable')
  const [saved, setSaved] = useState(false)

  const updateCable = (key: string, field: '현재고' | '기발주', value: string) => {
    const prev = inventory.cable[key] ?? { 현재고: 0, 기발주: 0 }
    setInventory({
      ...inventory,
      cable: { ...inventory.cable, [key]: { ...prev, [field]: parseInt(value) || 0 } },
    })
  }

  const updateHousing = (key: string, field: '현재고' | '기발주', value: string) => {
    const cur  = inventory.housing[key]
    const prev = Array.isArray(cur) ? cur[0] : (cur ?? { 현재고: 0, 기발주: 0 })
    setInventory({
      ...inventory,
      housing: { ...inventory.housing, [key]: { ...prev, [field]: parseInt(value) || 0 } },
    })
  }

  const save = () => {
    saveInventory(inventory)
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  const cableKeys  = Object.keys(metadata.cable).sort()
  const housingKeys = Object.keys(metadata.housing).sort()

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>재고 현황</b>: 현재고 및 기발주 수량을 입력하세요. STEP 3 발주계획 생성 시 자동 반영됩니다.
      </div>

      {/* 탭 — 활성: 노란색 */}
      <div className="flex gap-2">
        {(['cable', 'housing'] as const).map(t => (
          <button
            key={t}
            onClick={() => setTab(t)}
            className={`px-5 py-2 rounded font-semibold text-sm border transition ${
              tab === t
                ? 'bg-yellow-400 border-yellow-500 text-black'
                : 'bg-white border-gray-300 text-gray-600 hover:bg-gray-50'
            }`}
          >
            {t === 'cable' ? `케이블 (${cableKeys.length})` : `하우징 (${housingKeys.length})`}
          </button>
        ))}
      </div>

      {/* 케이블 */}
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
                      <input type="number" min="0"
                        value={inv.현재고}
                        onChange={e => updateCable(key, '현재고', e.target.value)}
                        className="w-28 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400"
                      />
                    </td>
                    <td className="px-2 py-1 text-right">
                      <input type="number" min="0"
                        value={inv.기발주 ?? 0}
                        onChange={e => updateCable(key, '기발주', e.target.value)}
                        className="w-28 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400"
                      />
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

      {/* 하우징 */}
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
                const mRaw  = metadata.housing[key]
                const m     = Array.isArray(mRaw) ? mRaw[0] : mRaw
                const invRaw = inventory.housing[key]
                const inv   = (Array.isArray(invRaw) ? invRaw[0] : invRaw) ?? { 현재고: 0, 기발주: 0 }
                return (
                  <tr key={key} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className={tdMono}>{pai}</td>
                    <td className={`${tdL} font-semibold`}>{ht}</td>
                    <td className={tdMono}>{m?.품번 ?? ''}</td>
                    <td className={`${tdL} max-w-xs truncate`} title={m?.품명 ?? ''}>{m?.품명 ?? ''}</td>
                    <td className="px-2 py-1 text-right">
                      <input type="number" min="0"
                        value={inv.현재고}
                        onChange={e => updateHousing(key, '현재고', e.target.value)}
                        className="w-28 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400"
                      />
                    </td>
                    <td className="px-2 py-1 text-right">
                      <input type="number" min="0"
                        value={inv.기발주}
                        onChange={e => updateHousing(key, '기발주', e.target.value)}
                        className="w-28 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400"
                      />
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

      <div className="flex items-center gap-3">
        <button onClick={save}
          className="bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold px-6 py-2 rounded-lg transition">
          💾 저장
        </button>
        {saved && <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>}
      </div>
    </div>
  )
}
