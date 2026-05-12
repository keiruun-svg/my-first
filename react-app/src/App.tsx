import { useState, useEffect } from 'react'
import {
  loadSettings, loadMetadata, loadInventory, loadSalesAnalysis
} from './lib/supabase'
import { DEFAULT_SETTINGS } from './lib/types'
import type { AppSettings, Metadata, Inventory, SalesAnalysis } from './lib/types'
import Step1 from './components/Step1'
import Step2 from './components/Step2'
import Step3 from './components/Step3'
import MetaManager from './components/MetaManager'
import InventoryComp from './components/Inventory'
import Settings from './components/Settings'

const TABS = [
  { id: 'step1', label: '📤 STEP 1', title: 'ERP 파일 파싱' },
  { id: 'step2', label: '📈 STEP 2', title: '판매 분석' },
  { id: 'step3', label: '📊 STEP 3', title: '발주계획 생성' },
  { id: 'meta', label: '📋 품번 관리', title: '품번 관리' },
  { id: 'inventory', label: '📦 재고 현황', title: '재고 현황' },
  { id: 'settings', label: '⚙️ 설정', title: '파라미터 & 설정' },
] as const

type TabId = typeof TABS[number]['id']

export default function App() {
  const [activeTab, setActiveTab] = useState<TabId>('step1')
  const [loading, setLoading] = useState(true)
  const [settings, setSettings] = useState<AppSettings>(DEFAULT_SETTINGS)
  const [metadata, setMetadata] = useState<Metadata>({ cable: {}, housing: {} })
  const [inventory, setInventory] = useState<Inventory>({ cable: {}, housing: {} })
  const [sales, setSales] = useState<SalesAnalysis>({})

  useEffect(() => {
    Promise.all([
      loadSettings(),
      loadMetadata(),
      loadInventory(),
      loadSalesAnalysis(),
    ]).then(([s, m, inv, sa]) => {
      setSettings(s)
      setMetadata(m)
      setInventory(inv)
      setSales(sa)
      setLoading(false)
    })
  }, [])

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-100">
        <div className="text-gray-500 text-lg">⏳ 데이터 로드 중...</div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-100">
      {/* Header */}
      <div className="bg-[#1F3864] text-white px-6 py-4 shadow-lg">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div>
            <h1 className="text-xl font-bold">AJW 생산자재 발주계획 시스템</h1>
            <p className="text-blue-200 text-xs mt-0.5">AJWorld SCM팀 — 연간 발주계획 자동화</p>
          </div>
          <div className="text-right text-xs text-blue-300">
            <div>케이블 {Object.keys(metadata.cable).length}종</div>
            <div>하우징 {Object.keys(metadata.housing).length}종</div>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div className="bg-white border-b shadow-sm sticky top-0 z-10">
        <div className="max-w-7xl mx-auto flex overflow-x-auto">
          {TABS.map(t => (
            <button
              key={t.id}
              onClick={() => setActiveTab(t.id)}
              className={`px-4 py-3 text-sm font-semibold whitespace-nowrap border-b-2 transition-colors ${
                activeTab === t.id
                  ? 'border-blue-600 text-blue-600 bg-blue-50'
                  : 'border-transparent text-gray-600 hover:text-blue-500 hover:bg-gray-50'
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>

      {/* Content */}
      <div className="max-w-7xl mx-auto px-4 py-6">
        <div className="bg-white rounded-xl shadow p-6">
          {activeTab === 'step1' && (
            <Step1 metadata={metadata} setMetadata={setMetadata} settings={settings} />
          )}
          {activeTab === 'step2' && (
            <Step2 sales={sales} setSales={setSales} />
          )}
          {activeTab === 'step3' && (
            <Step3 metadata={metadata} inventory={inventory} sales={sales} settings={settings} />
          )}
          {activeTab === 'meta' && (
            <MetaManager metadata={metadata} setMetadata={setMetadata} />
          )}
          {activeTab === 'inventory' && (
            <InventoryComp inventory={inventory} setInventory={setInventory} metadata={metadata} />
          )}
          {activeTab === 'settings' && (
            <Settings
              settings={settings} setSettings={setSettings}
              metadata={metadata} inventory={inventory} sales={sales}
            />
          )}
        </div>
      </div>
    </div>
  )
}
