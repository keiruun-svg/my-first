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
  { id: 'step1',     label: '📤 STEP 1 — ERP 파일 가공' },
  { id: 'step2',     label: '📈 STEP 2 — 판매 분석' },
  { id: 'step3',     label: '📊 STEP 3 — 발주계획 생성' },
  { id: 'meta',      label: '📋 품번 관리' },
  { id: 'inventory', label: '📦 재고 현황' },
  { id: 'settings',  label: '⚙️ 파라미터 & 양식 설정' },
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
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="text-gray-500">⏳ 데이터 로드 중...</div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <div className="bg-[#1F3864] text-white px-6 py-3 shadow">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold tracking-tight">📦 AJW 생산자재 발주계획 시스템</h1>
            <p className="text-blue-200 text-sm mt-0.5">(주)에이제이월드 SCM팀 — 로우데이터 업로드 후 버튼 클릭으로 Excel 자동 생성</p>
          </div>
          <div className="text-right text-xs text-blue-300 hidden md:block">
            <div>케이블 {Object.keys(metadata.cable).length}종</div>
            <div>하우징 {Object.keys(metadata.housing).length}종</div>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div className="bg-white border-b sticky top-0 z-10 shadow-sm">
        <div className="flex overflow-x-auto px-2">
          {TABS.map(t => (
            <button
              key={t.id}
              onClick={() => setActiveTab(t.id)}
              className={`px-4 py-3 text-sm font-medium whitespace-nowrap border-b-2 transition-colors flex-shrink-0 ${
                activeTab === t.id
                  ? 'border-[#E63946] text-[#E63946]'
                  : 'border-transparent text-gray-600 hover:text-gray-900 hover:border-gray-300'
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>

      {/* Content — wide, no card wrapper */}
      <div className="px-6 py-6 max-w-screen-xl mx-auto">
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
  )
}
