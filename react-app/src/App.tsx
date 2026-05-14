import { useState, useEffect } from 'react'
import {
  loadSettings, loadMetadata, loadInventory, loadSalesAnalysis, loadSalesAgg
} from './lib/supabase'
import { DEFAULT_SETTINGS } from './lib/types'
import type { AppSettings, Metadata, Inventory, SalesAnalysis } from './lib/types'
import type { SalesAggResult } from './lib/aggregate/salesAgg'
import Step1 from './components/Step1'
import Step2 from './components/Step2'
import Step3 from './components/Step3'
import MetaManager from './components/MetaManager'
import InventoryComp from './components/Inventory'
import Settings from './components/Settings'
import PartNumberGenerator from './components/PartNumberGenerator'

const TABS = [
  { id: 'step1',     label: '📤 STEP 1 — ERP 파일 가공' },
  { id: 'step2',     label: '📈 STEP 2 — 판매 분석' },
  { id: 'step3',     label: '📊 STEP 3 — 발주계획 생성' },
  { id: 'partnum',   label: '🏷 품번 생성기' },
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
  const [sales, setSales]         = useState<SalesAnalysis>({})
  const [salesAgg, setSalesAgg]   = useState<SalesAggResult | null>(null)

  useEffect(() => {
    Promise.all([
      loadSettings(),
      loadMetadata(),
      loadInventory(),
      loadSalesAnalysis(),
      loadSalesAgg(),
    ]).then(([s, m, inv, sa, sagg]) => {
      setSettings(s)
      setMetadata(m)
      setInventory(inv)
      setSales(sa)
      setSalesAgg(sagg)
      setLoading(false)
    })
  }, [])

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white">
        <div className="text-gray-500">⏳ 데이터 로드 중...</div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-white">
      {/* Title — matches Streamlit .main-title / .sub-title */}
      <div className="px-6 pt-6 pb-1 max-w-screen-2xl mx-auto">
        <div className="text-[1.6rem] font-bold text-[#1F3864] leading-tight mb-1">
          📦 AJW 생산자재 발주계획 시스템
        </div>
        <div className="text-[0.95rem] text-gray-500 mb-4">
          (주)에이제이월드 SCM팀 — 로우데이터 업로드 후 버튼 클릭으로 Excel 자동 생성
        </div>
      </div>

      {/* Tabs — matches Streamlit st.tabs */}
      <div className="border-b border-gray-200 sticky top-0 bg-white z-10">
        <div className="flex overflow-x-auto max-w-screen-2xl mx-auto px-4">
          {TABS.map(t => (
            <button
              key={t.id}
              onClick={() => setActiveTab(t.id)}
              className={`px-4 py-2.5 text-sm font-medium whitespace-nowrap border-b-2 transition-colors flex-shrink-0 ${
                activeTab === t.id
                  ? 'border-[#FF4B4B] text-[#FF4B4B]'
                  : 'border-transparent text-gray-600 hover:text-gray-900 hover:border-gray-300'
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>

      {/* Content — wide, no card wrapper */}
      <div className="max-w-screen-2xl mx-auto px-6 py-6">
        {activeTab === 'step1' && (
          <Step1 metadata={metadata} setMetadata={setMetadata} settings={settings} />
        )}
        {activeTab === 'step2' && (
          <Step2 salesAgg={salesAgg} setSalesAgg={setSalesAgg} />
        )}
        {activeTab === 'step3' && (
          <Step3 metadata={metadata} inventory={inventory} settings={settings} />
        )}
        {activeTab === 'partnum' && (
          <PartNumberGenerator />
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
