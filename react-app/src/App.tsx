import { useState, useEffect, useRef } from 'react'
import {
  loadSettings, loadMetadata, loadInventory, loadSalesAnalysis, loadSalesAgg,
  loadOjcProducts, saveOjcProducts, CAN_WRITE,
} from './lib/supabase'
import type { OjcProduct } from './lib/supabase'
import { DEFAULT_SETTINGS } from './lib/types'
import type { AppSettings, Metadata, Inventory, SalesAnalysis } from './lib/types'
import type { SalesAggResult } from './lib/aggregate/salesAgg'
import Step1 from './components/Step1'
import Step2 from './components/Step2'
import Step3 from './components/Step3'
import MaterialManager from './components/MaterialManager'
import Settings from './components/Settings'
import PartNumberGenerator from './components/PartNumberGenerator'
import SalesAnalysisTab from './components/SalesAnalysisTab'
import Dashboard from './components/Dashboard'

const ALL_TABS = [
  { id: 'home',      label: '🏠 홈',                       devOnly: false },
  { id: 'step1',     label: '📤 STEP 1 — ERP 파일 가공',  devOnly: false },
  { id: 'step2',     label: '📈 STEP 2 — 판매 분석',      devOnly: false },
  { id: 'step3',     label: '📊 STEP 3 — 발주계획 생성',  devOnly: false },
  { id: 'sales',     label: '🔍 판매 현황 분석',           devOnly: false },
  { id: 'partnum',   label: '🏷 품번 생성기',              devOnly: false },
  { id: 'material',  label: '📦 자재 관리',                devOnly: false },
  { id: 'settings',  label: '⚙️ 파라미터 & 양식 설정',   devOnly: false },
] as const

const TABS = ALL_TABS.filter(t => !t.devOnly || CAN_WRITE)

type TabId = typeof ALL_TABS[number]['id']

const ADMIN_PASS = import.meta.env.VITE_ADMIN_PASS as string | undefined

export default function App() {
  const [activeTab, setActiveTab] = useState<TabId>('home')
  const [loading, setLoading] = useState(true)
  const [settings, setSettings] = useState<AppSettings>(DEFAULT_SETTINGS)
  const [metadata, setMetadata] = useState<Metadata>({ cable: {}, housing: {}, ferrule: {} })
  const [inventory, setInventory] = useState<Inventory>({ cable: {}, housing: {}, ferrule: {} })
  const [sales, setSales]           = useState<SalesAnalysis>({})
  const [salesAgg, setSalesAgg]     = useState<SalesAggResult | null>(null)
  const [ojcProducts, setOjcProducts] = useState<OjcProduct[]>([])

  // Admin gate — session-level unlock, resets on page reload
  const [materialUnlocked, setMaterialUnlocked] = useState(false)
  const [showAdminGate, setShowAdminGate] = useState(false)
  const [pwInput, setPwInput] = useState('')
  const [pwError, setPwError] = useState(false)
  const pwRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    Promise.all([
      loadSettings(),
      loadMetadata(),
      loadInventory(),
      loadSalesAnalysis(),
      loadSalesAgg(),
      loadOjcProducts(),
    ]).then(([s, m, inv, sa, sagg, ojcp]) => {
      setSettings(s)
      setMetadata(m)
      setInventory(inv)
      setSales(sa)
      setSalesAgg(sagg)
      setOjcProducts(ojcp as OjcProduct[])
      setLoading(false)
    })
  }, [])

  // Focus password input when modal opens
  useEffect(() => {
    if (showAdminGate) {
      setPwInput('')
      setPwError(false)
      setTimeout(() => pwRef.current?.focus(), 50)
    }
  }, [showAdminGate])

  function handleTabClick(tabId: TabId) {
    if (tabId === 'material' && !CAN_WRITE && !materialUnlocked) {
      setShowAdminGate(true)
      return
    }
    setActiveTab(tabId)
  }

  function handleAdminConfirm() {
    if (ADMIN_PASS && pwInput === ADMIN_PASS) {
      setMaterialUnlocked(true)
      setShowAdminGate(false)
      setActiveTab('material')
    } else {
      setPwError(true)
      setPwInput('')
      setTimeout(() => pwRef.current?.focus(), 50)
    }
  }

  function handleAdminCancel() {
    setShowAdminGate(false)
    setPwInput('')
    setPwError(false)
  }

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
              onClick={() => handleTabClick(t.id)}
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
        {activeTab === 'home' && (
          <Dashboard
            metadata={metadata}
            inventory={inventory}
            settings={settings}
            salesAgg={salesAgg}
            onNavigate={(tab) => handleTabClick(tab as TabId)}
          />
        )}
        {activeTab === 'step1' && (
          <Step1 metadata={metadata} setMetadata={setMetadata} settings={settings} />
        )}
        {activeTab === 'step2' && (
          <Step2 salesAgg={salesAgg} setSalesAgg={setSalesAgg} />
        )}
        {activeTab === 'step3' && (
          <Step3 metadata={metadata} inventory={inventory} settings={settings} />
        )}
        {activeTab === 'sales' && (
          <SalesAnalysisTab />
        )}
        {activeTab === 'partnum' && (
          <PartNumberGenerator
            ojcProducts={ojcProducts}
            setOjcProducts={(p) => { setOjcProducts(p); saveOjcProducts(p) }}
          />
        )}
        {activeTab === 'material' && (
          <MaterialManager
            metadata={metadata}     setMetadata={setMetadata}
            inventory={inventory}   setInventory={setInventory}
          />
        )}
        {activeTab === 'settings' && (
          <Settings
            settings={settings} setSettings={setSettings}
            metadata={metadata} inventory={inventory} sales={sales}
            salesAgg={salesAgg}
          />
        )}
      </div>

      {/* Admin gate modal */}
      {showAdminGate && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50"
          onClick={(e) => { if (e.target === e.currentTarget) handleAdminCancel() }}
        >
          <div className="bg-white rounded-2xl shadow-2xl w-80 p-7 flex flex-col gap-4">
            <div className="flex flex-col items-center gap-2">
              <div className="text-4xl">🔒</div>
              <div className="text-lg font-bold text-[#1F3864]">관리자 전용</div>
              <div className="text-sm text-gray-500 text-center">
                자재 관리 탭은 관리자만 접근할 수 있습니다.<br />
                비밀번호를 입력하세요.
              </div>
            </div>
            <input
              ref={pwRef}
              type="password"
              value={pwInput}
              onChange={e => { setPwInput(e.target.value); setPwError(false) }}
              onKeyDown={e => { if (e.key === 'Enter') handleAdminConfirm() }}
              placeholder="비밀번호"
              className={`w-full border rounded-lg px-3 py-2 text-sm outline-none transition-colors ${
                pwError
                  ? 'border-red-400 bg-red-50 focus:border-red-500'
                  : 'border-gray-300 focus:border-[#1F3864]'
              }`}
            />
            {pwError && (
              <div className="text-xs text-red-500 -mt-2 text-center">
                비밀번호가 올바르지 않습니다.
              </div>
            )}
            <div className="flex gap-2">
              <button
                onClick={handleAdminCancel}
                className="flex-1 py-2 rounded-lg border border-gray-300 text-sm text-gray-600 hover:bg-gray-50 transition-colors"
              >
                취소
              </button>
              <button
                onClick={handleAdminConfirm}
                className="flex-1 py-2 rounded-lg bg-[#1F3864] text-white text-sm font-medium hover:bg-[#162a4e] transition-colors"
              >
                확인
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
