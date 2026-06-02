import { useState, useEffect, useRef } from 'react'
import {
  loadSettings, loadMetadata, loadInventory, loadSalesAnalysis, loadSalesAgg,
  loadOjcProducts, saveOjcProducts,
} from './lib/supabase'
import type { OjcProduct } from './lib/supabase'
import { DEFAULT_SETTINGS } from './lib/types'
import type { AppSettings, Metadata, Inventory, SalesAnalysis } from './lib/types'
import type { SalesAggResult } from './lib/aggregate/salesAgg'
import Step1 from './components/Step1'
import Step2 from './components/Step2'
import Step3 from './components/Step3'
import MaterialManager from './components/MaterialManager'
import OjcRulesEditor from './components/OjcRulesEditor'
import Settings from './components/Settings'
import PartNumberGenerator from './components/PartNumberGenerator'
import SalesAnalysisTab from './components/SalesAnalysisTab'
import Dashboard from './components/Dashboard'
import InventoryReconciliationTab from './components/InventoryReconciliationTab'
import ImportTab from './components/ImportTab'

const NAV_TABS = [
  { id: 'home',    label: '🏠 홈' },
  { id: 'steps',   label: '📋 생산자재 발주계획' },
  { id: 'sales',   label: '🔍 판매 현황 분석' },
  { id: 'recon',   label: '🗂 재고 대사' },
  { id: 'import',  label: '🚢 수입 관리' },
  { id: 'partnum', label: '🏷 품번 생성기' },
  { id: 'admin',   label: '🔐 관리자' },
] as const

const STEP_TABS = [
  { id: 'step1', label: 'STEP 1 — ERP 파일 가공' },
  { id: 'step2', label: 'STEP 2 — 판매 분석' },
  { id: 'step3', label: 'STEP 3 — 발주계획 생성' },
] as const

const ADMIN_TABS = [
  { id: 'admin_material', label: '자재 관리' },
  { id: 'admin_ojcrules', label: 'OJC 코드표' },
  { id: 'admin_settings', label: '설정' },
] as const

type StepId    = typeof STEP_TABS[number]['id']
type AdminTabId = typeof ADMIN_TABS[number]['id']
type NavId     = typeof NAV_TABS[number]['id']
type TabId     = Exclude<NavId, 'steps' | 'admin'> | StepId | AdminTabId

function isStepId(id: string): id is StepId {
  return STEP_TABS.some(t => t.id === id)
}
function isAdminTabId(id: string): id is AdminTabId {
  return ADMIN_TABS.some(t => t.id === id)
}

const ADMIN_PASS = import.meta.env.VITE_ADMIN_PASS as string | undefined

export default function App() {
  const [activeTab, setActiveTab] = useState<TabId>('home')
  const [loading, setLoading] = useState(true)
  const [settings, setSettings] = useState<AppSettings>(DEFAULT_SETTINGS)
  const [metadata, setMetadata] = useState<Metadata>({ cable: {}, housing: {}, ferrule: {} })
  const [inventory, setInventory] = useState<Inventory>({ cable: {}, housing: {}, ferrule: {} })
  const [sales, setSales]       = useState<SalesAnalysis>({})
  const [salesAgg, setSalesAgg] = useState<SalesAggResult | null>(null)
  const [ojcProducts, setOjcProducts] = useState<OjcProduct[]>([])

  const [adminUnlocked, setAdminUnlocked] = useState(import.meta.env.DEV as boolean)
  const [showAdminGate, setShowAdminGate] = useState(false)
  const [pwInput, setPwInput] = useState('')
  const [pwError, setPwError] = useState(false)
  const pwRef = useRef<HTMLInputElement>(null)
  const pendingTabAfterUnlock = useRef<TabId | null>(null)

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

  useEffect(() => {
    if (showAdminGate) {
      setPwInput('')
      setPwError(false)
      setTimeout(() => pwRef.current?.focus(), 50)
    }
  }, [showAdminGate])

  // 탭 전환의 유일한 통로 — 게이트 체크가 여기에만 존재
  function navigateTo(tabId: TabId) {
    if (isAdminTabId(tabId) && !adminUnlocked) {
      pendingTabAfterUnlock.current = tabId
      setShowAdminGate(true)
      return
    }
    setActiveTab(tabId)
  }

  // 네비 버튼 전용 — 그룹 ID를 실제 TabId로 해석 후 navigateTo에 위임
  function handleTabClick(navId: NavId) {
    if (navId === 'steps') return navigateTo(isStepId(activeTab) ? activeTab : 'step1')
    if (navId === 'admin') return navigateTo(isAdminTabId(activeTab) ? activeTab : 'admin_material')
    navigateTo(navId as TabId)
  }

  function handleAdminConfirm() {
    if (!ADMIN_PASS || pwInput === ADMIN_PASS) {
      setAdminUnlocked(true)
      setShowAdminGate(false)
      if (pendingTabAfterUnlock.current) {
        setActiveTab(pendingTabAfterUnlock.current)
        pendingTabAfterUnlock.current = null
      }
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
      <div className="px-6 pt-6 pb-1 max-w-screen-2xl mx-auto">
        <div className="text-[1.6rem] font-bold text-[#1F3864] leading-tight mb-1">
          📦 AJW 생산자재 발주계획 시스템
        </div>
        <div className="text-[0.95rem] text-gray-500 mb-4">
          (주)에이제이월드 SCM팀 — 로우데이터 업로드 후 버튼 클릭으로 Excel 자동 생성
        </div>
      </div>

      {/* 최상위 탭 */}
      <div className="border-b border-gray-200 sticky top-0 bg-white z-10">
        <div className="flex overflow-x-auto max-w-screen-2xl mx-auto px-4">
          {NAV_TABS.map(t => {
            const isActive =
              t.id === 'steps' ? isStepId(activeTab) :
              t.id === 'admin' ? isAdminTabId(activeTab) :
              activeTab === t.id
            return (
              <button
                key={t.id}
                onClick={() => handleTabClick(t.id)}
                className={`px-4 py-2.5 text-sm font-medium whitespace-nowrap border-b-2 transition-colors flex-shrink-0 ${
                  isActive
                    ? 'border-[#FF4B4B] text-[#FF4B4B]'
                    : 'border-transparent text-gray-600 hover:text-gray-900 hover:border-gray-300'
                }`}
              >
                {t.label}
              </button>
            )
          })}
        </div>
      </div>

      {/* Content */}
      <div className="max-w-screen-2xl mx-auto px-6 py-6">
        {activeTab === 'home' && (
          <Dashboard
            metadata={metadata}
            inventory={inventory}
            salesAgg={salesAgg}
            onNavigate={(tab) => navigateTo(tab as TabId)}
          />
        )}

        {isStepId(activeTab) && (
          <div>
            <div className="flex border-b border-gray-200 mb-6">
              {STEP_TABS.map(s => (
                <button
                  key={s.id}
                  onClick={() => setActiveTab(s.id)}
                  className={`px-5 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors whitespace-nowrap ${
                    activeTab === s.id
                      ? 'border-[#2E75B6] text-[#2E75B6]'
                      : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                  }`}
                >
                  {s.label}
                </button>
              ))}
            </div>
            {activeTab === 'step1' && <Step1 metadata={metadata} setMetadata={setMetadata} settings={settings} />}
            {activeTab === 'step2' && <Step2 salesAgg={salesAgg} setSalesAgg={setSalesAgg} />}
            {activeTab === 'step3' && <Step3 metadata={metadata} inventory={inventory} settings={settings} salesAgg={salesAgg} />}
          </div>
        )}

        {activeTab === 'sales'   && <SalesAnalysisTab />}
        {activeTab === 'recon'   && <InventoryReconciliationTab />}
        {activeTab === 'import'  && <ImportTab />}
        {activeTab === 'partnum' && (
          <PartNumberGenerator
            ojcProducts={ojcProducts}
            setOjcProducts={(p) => { setOjcProducts(p); saveOjcProducts(p) }}
          />
        )}

        {isAdminTabId(activeTab) && (
          <div>
            <div className="flex border-b border-gray-200 mb-6">
              {ADMIN_TABS.map(s => (
                <button
                  key={s.id}
                  onClick={() => setActiveTab(s.id)}
                  className={`px-5 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors whitespace-nowrap ${
                    activeTab === s.id
                      ? 'border-[#2E75B6] text-[#2E75B6]'
                      : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                  }`}
                >
                  {s.label}
                </button>
              ))}
            </div>
            {activeTab === 'admin_material' && (
              <MaterialManager
                metadata={metadata}   setMetadata={setMetadata}
                inventory={inventory} setInventory={setInventory}
              />
            )}
            {activeTab === 'admin_ojcrules' && <OjcRulesEditor />}
            {activeTab === 'admin_settings' && (
              <Settings
                settings={settings}   setSettings={setSettings}
                metadata={metadata}   setMetadata={setMetadata}
                inventory={inventory} setInventory={setInventory}
                sales={sales}         salesAgg={salesAgg}
              />
            )}
          </div>
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
                관리자 탭은 비밀번호가 필요합니다.
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
