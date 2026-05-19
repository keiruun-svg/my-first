import { useState } from 'react'
import { saveSettings, syncToSupabase, migratePigtailKeys, sb, CAN_WRITE } from '../lib/supabase'
import type { AppSettings, Metadata, Inventory, SalesAnalysis } from '../lib/types'
import type { SalesAggResult } from '../lib/aggregate/salesAgg'

interface Props {
  settings:     AppSettings
  setSettings:  (s: AppSettings) => void
  metadata:     Metadata
  setMetadata:  (m: Metadata) => void
  inventory:    Inventory
  setInventory: (i: Inventory) => void
  sales:        SalesAnalysis
  salesAgg?:    SalesAggResult | null
}

export default function Settings({ settings, setSettings, metadata, setMetadata, inventory, setInventory, sales, salesAgg }: Props) {
  const [saved, setSaved]                 = useState(false)
  const [syncStatus, setSyncStatus]       = useState<string | null>(null)
  const [syncing, setSyncing]             = useState(false)
  const [migrStatus, setMigrStatus]       = useState<string | null>(null)
  const [migrating, setMigrating]         = useState(false)

  const save = () => {
    saveSettings(settings)
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  const sync = async () => {
    if (!sb) {
      setSyncStatus('❌ Supabase 연결 없음 — .env.local에 VITE_SUPABASE_URL / VITE_SUPABASE_KEY를 설정하세요.')
      return
    }
    setSyncing(true); setSyncStatus(null)
    try {
      const results = await syncToSupabase(settings, metadata, inventory, sales, salesAgg)
      const failed  = Object.entries(results).filter(([, v]) => !v).map(([k]) => k)
      setSyncStatus(failed.length
        ? `❌ 동기화 실패: ${failed.join(', ')}`
        : '✅ Supabase 동기화 완료 (설정 / 품번 메타 / 재고 / 판매 분석 / 판매 집계)')
    } catch (e) {
      setSyncStatus(`❌ 오류: ${e}`)
    } finally {
      setSyncing(false)
    }
  }

  const runMigration = async () => {
    setMigrating(true); setMigrStatus(null)
    try {
      const { metadata: newMeta, inventory: newInv, changed } = await migratePigtailKeys(metadata, inventory)
      if (changed === 0) {
        setMigrStatus('ℹ 변경 대상 없음 — 이미 마이그레이션되었거나 해당 키가 없습니다.')
      } else {
        setMetadata(newMeta)
        setInventory(newInv)
        setMigrStatus(`✅ ${changed}개 키 변경 완료 (pigtail-연청→AQUA, pigtail-연등→rose)${CAN_WRITE ? ' — Supabase 동기화 완료' : ' — 로컬만 반영. Supabase 동기화 필요'}`)
      }
    } catch (e) {
      setMigrStatus(`❌ 오류: ${e}`)
    } finally {
      setMigrating(false)
    }
  }

  const isConnected = !!sb
  const nCable   = Object.keys(metadata.cable).length
  const nHousing = Object.keys(metadata.housing).length

  return (
    <div className="space-y-6 max-w-2xl">

      {/* 파라미터 */}
      <div>
        <div className="text-sm font-semibold text-gray-700 mb-3">⚙️ 발주 파라미터</div>
        <div className="bg-white border border-gray-200 rounded-lg divide-y divide-gray-100">

          <div className="flex items-center justify-between px-5 py-3">
            <div>
              <div className="text-sm font-medium text-gray-700">기본 리드타임</div>
              <div className="text-xs text-gray-400 mt-0.5">품번별 리드타임이 없을 때 사용</div>
            </div>
            <div className="flex items-center gap-2">
              <input type="number" min="1" max="365"
                value={settings.lead_time_default}
                onChange={e => setSettings({ ...settings, lead_time_default: parseInt(e.target.value) || 60 })}
                className="w-24 border border-gray-300 rounded px-3 py-1.5 text-sm text-right focus:outline-none focus:border-blue-400"
              />
              <span className="text-sm text-gray-500">일</span>
            </div>
          </div>

          <div className="flex items-center justify-between px-5 py-3">
            <div>
              <div className="text-sm font-medium text-gray-700">안전재고 계수 k</div>
              <div className="text-xs text-gray-400 mt-0.5">안전재고 = 월평균 × (LT/30) × k</div>
            </div>
            <div className="flex items-center gap-2">
              <input type="number" min="0.5" max="3.0" step="0.1"
                value={settings.safety_stock_k ?? 1.5}
                onChange={e => setSettings({ ...settings, safety_stock_k: parseFloat(e.target.value) || 1.5 })}
                className="w-24 border border-gray-300 rounded px-3 py-1.5 text-sm text-right focus:outline-none focus:border-blue-400"
              />
              <span className="text-xs text-gray-400">권장 1.5~2.0</span>
            </div>
          </div>

        </div>
      </div>

      {/* 색상 */}
      <div>
        <div className="text-sm font-semibold text-gray-700 mb-3">🎨 연도별 색상 (HEX, # 제외)</div>
        <div className="bg-white border border-gray-200 rounded-lg p-4">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            {([
              ['main_header', '메인 헤더'],
              ['year_23',     '2023년'],
              ['year_24',     '2024년'],
              ['year_25',     '2025년'],
            ] as const).map(([key, label]) => (
              <div key={key}>
                <div className="text-xs text-gray-500 mb-1.5">{label}</div>
                <div className="flex items-center gap-2">
                  <div className="w-7 h-7 rounded border border-gray-200 flex-shrink-0"
                       style={{ background: '#' + settings.colors[key] }} />
                  <input
                    value={settings.colors[key]}
                    onChange={e => setSettings({ ...settings, colors: { ...settings.colors, [key]: e.target.value } })}
                    className="flex-1 min-w-0 border border-gray-300 rounded px-2 py-1 text-xs font-mono focus:outline-none focus:border-blue-400"
                    maxLength={6} placeholder="1F3864"
                  />
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* 저장 버튼 */}
      <div className="flex items-center gap-3">
        <button onClick={save}
          className="bg-[#FF4B4B] hover:bg-[#e03030] text-white font-semibold px-6 py-2 rounded-lg transition text-sm">
          💾 파라미터 저장
        </button>
        {saved && <span className="text-green-600 text-sm font-medium">✅ 저장됐습니다.</span>}
      </div>

      <hr className="border-gray-200" />

      {/* Supabase 동기화 */}
      <div>
        <div className="flex items-center gap-2 mb-3">
          <div className="text-sm font-semibold text-gray-700">☁️ Supabase 동기화</div>
          <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${isConnected ? 'bg-green-100 text-green-700' : 'bg-gray-100 text-gray-500'}`}>
            {isConnected ? '● 연결됨' : '○ 미연결'}
          </span>
        </div>

        {CAN_WRITE ? (
          <div className="bg-gray-50 border border-gray-200 rounded-lg p-4 space-y-3">
            <div className="flex items-start justify-between gap-4">
              <div>
                <div className="text-sm text-gray-600">
                  로컬 데이터를 Supabase에 업로드합니다. Vercel에서 즉시 반영됩니다.
                </div>
                <div className="text-xs text-gray-400 mt-1">
                  케이블 {nCable}타입 · 하우징 {nHousing}타입 · 재고 · 판매 집계
                </div>
              </div>
              <button
                onClick={sync}
                disabled={syncing || !isConnected}
                className="flex-shrink-0 bg-[#2E75B6] hover:bg-[#1a5fa0] disabled:bg-gray-300 text-white font-semibold px-5 py-2 rounded-lg transition text-sm"
              >
                {syncing ? '⏳ 동기화 중...' : '☁️ 동기화'}
              </button>
            </div>
            {!isConnected && (
              <p className="text-xs text-orange-600">
                ⚠ .env.local에 VITE_SUPABASE_URL / VITE_SUPABASE_KEY를 설정하면 활성화됩니다.
              </p>
            )}
            {syncStatus && (
              <div className={`rounded-md px-3 py-2 text-sm font-medium ${syncStatus.startsWith('✅') ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>
                {syncStatus}
              </div>
            )}
          </div>
        ) : (
          <div className="bg-gray-50 border border-gray-200 rounded-lg p-4">
            <div className="text-sm text-gray-500">
              🔒 프로덕션 환경 — Supabase 쓰기는 로컬 개발 환경에서만 가능합니다.
            </div>
            <div className="text-xs text-gray-400 mt-1">
              데이터는 Supabase에서 읽기 전용으로 로드됩니다.
              {isConnected ? ' (현재 Supabase 연결됨 ✓)' : ' (Supabase 미연결 — 로컬 저장소 사용)'}
            </div>
          </div>
        )}
      </div>

      <hr className="border-gray-200" />

      {/* 피그테일 키 마이그레이션 */}
      <div>
        <div className="text-sm font-semibold text-gray-700 mb-3">🔧 데이터 마이그레이션</div>
        <div className="bg-amber-50 border border-amber-200 rounded-lg p-4 space-y-3">
          <div className="flex items-start justify-between gap-4">
            <div>
              <div className="text-sm text-gray-700 font-medium">피그테일 키 이름 변경</div>
              <div className="text-xs text-gray-500 mt-1 space-y-0.5">
                <div><code className="bg-white px-1 rounded">pigtail-연청</code> → <code className="bg-white px-1 rounded">pigtail-AQUA</code> (P14-RM-417D)</div>
                <div><code className="bg-white px-1 rounded">pigtail-연등</code> → <code className="bg-white px-1 rounded">pigtail-rose</code> (P14-RM-417C)</div>
              </div>
            </div>
            <button
              onClick={runMigration}
              disabled={migrating}
              className="flex-shrink-0 bg-amber-500 hover:bg-amber-600 disabled:bg-gray-300 text-white font-semibold px-4 py-2 rounded-lg transition text-sm"
            >
              {migrating ? '⏳ 처리 중...' : '🔄 마이그레이션 실행'}
            </button>
          </div>
          {migrStatus && (
            <div className={`rounded-md px-3 py-2 text-sm font-medium ${migrStatus.startsWith('✅') ? 'bg-green-100 text-green-800' : migrStatus.startsWith('ℹ') ? 'bg-blue-50 text-blue-700' : 'bg-red-100 text-red-800'}`}>
              {migrStatus}
            </div>
          )}
        </div>
      </div>

    </div>
  )
}
