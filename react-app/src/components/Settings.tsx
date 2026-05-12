import { useState } from 'react'
import { saveSettings, syncToSupabase, sb } from '../lib/supabase'
import type { AppSettings, Metadata, Inventory, SalesAnalysis } from '../lib/types'

interface Props {
  settings: AppSettings
  setSettings: (s: AppSettings) => void
  metadata: Metadata
  inventory: Inventory
  sales: SalesAnalysis
}

export default function Settings({ settings, setSettings, metadata, inventory, sales }: Props) {
  const [saved, setSaved] = useState(false)
  const [syncStatus, setSyncStatus] = useState<string | null>(null)
  const [syncing, setSyncing] = useState(false)

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
      const results = await syncToSupabase(settings, metadata, inventory, sales)
      const failed = Object.entries(results).filter(([, v]) => !v).map(([k]) => k)
      if (!failed.length) {
        setSyncStatus('✅ Supabase 동기화 완료! (설정 / 품번 메타 / 재고 / 판매 분석)')
      } else {
        setSyncStatus(`❌ 동기화 실패 항목: ${failed.join(', ')}`)
      }
    } catch (e) {
      setSyncStatus(`❌ 오류: ${e}`)
    } finally {
      setSyncing(false)
    }
  }

  return (
    <div className="space-y-6">
      <div className="bg-gray-50 border border-gray-200 rounded p-4">
        <h3 className="font-bold text-gray-800 mb-1">⚙️ 파라미터 설정</h3>
        <p className="text-sm text-gray-600">변경 후 저장 버튼을 눌러야 반영됩니다.</p>
      </div>

      <div className="bg-white border rounded-lg p-5 space-y-4">
        <div>
          <label className="block text-sm font-semibold text-gray-700 mb-1">기본 리드타임 (일)</label>
          <input type="number" min="1" max="365"
            value={settings.lead_time_default}
            onChange={e => setSettings({ ...settings, lead_time_default: parseInt(e.target.value) || 60 })}
            className="w-32 border rounded px-3 py-1.5 text-sm focus:outline-blue-400"
          />
          <p className="text-xs text-gray-400 mt-1">품번별 리드타임이 없을 때 사용되는 기본값</p>
        </div>

        <div>
          <label className="block text-sm font-semibold text-gray-700 mb-2">연도별 색상 (HEX, # 제외)</label>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            {([
              ['main_header','메인 헤더'],
              ['year_23','2023년'],
              ['year_24','2024년'],
              ['year_25','2025년'],
            ] as const).map(([key, label]) => (
              <div key={key}>
                <label className="text-xs text-gray-500 mb-1 block">{label}</label>
                <div className="flex items-center gap-2">
                  <div className="w-6 h-6 rounded border" style={{ background: '#' + settings.colors[key] }} />
                  <input
                    value={settings.colors[key]}
                    onChange={e => setSettings({ ...settings, colors: { ...settings.colors, [key]: e.target.value } })}
                    className="flex-1 border rounded px-2 py-1 text-xs font-mono focus:outline-blue-400"
                    maxLength={6}
                  />
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="flex items-center gap-3">
        <button onClick={save}
          className="bg-gray-700 hover:bg-gray-800 text-white font-bold px-6 py-2 rounded-lg transition">
          💾 설정 저장
        </button>
        {saved && <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>}
      </div>

      <hr />

      <div className="bg-blue-50 border border-blue-200 rounded-lg p-5 space-y-3">
        <h4 className="font-bold text-blue-800">☁️ Supabase 동기화</h4>
        <p className="text-sm text-blue-700">
          로컬(localStorage)에 저장된 데이터를 Supabase 클라우드에 수동 업로드합니다.
          {!sb && <span className="ml-1 text-red-600 font-semibold">⚠ .env.local에 API 키가 없어 비활성화됨</span>}
        </p>
        <button
          onClick={sync}
          disabled={syncing || !sb}
          className="bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white font-semibold px-6 py-2 rounded-lg transition"
        >
          {syncing ? '⏳ 동기화 중...' : '☁️ Supabase에 동기화'}
        </button>
        {syncStatus && (
          <div className={`rounded p-3 text-sm font-medium ${syncStatus.startsWith('✅') ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>
            {syncStatus}
          </div>
        )}
      </div>

      <hr />

      <div className="bg-gray-50 border rounded-lg p-4">
        <h4 className="font-semibold text-gray-700 mb-2">워크플로우 안내</h4>
        <ol className="text-sm text-gray-600 space-y-1 list-decimal list-inside">
          <li>맥산 ERP에서 구매조회 또는 구매현황 파일 추출</li>
          <li>STEP 1 실행 → 사용 케이블·하우징 타입 자동 감지</li>
          <li>품번 관리 탭에서 품번·품명·구매처·리드타임 입력 후 저장</li>
          <li>재고 현황 탭에서 현재고·기발주 입력 후 저장</li>
          <li>STEP 2 실행 (선택) → 판매량 파일 분석</li>
          <li>STEP 3 실행 → 2026_연간발주계획.xlsx 다운로드</li>
          <li>노란색 셀에 2026 목표 발주량 입력</li>
        </ol>
      </div>
    </div>
  )
}
