import { useState, useEffect } from 'react'
import type { Metadata, Inventory } from '../lib/types'
import type { SalesAggResult } from '../lib/aggregate/salesAgg'
import { loadImportOrders, loadReconHistory } from '../lib/supabase'
import type { ImportOrder, ReconHistorySummary } from '../lib/supabase'

interface Props {
  metadata:   Metadata
  inventory:  Inventory
  salesAgg:   SalesAggResult | null
  onNavigate: (tab: string) => void
}

export default function Dashboard({ metadata, inventory, salesAgg, onNavigate }: Props) {
  const [orders,       setOrders]       = useState<ImportOrder[]>([])
  const [reconHistory, setReconHistory] = useState<ReconHistorySummary[]>([])

  useEffect(() => {
    setOrders(loadImportOrders())
    setReconHistory(loadReconHistory())
  }, [])

  // ── 품번 등록 현황 ────────────────────────────────────────────
  const cableEntries   = Object.entries(metadata.cable)
  const housingEntries = Object.entries(metadata.housing)
  const cableMissing   = cableEntries.filter(([, v]) => !v.품번?.trim())
  const housingMissing = housingEntries.filter(([, v]) => {
    const first = Array.isArray(v) ? v[0] : v
    return !first?.품번?.trim()
  })
  const totalMissing = cableMissing.length + housingMissing.length

  // ── 재고 입력 현황 ────────────────────────────────────────────
  const cableZero   = Object.entries(inventory.cable).filter(([, v]) => v.현재고 === 0)
  const housingZero = Object.entries(inventory.housing).filter(([, v]) => {
    const first = Array.isArray(v) ? v[0] : v
    return (first as { 현재고: number })?.현재고 === 0
  })
  const totalZero = cableZero.length + housingZero.length

  // ── 3-STEP 준비 상태 ─────────────────────────────────────────
  const hasMetadata = cableEntries.length > 0 || housingEntries.length > 0
  const hasSalesAgg = salesAgg !== null
  const step3Ready  = hasMetadata && totalMissing === 0

  // ── 수입 발주 현황 ────────────────────────────────────────────
  function dday(expectedDate: string): number {
    const exp = new Date(expectedDate + 'T00:00:00')
    const tod = new Date(); tod.setHours(0, 0, 0, 0)
    return Math.round((exp.getTime() - tod.getTime()) / 86400000)
  }
  function orderLabel(o: ImportOrder) {
    return `${o.vendorCode} ${o.year}-${String(o.seq).padStart(2, '0')}차`
  }

  const activeOrders  = orders.filter(o => !o.actualDate)
  const overdueOrders = activeOrders.filter(o => dday(o.expectedDate) < 0)
  const urgentOrders  = activeOrders.filter(o => { const d = dday(o.expectedDate); return d >= 0 && d <= 14 })
  const waitingOrders = activeOrders.filter(o => dday(o.expectedDate) > 14)
  const ddayItems     = [...overdueOrders, ...urgentOrders]
    .sort((a, b) => dday(a.expectedDate) - dday(b.expectedDate))
    .slice(0, 5)

  // ── 재고 대사 현황 ────────────────────────────────────────────
  const latestRecon    = [...reconHistory].sort((a, b) => b.savedAt.localeCompare(a.savedAt))[0] ?? null
  const reconDaysAgo   = latestRecon
    ? Math.floor((Date.now() - new Date(latestRecon.savedAt).getTime()) / 86400000)
    : null
  const reconAnomalies = latestRecon
    ? latestRecon.diff + latestRecon.emp_only + latestRecon.ecount_only
    : 0

  // ── OJC 판매 CAGR ────────────────────────────────────────────
  const latestYear  = salesAgg?.years[salesAgg.years.length - 1] ?? ''
  const cagrEntries = salesAgg
    ? Object.entries(salesAgg.salesCagr)
        .map(([kind, cagr]) => ({
          kind, cagr,
          latestSales: salesAgg.byType[kind]?.[latestYear]?.sales ?? 0,
        }))
        .filter(e => e.latestSales > 0)
        .sort((a, b) => b.latestSales - a.latestSales)
        .slice(0, 5)
    : []

  return (
    <div className="space-y-6">

      {/* ── KPI 카드 6개 ───────────────────────────────────────── */}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4">
        <KpiCard
          label="품번 미등록"
          value={totalMissing === 0 ? '모두 완료' : `${totalMissing}개`}
          sub={totalMissing === 0 ? `케이블 ${cableEntries.length} / 하우징 ${housingEntries.length}` : '발주계획 전 입력 필요'}
          status={totalMissing === 0 ? 'good' : 'warn'}
          onClick={() => onNavigate('admin_material')}
        />
        <KpiCard
          label="재고 미입력"
          value={totalZero === 0 ? '모두 입력됨' : `${totalZero}개`}
          sub={totalZero === 0 ? '현재고 입력 완료' : '0으로 처리된 항목'}
          status={totalZero === 0 ? 'good' : 'info'}
          onClick={() => onNavigate('admin_material')}
        />
        <KpiCard
          label="판매 분석"
          value={hasSalesAgg ? `${salesAgg!.years.length}개 연도` : '미업로드'}
          sub={hasSalesAgg ? `${salesAgg!.years[0]}~${latestYear}년` : 'STEP 2에서 업로드 필요'}
          status={hasSalesAgg ? 'good' : 'info'}
          onClick={() => onNavigate('step2')}
        />
        <KpiCard
          label="발주계획 생성"
          value={step3Ready ? '준비 완료' : '준비 중'}
          sub={step3Ready ? 'STEP 3으로 이동' : `미등록 ${totalMissing}개 해결 필요`}
          status={step3Ready ? 'good' : totalMissing > 0 ? 'warn' : 'info'}
          onClick={() => onNavigate('step3')}
        />
        <KpiCard
          label="수입 발주"
          value={
            activeOrders.length === 0 ? '진행 없음' :
            overdueOrders.length > 0  ? `지연 ${overdueOrders.length}건` :
            `대기 ${activeOrders.length}건`
          }
          sub={
            activeOrders.length === 0   ? '등록된 발주 없음' :
            urgentOrders.length > 0     ? `14일 이내 ${urgentOrders.length}건` :
            `대기 ${waitingOrders.length}건`
          }
          status={overdueOrders.length > 0 ? 'warn' : activeOrders.length > 0 ? 'info' : 'good'}
          onClick={() => onNavigate('import')}
        />
        <KpiCard
          label="재고 대사"
          value={latestRecon === null ? '미실행' : reconDaysAgo === 0 ? '오늘 실행' : `${reconDaysAgo}일 전`}
          sub={
            latestRecon === null    ? '대사 실행 필요' :
            reconAnomalies === 0    ? '불일치 없음 ✓' :
            `불일치 ${reconAnomalies}건`
          }
          status={
            latestRecon === null                            ? 'warn' :
            reconAnomalies > 0                             ? 'warn' :
            reconDaysAgo !== null && reconDaysAgo > 7      ? 'info' :
            'good'
          }
          onClick={() => onNavigate('recon')}
        />
      </div>

      {/* ── 메인 3열 ───────────────────────────────────────────── */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">

        {/* 3-STEP 워크플로우 */}
        <div className="border border-gray-200 rounded-lg p-5">
          <h2 className="text-sm font-semibold text-gray-700 mb-4">생산자재 발주계획 진행 상태</h2>
          <div className="space-y-2">
            <StepRow
              step="STEP 1" label="ERP 파일 가공" done={hasMetadata}
              detail={hasMetadata
                ? `케이블 ${cableEntries.length}종 / 하우징 ${housingEntries.length}종`
                : '가공파일 미업로드'}
              onClick={() => onNavigate('step1')}
            />
            <StepRow
              step="STEP 2" label="판매 분석" done={hasSalesAgg}
              detail={hasSalesAgg
                ? `${salesAgg!.years[0]}~${latestYear}년 로드됨`
                : '판매 파일 미업로드'}
              onClick={() => onNavigate('step2')}
            />
            <StepRow
              step="STEP 3" label="발주계획 생성" done={step3Ready}
              detail={step3Ready
                ? '가공파일 업로드 후 Excel 생성 가능'
                : `품번 미등록 ${totalMissing}개 해결 필요`}
              onClick={() => onNavigate('step3')}
            />
          </div>
        </div>

        {/* 수입 발주 D-day */}
        <div className="border border-gray-200 rounded-lg p-5">
          <h2 className="text-sm font-semibold text-gray-700 mb-4">
            수입 발주 현황
            {overdueOrders.length > 0 && (
              <span className="ml-2 px-2 py-0.5 bg-red-100 text-red-700 rounded text-xs">
                지연 {overdueOrders.length}건
              </span>
            )}
          </h2>
          {activeOrders.length === 0 ? (
            <p className="text-sm text-gray-400">진행 중인 발주가 없습니다.</p>
          ) : (
            <div className="space-y-2">
              {ddayItems.map(o => (
                <div key={o.id} className="flex items-center justify-between">
                  <div className="min-w-0">
                    <span className="text-sm font-medium text-gray-700">{orderLabel(o)}</span>
                    <span className="ml-1.5 text-xs text-gray-400">{o.vendorName}</span>
                  </div>
                  <DdayBadge d={dday(o.expectedDate)} />
                </div>
              ))}
              {waitingOrders.length > 0 && (
                <p className="text-xs text-gray-400 pt-1">외 {waitingOrders.length}건 대기 중</p>
              )}
              <button onClick={() => onNavigate('import')} className="mt-1 text-xs text-blue-600 hover:underline">
                전체 발주 현황 →
              </button>
            </div>
          )}
        </div>

        {/* 재고 대사 + OJC CAGR */}
        <div className="border border-gray-200 rounded-lg p-5 space-y-5">

          {/* 재고 대사 */}
          <div>
            <h2 className="text-sm font-semibold text-gray-700 mb-3">재고 대사 현황</h2>
            {latestRecon === null ? (
              <p className="text-sm text-orange-600">아직 실행된 재고 대사가 없습니다.</p>
            ) : (
              <div className="space-y-2 text-sm">
                <div className="flex justify-between">
                  <span className="text-gray-500">마지막 실행</span>
                  <span className="font-medium text-gray-700">
                    {latestRecon.date.slice(0, 2)}/{latestRecon.date.slice(2)}
                    <span className="ml-1 text-gray-400 font-normal">
                      ({reconDaysAgo === 0 ? '오늘' : `${reconDaysAgo}일 전`})
                    </span>
                  </span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-500">총 품목</span>
                  <span className="text-gray-700">{latestRecon.total}개</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-500">불일치</span>
                  <span className={`font-medium ${reconAnomalies > 0 ? 'text-red-600' : 'text-green-600'}`}>
                    {reconAnomalies}건{reconAnomalies === 0 && ' ✓'}
                  </span>
                </div>
                {reconDaysAgo !== null && reconDaysAgo > 7 && (
                  <p className="text-xs text-orange-500">⚠ {reconDaysAgo}일 경과 — 재실행 권장</p>
                )}
              </div>
            )}
            <button onClick={() => onNavigate('recon')} className="mt-2 text-xs text-blue-600 hover:underline">
              재고 대사 실행 →
            </button>
          </div>

          {/* OJC 판매 CAGR */}
          {cagrEntries.length > 0 && (
            <div className="border-t border-gray-100 pt-4">
              <h2 className="text-sm font-semibold text-gray-700 mb-3">
                OJC 판매 성장률
                <span className="ml-1 text-xs text-gray-400 font-normal">{latestYear}년 기준</span>
              </h2>
              <div className="space-y-1.5">
                {cagrEntries.map(e => (
                  <div key={e.kind} className="flex items-center justify-between text-xs">
                    <span className="text-gray-600 truncate max-w-[65%]">{e.kind}</span>
                    <span className={`font-mono font-bold tabular-nums ${e.cagr >= 0 ? 'text-green-600' : 'text-red-500'}`}>
                      {e.cagr >= 0 ? '+' : ''}{(e.cagr * 100).toFixed(1)}%
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* ── 미등록 품번 상세 (있을 때만) ──────────────────────── */}
      {totalMissing > 0 && (
        <div className="border border-orange-200 rounded-lg p-5 bg-orange-50">
          <h2 className="text-sm font-semibold text-orange-800 mb-3">
            미등록 품번 현황
            <span className="ml-2 px-2 py-0.5 bg-orange-200 text-orange-800 rounded text-xs">{totalMissing}건</span>
          </h2>
          <div className="flex gap-8 flex-wrap">
            {cableMissing.length > 0 && (
              <div>
                <div className="text-xs font-medium text-gray-500 mb-1">케이블 ({cableMissing.length})</div>
                <div className="space-y-1 max-h-32 overflow-y-auto">
                  {cableMissing.map(([key]) => (
                    <div key={key} className="text-xs text-gray-700 px-2 py-0.5 bg-white border border-orange-100 rounded">{key}</div>
                  ))}
                </div>
              </div>
            )}
            {housingMissing.length > 0 && (
              <div>
                <div className="text-xs font-medium text-gray-500 mb-1">하우징 ({housingMissing.length})</div>
                <div className="space-y-1 max-h-32 overflow-y-auto">
                  {housingMissing.map(([key]) => (
                    <div key={key} className="text-xs text-gray-700 px-2 py-0.5 bg-white border border-orange-100 rounded">{key}</div>
                  ))}
                </div>
              </div>
            )}
          </div>
          <button onClick={() => onNavigate('admin_material')} className="mt-3 text-xs text-blue-600 hover:underline">
            자재 관리로 이동 →
          </button>
        </div>
      )}
    </div>
  )
}

// ── 서브 컴포넌트 ──────────────────────────────────────────────

type Status = 'good' | 'warn' | 'info'

function KpiCard({ label, value, sub, status, onClick }: {
  label: string; value: string; sub: string; status: Status; onClick: () => void
}) {
  const border: Record<Status, string> = {
    good: 'border-green-200 bg-green-50',
    warn: 'border-orange-200 bg-orange-50',
    info: 'border-gray-200 bg-gray-50',
  }
  const valueColor: Record<Status, string> = {
    good: 'text-green-700',
    warn: 'text-orange-700',
    info: 'text-gray-700',
  }
  return (
    <button onClick={onClick} className={`text-left border rounded-lg p-4 hover:shadow-sm transition ${border[status]}`}>
      <div className="text-xs text-gray-500 mb-1">{label}</div>
      <div className={`text-lg font-bold mb-1 ${valueColor[status]}`}>{value}</div>
      <div className="text-xs text-gray-500">{sub}</div>
    </button>
  )
}

function StepRow({ step, label, done, detail, onClick }: {
  step: string; label: string; done: boolean; detail: string; onClick: () => void
}) {
  return (
    <button onClick={onClick} className="w-full text-left flex items-start gap-3 p-3 rounded-lg hover:bg-gray-50 transition">
      <div className={`mt-0.5 w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold flex-shrink-0 ${
        done ? 'bg-green-100 text-green-700' : 'bg-gray-100 text-gray-400'
      }`}>
        {done ? '✓' : '—'}
      </div>
      <div>
        <div className="text-xs font-semibold text-gray-700">{step} — {label}</div>
        <div className="text-xs text-gray-400 mt-0.5">{detail}</div>
      </div>
    </button>
  )
}

function DdayBadge({ d }: { d: number }) {
  if (d < 0)  return <span className="text-xs font-bold text-red-600    bg-red-50    px-2 py-0.5 rounded flex-shrink-0">D+{Math.abs(d)}</span>
  if (d === 0) return <span className="text-xs font-bold text-orange-600 bg-orange-50 px-2 py-0.5 rounded flex-shrink-0">D-Day</span>
  if (d <= 7)  return <span className="text-xs font-bold text-orange-500 bg-orange-50 px-2 py-0.5 rounded flex-shrink-0">D-{d}</span>
  return              <span className="text-xs font-medium text-gray-500  bg-gray-100  px-2 py-0.5 rounded flex-shrink-0">D-{d}</span>
}
