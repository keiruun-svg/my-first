import type { AppSettings, Metadata, Inventory } from '../lib/types'
import type { SalesAggResult } from '../lib/aggregate/salesAgg'

interface Props {
  metadata:  Metadata
  inventory: Inventory
  settings:  AppSettings
  salesAgg:  SalesAggResult | null
  onNavigate: (tab: string) => void
}

export default function Dashboard({ metadata, inventory, salesAgg, onNavigate }: Props) {
  // ── 품번 등록 현황 ────────────────────────────────────────────
  const cableEntries   = Object.entries(metadata.cable)
  const housingEntries = Object.entries(metadata.housing)

  const cableMissing   = cableEntries  .filter(([, v]) => !v.품번?.trim())
  const housingMissing = housingEntries.filter(([, v]) => {
    const first = Array.isArray(v) ? v[0] : v
    return !first?.품번?.trim()
  })
  const totalMissing = cableMissing.length + housingMissing.length

  // ── 재고 입력 현황 ────────────────────────────────────────────
  const cableZero   = Object.entries(inventory.cable)  .filter(([, v]) => v.현재고 === 0)
  const housingZero = Object.entries(inventory.housing).filter(([, v]) => {
    const first = Array.isArray(v) ? v[0] : v
    return (first as { 현재고: number })?.현재고 === 0
  })
  const totalZero = cableZero.length + housingZero.length

  // ── 3-STEP 준비 상태 ─────────────────────────────────────────
  const hasMetadata  = cableEntries.length > 0 || housingEntries.length > 0
  const hasSalesAgg  = salesAgg !== null
  const step3Ready   = hasMetadata && totalMissing === 0

  // ── OJC 판매 CAGR 요약 ───────────────────────────────────────
  const latestYear = salesAgg?.years[salesAgg.years.length - 1] ?? ''
const cagrEntries = salesAgg
    ? Object.entries(salesAgg.salesCagr)
        .map(([kind, cagr]) => ({
          kind,
          cagr,
          latestSales: salesAgg.byType[kind]?.[latestYear]?.sales ?? 0,
        }))
        .filter(e => e.latestSales > 0)
        .sort((a, b) => b.latestSales - a.latestSales)
    : []

  return (
    <div className="space-y-6">

      {/* ── KPI 카드 4개 ───────────────────────────────────────── */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">

        {/* 품번 미등록 */}
        <KpiCard
          label="품번 미등록"
          value={totalMissing === 0 ? '모두 완료' : `${totalMissing}개`}
          sub={totalMissing === 0 ? `케이블 ${cableEntries.length} / 하우징 ${housingEntries.length}` : '발주계획 생성 전 입력 필요'}
          status={totalMissing === 0 ? 'good' : 'warn'}
          onClick={() => onNavigate('meta')}
        />

        {/* 재고 미입력 */}
        <KpiCard
          label="재고 미입력"
          value={totalZero === 0 ? '모두 입력됨' : `${totalZero}개`}
          sub={totalZero === 0 ? '현재고 입력 완료' : '0으로 처리된 항목'}
          status={totalZero === 0 ? 'good' : 'info'}
          onClick={() => onNavigate('inventory')}
        />

        {/* 판매 분석 */}
        <KpiCard
          label="판매 분석"
          value={hasSalesAgg ? `${salesAgg!.years.length}개 연도` : '미업로드'}
          sub={hasSalesAgg ? `OJC ${cagrEntries.length}종 분석됨` : 'STEP 2에서 업로드 필요'}
          status={hasSalesAgg ? 'good' : 'info'}
          onClick={() => onNavigate('step2')}
        />

        {/* STEP 3 준비 */}
        <KpiCard
          label="발주계획 생성"
          value={step3Ready ? '준비 완료' : '준비 중'}
          sub={step3Ready ? 'STEP 3으로 이동' : `미등록 품번 ${totalMissing}개 해결 필요`}
          status={step3Ready ? 'good' : totalMissing > 0 ? 'warn' : 'info'}
          onClick={() => onNavigate('step3')}
        />
      </div>

      {/* ── 메인 콘텐츠 2열 ────────────────────────────────────── */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">

        {/* 3-STEP 워크플로우 상태 */}
        <div className="border border-gray-200 rounded-lg p-5">
          <h2 className="text-sm font-semibold text-gray-700 mb-4">3-STEP 워크플로우 진행 상태</h2>
          <div className="space-y-3">
            <StepRow
              step="STEP 1"
              label="ERP 파일 가공"
              done={hasMetadata}
              detail={hasMetadata ? `케이블 ${cableEntries.length}종 / 하우징 ${housingEntries.length}종 등록됨` : '가공파일 아직 업로드 안 됨'}
              onClick={() => onNavigate('step1')}
            />
            <StepRow
              step="STEP 2"
              label="판매 분석"
              done={hasSalesAgg}
              detail={hasSalesAgg ? `${salesAgg!.years[0]}~${latestYear}년 데이터 로드됨` : '판매 파일 아직 업로드 안 됨'}
              onClick={() => onNavigate('step2')}
            />
            <StepRow
              step="STEP 3"
              label="발주계획 생성"
              done={step3Ready}
              detail={step3Ready ? '가공파일 업로드 후 Excel 생성 가능' : `품번 미등록 ${totalMissing}개 해결 후 사용 가능`}
              onClick={() => onNavigate('step3')}
            />
          </div>
        </div>

        {/* 미등록 품번 목록 */}
        <div className="border border-gray-200 rounded-lg p-5">
          <h2 className="text-sm font-semibold text-gray-700 mb-4">
            미등록 품번 현황
            {totalMissing > 0 && (
              <span className="ml-2 px-2 py-0.5 bg-orange-100 text-orange-700 rounded text-xs">{totalMissing}건</span>
            )}
          </h2>
          {totalMissing === 0 ? (
            <p className="text-sm text-green-600">모든 품번이 등록되어 있습니다.</p>
          ) : (
            <div className="space-y-2 max-h-48 overflow-y-auto">
              {cableMissing.length > 0 && (
                <div>
                  <div className="text-xs font-medium text-gray-500 mb-1">케이블 ({cableMissing.length})</div>
                  {cableMissing.map(([key]) => (
                    <div key={key} className="text-xs text-gray-700 px-2 py-1 bg-orange-50 rounded mb-1">{key}</div>
                  ))}
                </div>
              )}
              {housingMissing.length > 0 && (
                <div>
                  <div className="text-xs font-medium text-gray-500 mb-1">하우징 ({housingMissing.length})</div>
                  {housingMissing.map(([key]) => (
                    <div key={key} className="text-xs text-gray-700 px-2 py-1 bg-orange-50 rounded mb-1">{key}</div>
                  ))}
                </div>
              )}
              <button
                onClick={() => onNavigate('meta')}
                className="mt-2 text-xs text-blue-600 hover:underline"
              >
                품번 관리로 이동 →
              </button>
            </div>
          )}
        </div>
      </div>


    </div>
  )
}

// ── 공용 서브 컴포넌트 ─────────────────────────────────────────

type Status = 'good' | 'warn' | 'info'

function KpiCard({ label, value, sub, status, onClick }: {
  label: string; value: string; sub: string; status: Status; onClick: () => void
}) {
  const colors: Record<Status, string> = {
    good: 'border-green-200 bg-green-50',
    warn: 'border-orange-200 bg-orange-50',
    info: 'border-gray-200 bg-gray-50',
  }
  const valueColors: Record<Status, string> = {
    good: 'text-green-700',
    warn: 'text-orange-700',
    info: 'text-gray-700',
  }
  return (
    <button
      onClick={onClick}
      className={`text-left border rounded-lg p-4 hover:shadow-sm transition ${colors[status]}`}
    >
      <div className="text-xs text-gray-500 mb-1">{label}</div>
      <div className={`text-lg font-bold mb-1 ${valueColors[status]}`}>{value}</div>
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
