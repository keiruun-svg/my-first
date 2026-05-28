import type { CostEntry } from './parse/parseCostFile'

// ─────────────────────────────────────────────────────────────────────────────
// 수익성 분석 엔진
//
// 수입(FLC) vs 국내생산(맥산) 권고 로직을 담은 순수 함수 모음.
// UI 의존성 없음 — 다른 탭에서도 import 가능.
// ─────────────────────────────────────────────────────────────────────────────

export interface ProfitRow {
  code:           string
  name:           string
  category:       string
  maeksanApplied: number | null   // 품목 카테고리 기준 적용 맥산 원가 (KT→표준, 기타→생산)
  flcApplied:     number | null   // 품목 카테고리 기준 적용 FLC 원가
  판매가:         number | null   // 계약 단가 우선, 없으면 판매 이력 평균
  isContract:     boolean
  avgAnnualQty:   number
  마진율맥산:     number | null   // (판매가 - 맥산원가) / 판매가 × 100
  마진율FLC:      number | null   // (판매가 - FLC원가)  / 판매가 × 100
  연간이익맥산:   number | null
  연간이익FLC:    number | null
}

export type RecoBadge = { label: string; bg: string; text: string }

/** KT향 → 표준원가, LG향·기타 → 생산원가 */
export function applyCost(entry: CostEntry, category: string): number {
  return category.startsWith('KT') ? entry.표준원가 : entry.생산원가
}

export function 적용원가명(category: string): string {
  return category.startsWith('KT') ? '표준원가' : '생산원가'
}

/**
 * 품목별 생산방식 권고를 반환한다.
 *
 * @param threshold     맥산/FLC 원가 비율 기준 (예: 105 = "맥산이 FLC 대비 5%까지 비싸도 생산 유지")
 * @param minMarginPct  이 값 미만 마진이면 무조건 철수검토 (영업 부서 기준값 사용 권장)
 *
 * 생산유지기준 = max(20, minMarginPct × 3)
 *   - 검토필요 구간: minMarginPct ~ 생산유지기준
 *   - 생산/수입 권고: 생산유지기준 이상
 *
 * ⚖ 맥산 생산 가능: 맥산이 FLC보다 비싸지만 가중치(threshold) 덕분에 수입 권고를 피한 케이스.
 *   가중치 OFF(threshold=100)이거나 맥산이 FLC보다 원래 저렴하면 ✅ 맥산 생산 권고.
 */
export function getReco(row: ProfitRow, threshold: number, minMarginPct: number): RecoBadge {
  const { 판매가, maeksanApplied, flcApplied, 마진율맥산, 마진율FLC } = row
  if (판매가 === null || 판매가 <= 0)
    return { label: '판매가 없음', bg: 'bg-gray-100', text: 'text-gray-400' }

  const 생산유지기준 = Math.max(20, minMarginPct * 3)
  const 맥산미달 = 마진율맥산 !== null && 마진율맥산 < minMarginPct
  const FLC미달  = 마진율FLC  !== null && 마진율FLC  < minMarginPct

  // ── 양쪽 원가 모두 있을 때 ──────────────────────────────────────────────────
  if (maeksanApplied !== null && flcApplied !== null) {
    if (맥산미달 && FLC미달)
      return { label: '⚠ 철수검토', bg: 'bg-red-100', text: 'text-red-700' }

    if (맥산미달 && !FLC미달)
      return { label: '📦 수입 권고', bg: 'bg-blue-100', text: 'text-blue-700' }

    if (!맥산미달 && !FLC미달 && flcApplied > 0) {
      if (maeksanApplied > flcApplied * (threshold / 100))
        return { label: '📦 수입 권고', bg: 'bg-blue-100', text: 'text-blue-700' }
    }

    if (마진율맥산 === null) return { label: '—', bg: 'bg-gray-100', text: 'text-gray-400' }

    if (마진율맥산 >= 생산유지기준) {
      // 맥산이 FLC보다 비싸지만 가중치 덕분에 수입 권고를 피한 경우 → 가능(caution)
      if (flcApplied > 0 && maeksanApplied > flcApplied)
        return { label: '⚖ 맥산 생산 가능', bg: 'bg-orange-100', text: 'text-orange-700' }
      return { label: '✅ 맥산 생산 권고', bg: 'bg-green-100', text: 'text-green-700' }
    }
    return { label: '🔍 검토필요', bg: 'bg-yellow-100', text: 'text-yellow-700' }
  }

  // ── 맥산 원가만 있을 때 ─────────────────────────────────────────────────────
  if (maeksanApplied !== null) {
    if (마진율맥산 === null || 맥산미달)
      return { label: '⚠ 철수검토', bg: 'bg-red-100', text: 'text-red-700' }
    if (마진율맥산 >= 생산유지기준)
      return { label: '✅ 맥산 생산 권고', bg: 'bg-green-100', text: 'text-green-700' }
    return { label: '🔍 검토필요', bg: 'bg-yellow-100', text: 'text-yellow-700' }
  }

  // ── FLC 원가만 있을 때 ──────────────────────────────────────────────────────
  if (flcApplied !== null) {
    if (마진율FLC === null || FLC미달)
      return { label: '⚠ 철수검토', bg: 'bg-red-100', text: 'text-red-700' }
    if (마진율FLC >= 생산유지기준)
      return { label: '📦 수입 권고', bg: 'bg-blue-100', text: 'text-blue-700' }
    return { label: '🔍 검토필요', bg: 'bg-yellow-100', text: 'text-yellow-700' }
  }

  return { label: '—', bg: 'bg-gray-100', text: 'text-gray-400' }
}

/**
 * 권고 레이블에 대응하는 마진율을 반환한다.
 * 신호등(marginSignal) 색상을 결정하는 데 사용.
 * - 수입 권고 → FLC 마진율
 * - 맥산 생산 권고/가능, 검토필요 → 맥산 마진율
 * - 철수검토 → 양쪽 중 더 나은 값 (얼마나 심각한지 표시)
 */
export function getSignalMargin(row: ProfitRow, label: string): number | null {
  if (label.includes('수입 권고')) return row.마진율FLC
  if (label.includes('맥산 생산') || label.includes('검토필요')) return row.마진율맥산
  if (label.includes('철수검토')) {
    const candidates = [row.마진율맥산, row.마진율FLC].filter((v): v is number => v !== null)
    return candidates.length ? Math.max(...candidates) : null
  }
  return null
}

/**
 * 마진율 → 신호등 이모지
 * 🟢 ≥ 20%  /  🟡 0~20%  /  🔴 < 0%
 */
export function marginSignal(v: number | null): string {
  if (v === null) return ''
  if (v < 0)  return '🔴'
  if (v < 20) return '🟡'
  return '🟢'
}
