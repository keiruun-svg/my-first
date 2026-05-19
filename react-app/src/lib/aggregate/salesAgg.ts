import { classifyOjc } from '../ojcFilter'
import { deriveKind, deriveCore } from '../parse/classify'
import type { SalesRow } from '../parse/parseSales'

export interface SalesTypeEntry {
  sales:      number  // 전체 판매 수량 (EA)
  production: number  // 맥산 생산 수량 (EA)
  imported:   number  // 수입 완제품 (= sales - production)
  ratio:      number  // 생산비중 (production / sales), sales=0 이면 0
}

export interface SalesProductEntry {
  code:   string
  name:   string
  kind:   string
  byYear: Record<string, { sales: number; production: number }>
}

export interface SalesAggResult {
  /** 타입(kind)별 연도별 집계 — 발주 예측(B안) 핵심 */
  byType:         Record<string, Record<string, SalesTypeEntry>>
  /** 품목별 연도별 집계 — 수익률 분석용 */
  byProduct:      SalesProductEntry[]
  /** 생산량 기반 CAGR per kind */
  productionCagr: Record<string, number>
  /** 판매량 기반 CAGR per kind (시장 성장률) */
  salesCagr:      Record<string, number>
  years:          string[]
  logs:           string[]
}

// ── B안 예측량 계산 ─────────────────────────────────────────────
// 26년 예측 생산 수량 = 25년 판매 × (1 + 판매CAGR) × 생산비중_25년
export function forecastProduction(
  result: SalesAggResult,
  _targetYear: string,
  baseYear: string,
): Record<string, number> {
  const forecast: Record<string, number> = {}
  for (const [kind, byYear] of Object.entries(result.byType)) {
    const base = byYear[baseYear]
    if (!base || base.sales === 0) continue
    const cagr = result.salesCagr[kind] ?? 0
    const ratio = base.ratio
    forecast[kind] = Math.round(base.sales * (1 + cagr) * ratio)
  }
  return forecast
}

// ── 내부 헬퍼 ───────────────────────────────────────────────────
function kindOf(name: string): string {
  const core = deriveCore(name)
  return deriveKind(name, '', core)
}

function calcCagr(vals: number[]): number {
  // 첫 비-제로와 마지막 비-제로 사이의 CAGR
  const firstIdx = vals.findIndex(v => v > 0)
  const lastIdx  = vals.length - 1 - [...vals].reverse().findIndex(v => v > 0)
  if (firstIdx < 0 || lastIdx <= firstIdx) return 0
  const n = lastIdx - firstIdx
  return Math.pow(vals[lastIdx] / vals[firstIdx], 1 / n) - 1
}

// ── 메인 집계 ───────────────────────────────────────────────────
export function aggregateSales(
  salesRows: SalesRow[],
  prodRows:  SalesRow[],
  logs:      string[],
): SalesAggResult {
  // OJC 제품만 필터링
  const ojcSales = salesRows.filter(r => classifyOjc(r.name) !== null)
  const ojcProd  = prodRows .filter(r => classifyOjc(r.name) !== null)
  logs.push(`OJC 판매: ${ojcSales.length.toLocaleString()}건 / OJC 생산(맥산): ${ojcProd.length.toLocaleString()}건`)

  // 연도 수집 (정렬)
  const yearSet = new Set<string>([...ojcSales, ...ojcProd].map(r => r.year))
  const years   = [...yearSet].sort()

  // ── 타입별 집계 ────────────────────────────────────────────
  const salesByKind: Record<string, Record<string, number>> = {}
  const prodByKind:  Record<string, Record<string, number>> = {}

  const addTo = (map: typeof salesByKind, row: SalesRow) => {
    const kind = kindOf(row.name)
    if (!map[kind])           map[kind] = {}
    if (!map[kind][row.year]) map[kind][row.year] = 0
    map[kind][row.year] += row.qty
  }

  ojcSales.forEach(r => addTo(salesByKind, r))
  ojcProd .forEach(r => addTo(prodByKind,  r))

  const allKinds = [...new Set([...Object.keys(salesByKind), ...Object.keys(prodByKind)])].sort()

  const byType:         SalesAggResult['byType']         = {}
  const productionCagr: SalesAggResult['productionCagr'] = {}
  const salesCagr:      SalesAggResult['salesCagr']      = {}

  for (const kind of allKinds) {
    byType[kind] = {}
    const prodVals:  number[] = []
    const salesVals: number[] = []

    for (const yr of years) {
      const sales      = salesByKind[kind]?.[yr] ?? 0
      const rawProd    = prodByKind [kind]?.[yr] ?? 0
      // 생산이 판매를 초과하는 경우 판매량으로 cap (데이터 불일치 보정)
      const production = rawProd > sales && sales > 0 ? sales : rawProd
      const imported   = Math.max(sales - production, 0)
      const ratio      = sales > 0 ? production / sales : 0
      byType[kind][yr] = { sales, production, imported, ratio }
      prodVals .push(production)
      salesVals.push(sales)
    }

    productionCagr[kind] = calcCagr(prodVals)
    salesCagr     [kind] = calcCagr(salesVals)
  }

  // ── 품목별 집계 (수익률 분석용) ──────────────────────────
  const productMap: Record<string, SalesProductEntry> = {}

  for (const r of ojcSales) {
    const key = r.code || r.name
    if (!productMap[key]) {
      productMap[key] = { code: r.code, name: r.name, kind: kindOf(r.name), byYear: {} }
    }
    const e = productMap[key]
    if (!e.byYear[r.year]) e.byYear[r.year] = { sales: 0, production: 0 }
    e.byYear[r.year].sales += r.qty
  }

  // 생산 파일에서도 같은 코드가 있으면 production 채움
  for (const r of ojcProd) {
    const key = r.code || r.name
    if (!productMap[key]) {
      productMap[key] = { code: r.code, name: r.name, kind: kindOf(r.name), byYear: {} }
    }
    const e = productMap[key]
    if (!e.byYear[r.year]) e.byYear[r.year] = { sales: 0, production: 0 }
    e.byYear[r.year].production += r.qty
  }

  const byProduct = Object.values(productMap).sort((a, b) => a.kind.localeCompare(b.kind) || a.name.localeCompare(b.name))

  // 로그
  logs.push(`타입 분류: ${allKinds.length}종 — ${allKinds.join(', ')}`)
  logs.push(`품목 수: ${byProduct.length.toLocaleString()}개`)
  for (const kind of allKinds) {
    const latest = years[years.length - 1]
    const e = byType[kind][latest]
    if (!e || e.sales === 0) continue
    const pct = Math.round(e.ratio * 100)
    const cagr = Math.round((salesCagr[kind] ?? 0) * 100)
    logs.push(`  ${kind.padEnd(12)} ${latest}년 판매 ${e.sales.toLocaleString()}ea | 생산비중 ${pct}% | 판매CAGR ${cagr > 0 ? '+' : ''}${cagr}%`)
  }

  return { byType, byProduct, productionCagr, salesCagr, years, logs }
}
