import type { ErpRow } from '../parse/parseERP'
import { classifyRow } from '../parse/classify'
import type { ProductInfo } from '../parse/classify'

export interface PivotRow extends ProductInfo {
  code:    string
  name:    string
  spec:    string
  monthly: number[]  // 인덱스 0 = 1월, ..., 11 = 12월
}

/** year → 정렬된 PivotRow[] */
export type YearPivot = Map<string, PivotRow[]>

/**
 * ERP 행을 연도별·품목별로 집계합니다.
 * @param rows   parseERP() 결과
 * @param years  포함할 연도 배열 (예: ['2023','2024','2025']). 미지정시 전체 연도.
 */
export function buildPivot(rows: ErpRow[], years?: string[]): YearPivot {
  const targetYears = years ?? [...new Set(rows.map(r => r.year))].sort()
  const filtered    = rows.filter(r => targetYears.includes(r.year))

  // 연도별 그룹
  const byYear = new Map<string, ErpRow[]>()
  for (const row of filtered) {
    if (!byYear.has(row.year)) byYear.set(row.year, [])
    byYear.get(row.year)!.push(row)
  }

  const result: YearPivot = new Map()
  for (const [year, yRows] of byYear) {
    const map = new Map<string, PivotRow>()
    for (const row of yRows) {
      const key = `${row.code}||${row.name}||${row.spec}`
      if (!map.has(key)) {
        map.set(key, {
          ...classifyRow(row.name, row.spec),
          code:    row.code,
          name:    row.name,
          spec:    row.spec,
          monthly: new Array(12).fill(0),
        })
      }
      const pr = map.get(key)!
      if (row.month >= 1 && row.month <= 12)
        pr.monthly[row.month - 1] += row.qty
    }
    // 품목코드 기준 정렬
    result.set(year, [...map.values()].sort((a, b) => a.code.localeCompare(b.code)))
  }
  return result
}
