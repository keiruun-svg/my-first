/**
 * 연간발주계획 검증 스크립트
 * 실행: npx tsx validate_plan.ts  (react-app 디렉토리에서)
 */
import * as fs from 'fs'
import * as path from 'path'
import { fileURLToPath } from 'url'
import { aggregateStats, buildOrderPlan } from './src/lib/step2Core'
import type { Metadata, Inventory, SalesAnalysis } from './src/lib/types'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const GAONG_FILE = path.join(__dirname, '..', '검증_가공파일.xlsx')

const buf = fs.readFileSync(GAONG_FILE)
const stats = aggregateStats(buf.buffer as ArrayBuffer)

console.log('\n========== 감지된 연도 ==========')
console.log(' ', stats.logs.join('\n  '))

// ── ① 케이블 집계 결과 ──────────────────────────────────────
console.log('\n========== ① 케이블 집계 (상위 20개) ==========')
console.log('키'.padEnd(28) + stats.years.map(y => `20${y}년 연간`.padStart(12)).join('') + '  3개년 평균')
console.log('─'.repeat(28 + stats.years.length * 12 + 12))
const cEntries = Object.entries(stats.cableStats)
  .map(([k, yd]) => ({
    k,
    annuals: stats.years.map(yr => yd[yr]?.annual ?? 0),
    avg: Math.round(stats.years.map(yr => yd[yr]?.annual ?? 0).reduce((a,b)=>a+b,0) / stats.years.length),
  }))
  .sort((a, b) => b.avg - a.avg)
  .slice(0, 20)

for (const { k, annuals, avg } of cEntries) {
  const row = k.padEnd(28) + annuals.map(v => v.toLocaleString().padStart(12)).join('') + '  ' + avg.toLocaleString().padStart(8)
  console.log(row)
}

// ── ② 하우징 집계 결과 ──────────────────────────────────────
console.log('\n========== ② 하우징 집계 (상위 20개) ==========')
console.log('키'.padEnd(28) + stats.years.map(y => `20${y}년 연간`.padStart(12)).join('') + '  3개년 평균')
console.log('─'.repeat(28 + stats.years.length * 12 + 12))
const hEntries = Object.entries(stats.housingStats)
  .map(([k, yd]) => ({
    k,
    annuals: stats.years.map(yr => yd[yr]?.annual ?? 0),
    avg: Math.round(stats.years.map(yr => yd[yr]?.annual ?? 0).reduce((a,b)=>a+b,0) / stats.years.length),
  }))
  .sort((a, b) => b.avg - a.avg)
  .slice(0, 20)

for (const { k, annuals, avg } of hEntries) {
  const row = k.padEnd(28) + annuals.map(v => v.toLocaleString().padStart(12)).join('') + '  ' + avg.toLocaleString().padStart(8)
  console.log(row)
}

// ── ③ PIGTAIL 색상 분배 검증 ────────────────────────────────
console.log('\n========== ③ PIGTAIL 색상 분배 (step2Core 로직) ==========')
const pigtailKeys = Object.keys(stats.cableStats).filter(k => k.includes('pigtail'))
if (!pigtailKeys.length) {
  console.log('  ⚠ PIGTAIL 항목 없음')
} else {
  console.log(`  발견된 PIGTAIL 키 (${pigtailKeys.length}개):`)
  for (const k of pigtailKeys) {
    const annuals = stats.years.map(yr => stats.cableStats[k][yr]?.annual ?? 0)
    console.log(`  ${k.padEnd(30)} ${annuals.map(v => v.toLocaleString().padStart(10)).join('  ')}`)
  }
  const colors = ['청','등','녹','적','황','자','갈','흑','백','회','연청','연등']
  const foundColors = [...new Set(pigtailKeys.map(k => {
    const part = k.split('pigtail-')[1] ?? ''
    return colors.includes(part) ? part : '?'
  }))]
  const missingColors = colors.filter(c => !foundColors.includes(c))
  console.log(`\n  발견된 색상: ${foundColors.join(', ')}`)
  if (missingColors.length) console.log(`  ⚠ 없는 색상: ${missingColors.join(', ')}`)
  else console.log('  ✅ 12색상 모두 존재')
}

// ── ④ buildOrderPlan 검증 ────────────────────────────────────
console.log('\n========== ④ 발주계획 계산 검증 ==========')
const emptyMeta: Metadata = { cable: {}, housing: {} }
const emptyInv: Inventory = { cable: {}, housing: {} }
const emptySales: SalesAnalysis = {}
const plan = buildOrderPlan(stats, emptyMeta, emptyInv, emptySales, 60)

// 케이블 상위 15개
const cPlan = plan.filter(r => r.type === 'cable').slice(0, 15)
console.log('\n[케이블 발주계획 상위 15개] (리드타임 기본 60일)')
console.log(
  '케이블 타입'.padEnd(28) +
  stats.years.map(y => `20${y}년`.padStart(10)).join('') +
  '  CAGR'.padStart(8) + '  제안량'.padStart(10) + '  안전재고'.padStart(10)
)
console.log('─'.repeat(28 + stats.years.length * 10 + 30))
for (const r of cPlan) {
  const key = `${r.pai}|${r.ctype}`
  const annuals = stats.years.map(yr => r.yearStats[yr]?.annual ?? 0)
  const firstYr = stats.years[0], lastYr = stats.years[stats.years.length - 1]
  const aFirst = r.yearStats[firstYr]?.annual ?? 0
  const aLast  = r.yearStats[lastYr]?.annual ?? 0
  const nYrs = stats.years.length - 1
  let cagr = 0
  if (aFirst > 0 && aLast > 0 && nYrs > 0) cagr = Math.pow(aLast / aFirst, 1 / nYrs) - 1

  const cagrStr = (cagr * 100).toFixed(1) + '%'
  const row =
    key.padEnd(28) +
    annuals.map(v => v.toLocaleString().padStart(10)).join('') +
    cagrStr.padStart(8) +
    (r.제안량 ?? 0).toLocaleString().padStart(10) +
    r.안전재고.toLocaleString().padStart(10)
  console.log(row)
}

// 하우징 상위 10개
const hPlan = plan.filter(r => r.type === 'housing').slice(0, 10)
console.log('\n[하우징 발주계획 상위 10개]')
console.log(
  '하우징 타입'.padEnd(28) +
  stats.years.map(y => `20${y}년`.padStart(10)).join('') +
  '  CAGR'.padStart(8) + '  제안량'.padStart(10) + '  안전재고'.padStart(10)
)
console.log('─'.repeat(28 + stats.years.length * 10 + 30))
for (const r of hPlan) {
  const key = `${r.pai}|${r.ctype}`
  const annuals = stats.years.map(yr => r.yearStats[yr]?.annual ?? 0)
  const firstYr = stats.years[0], lastYr = stats.years[stats.years.length - 1]
  const aFirst = r.yearStats[firstYr]?.annual ?? 0
  const aLast  = r.yearStats[lastYr]?.annual ?? 0
  const nYrs = stats.years.length - 1
  let cagr = 0
  if (aFirst > 0 && aLast > 0 && nYrs > 0) cagr = Math.pow(aLast / aFirst, 1 / nYrs) - 1

  const cagrStr = (cagr * 100).toFixed(1) + '%'
  const row =
    key.padEnd(28) +
    annuals.map(v => v.toLocaleString().padStart(10)).join('') +
    cagrStr.padStart(8) +
    (r.제안량 ?? 0).toLocaleString().padStart(10) +
    r.안전재고.toLocaleString().padStart(10)
  console.log(row)
}

// ── ⑤ 26년 부분연도 영향 분석 ──────────────────────────────
console.log('\n========== ⑤ 26년 부분연도 영향 분석 ==========')
const hasYr26 = stats.years.includes('26')
if (!hasYr26) {
  console.log('  26년 데이터 없음')
} else {
  console.log('  26년은 5월까지만 데이터 있음 → CAGR 및 제안량에 영향 있음\n')
  console.log('  케이블 타입'.padEnd(28) + '25년 연간'.padStart(12) + '26년(~5월)'.padStart(12) + '  비율(월평균×12)'.padStart(18) + '  CAGR'.padStart(8) + '  제안량'.padStart(10))
  console.log('  ' + '─'.repeat(90))
  for (const r of cPlan.slice(0, 10)) {
    const key = `${r.pai}|${r.ctype}`
    const a25 = r.yearStats['25']?.annual ?? 0
    const a26 = r.yearStats['26']?.annual ?? 0
    const months26 = r.yearStats['26'] ? r.yearStats['26'].monthly.filter(v => v > 0).length : 0
    const annualized26 = months26 > 0 ? Math.round(a26 / months26 * 12) : 0
    const yoy = a25 > 0 ? ((annualized26 - a25) / a25 * 100).toFixed(1) + '%' : 'N/A'

    const aFirst = r.yearStats[stats.years[0]]?.annual ?? 0
    const aLast  = r.yearStats['26']?.annual ?? 0
    const nYrs = stats.years.length - 1
    let cagr = 0
    if (aFirst > 0 && aLast > 0 && nYrs > 0) cagr = Math.pow(aLast / aFirst, 1 / nYrs) - 1
    cagr = Math.min(Math.max(cagr, -0.5), 1.0)

    console.log('  ' +
      key.padEnd(28) +
      a25.toLocaleString().padStart(12) +
      a26.toLocaleString().padStart(12) +
      `(연환산: ${annualized26.toLocaleString()}, ${yoy})`.padStart(18) +
      (cagr * 100).toFixed(1).padStart(6) + '%' +
      (r.제안량 ?? 0).toLocaleString().padStart(10)
    )
  }
  console.log('\n  ⚠ 26년 실적(~5월)이 CAGR 기준연도가 되면 제안량이 과소 산정될 수 있습니다.')
  console.log('  → 25년을 기준으로 제안량 계산하거나, 26년을 연환산 처리하는 것이 정확합니다.')
}

// ── ⑥ trend 변수명 검증 ──────────────────────────────────────
console.log('\n========== ⑥ trend 변수 검증 ==========')
console.log(`  활성 연도 순서: ${stats.years.map(y=>'20'+y).join(' → ')}`)
const samplePlan = plan.filter(r => r.type === 'cable')[0]
if (samplePlan) {
  console.log(`  샘플 타입: ${samplePlan.pai}|${samplePlan.ctype}`)
  console.log(`  연간 실적: ${stats.years.map(yr => `20${yr}=${(samplePlan.yearStats[yr]?.annual??0).toLocaleString()}`).join(', ')}`)
  console.log(`  trend2324 = ${samplePlan.trend2324 != null ? (samplePlan.trend2324*100).toFixed(1)+'%' : 'null'}  ← 실제 의미: ${stats.years[stats.years.length-2]}→${stats.years[stats.years.length-1]}년 증감률`)
  console.log(`  trend2425 = ${samplePlan.trend2425 != null ? (samplePlan.trend2425*100).toFixed(1)+'%' : 'null'}  ← 실제 의미: ${stats.years[stats.years.length-3]}→${stats.years[stats.years.length-2]}년 증감률`)
  console.log()
  if (stats.years.length === 4) {
    console.log('  ⚠ 연도가 4개(23/24/25/26)일 때:')
    console.log('     trend2324 = (aLast - aSecondLast) / aSecondLast = 25→26 변화율  (이름과 불일치)')
    console.log('     trend2425 = (aSecondLast - aThirdLast) / aThirdLast = 24→25 변화율  (이름과 일치)')
  }
}

console.log('\n========== 검증 완료 ==========\n')
