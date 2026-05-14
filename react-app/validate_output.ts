/**
 * 가공파일 검증 스크립트
 * 실행: npx tsx validate_output.ts
 */
import * as fs from 'fs'
import * as path from 'path'
import { fileURLToPath } from 'url'
import * as XLSX from 'xlsx'
import { parseERP } from './src/lib/parse/parseERP'
import { buildPivot } from './src/lib/aggregate/pivot'
import { writeGaong } from './src/lib/output/writeGaong'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const INPUT_FILE = path.join(__dirname, '..', '의왕 생산 내역 25년.xlsx')
const OUTPUT_FILE = path.join(__dirname, '..', '검증_가공파일.xlsx')

// ── 1. 변환 실행 ────────────────────────────────────────────
console.log('\n========== STEP 1: 변환 실행 ==========')
const inputBuf = fs.readFileSync(INPUT_FILE)
const logs: string[] = []
const rows   = parseERP(inputBuf.buffer as ArrayBuffer, logs)
const pivot  = buildPivot(rows)
const outArr = writeGaong(pivot)
logs.forEach(l => console.log(' ', l))

fs.writeFileSync(OUTPUT_FILE, Buffer.from(outArr))
console.log(`\n✅ 가공파일 저장: ${OUTPUT_FILE}`)

// ── 2. 출력 파일 로드 ────────────────────────────────────────
const wb = XLSX.read(new Uint8Array(outArr), { type: 'array' })
console.log('\n========== 시트 목록 ==========')
wb.SheetNames.forEach(s => console.log(' ', s))

// ── 3. 케이블 타입 분류 검증 ─────────────────────────────────
console.log('\n========== ① 케이블 타입 분류 ==========')
const cableSheets = wb.SheetNames.filter(s => s.includes('케이블'))
for (const sName of cableSheets) {
  const rows = XLSX.utils.sheet_to_json<string[]>(wb.Sheets[sName], { header: 1, defval: '' }) as string[][]
  const data = rows.slice(1).filter(r => r[0] || r[1])
  const types = new Map<string, number>()
  for (const r of data) {
    const k = `${r[3]}|${r[4]}` // 케이블종류|파이
    types.set(k, (types.get(k) ?? 0) + 1)
  }
  console.log(`\n[${sName}] 품목 수: ${data.length}`)
  const sorted = [...types.entries()].sort((a,b) => b[1]-a[1]).slice(0, 15)
  sorted.forEach(([k,n]) => console.log(`  ${String(n).padStart(3)}건 │ ${k}`))
}

// ── 4. 월별 수량 합계 검증 ───────────────────────────────────
console.log('\n========== ② 월별 수량 합계 ==========')
for (const sName of cableSheets) {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sName], { header: 1, defval: 0 }) as unknown[][]
  const data = rows.slice(1).filter(r => (r as string[])[0] || (r as string[])[1])
  const monthly = new Array(12).fill(0)
  let annualTotal = 0
  for (const r of data) {
    for (let m = 0; m < 12; m++) {
      const v = Number(r[9 + m]) || 0
      monthly[m] += v
      annualTotal += v
    }
  }
  const yr = sName.replace(/[^0-9]/g, '').slice(0, 2)
  console.log(`\n[${sName}] 연간 총합: ${annualTotal.toLocaleString()}m`)
  const labels = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']
  monthly.forEach((v,i) => process.stdout.write(`  ${labels[i]}: ${String(v.toLocaleString()).padStart(8)}`+(i%4===3?'\n':'')))
}

// ── 5. 하우징 계산 검증 ─────────────────────────────────────
console.log('\n\n========== ③ 하우징 계산 검증 ==========')
const housingSheets = wb.SheetNames.filter(s => s.includes('하우징'))
for (const sName of housingSheets.slice(-1)) { // 최신 연도만
  const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sName], { header: 1, defval: 0 }) as unknown[][]
  const data = rows.slice(1).filter(r => (r as string[])[0] || (r as string[])[1])
  console.log(`\n[${sName}] 상위 10개 품목 하우징 합계:`)
  console.log('  품목명'.padEnd(45) + '타입1'.padEnd(10) + '코어' + '  검증' + '  계산수량')
  console.log('  ' + '─'.repeat(80))
  let shown = 0
  for (const r of data) {
    if (shown >= 10) break
    const arr = r as unknown[]
    const 품목명 = String(arr[1] ?? '').slice(0, 40)
    const 타입1 = String(arr[7] ?? '')
    const 코어 = Number(arr[5]) || 1
    const 검증 = Number(arr[47]) || 0
    const 계산수량 = Number(arr[48]) || 0
    const 월합 = Array.from({length:12}, (_,i) => Number(arr[9+i])||0).reduce((a,b)=>a+b,0)
    if (월합 === 0) continue
    const expected = 코어 === 1 ? 월합 * 2 : 월합 * 코어 * 2
    const match = Math.abs(검증 - expected) < 1 ? '✅' : '❌'
    console.log(`  ${품목명.padEnd(45)}${타입1.padEnd(10)}${String(코어).padStart(3)}  ${match}${String(검증).padStart(7)}  ${String(계산수량).padStart(7)}`)
    shown++
  }
}

// ── 6. PIGTAIL 색상 분배 검증 ────────────────────────────────
console.log('\n========== ④ PIGTAIL 색상 분배 ==========')
const COLORS = ['청','등','녹','적','황','자','갈','흑','백','회','연청','연등']
for (const sName of cableSheets) {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sName], { header: 1, defval: '' }) as unknown[][]
  const data = rows.slice(1).filter(r => (r as string[])[0] || (r as string[])[1])
  const pigtails = data.filter(r => String((r as string[])[3]).toLowerCase().includes('pigtail'))
  if (!pigtails.length) continue
  console.log(`\n[${sName}] PIGTAIL 항목 ${pigtails.length}개`)
  const colorGroups = new Map<string, number>()
  for (const r of pigtails) {
    const 품명 = String((r as string[])[1])
    const annual = Array.from({length:12}, (_,i) => Number((r as unknown[])[9+i])||0).reduce((a,b)=>a+b,0)
    if (annual === 0) continue
    colorGroups.set(품명.slice(-3), (colorGroups.get(품명.slice(-3)) ?? 0) + 1)
    if (colorGroups.size <= 12) console.log(`  ${품명.slice(0,50).padEnd(52)} ${annual.toLocaleString()}m`)
  }
  const foundColors = [...new Set(pigtails.map(r => {
    const n = String((r as string[])[1])
    return COLORS.find(c => n.endsWith(c)) ?? '?'
  }))]
  console.log(`  발견된 색상: ${foundColors.join(', ')}`)
  const missing = COLORS.filter(c => !foundColors.includes(c))
  if (missing.length) console.log(`  ⚠ 없는 색상: ${missing.join(', ')} (해당 연도 사용량 0)`)
  else console.log(`  ✅ 12색상 모두 확인`)
}

console.log('\n========== 검증 완료 ==========')
console.log(`📂 Excel 파일 열기: ${OUTPUT_FILE}\n`)
