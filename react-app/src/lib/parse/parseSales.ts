import * as XLSX from 'xlsx'

export interface SalesRow {
  code: string
  name: string
  year: string  // '23', '24', '25'
  qty:  number
}

function toYY(raw: unknown): string | null {
  const s = String(raw ?? '').trim()
  const n = parseInt(s)
  if (isNaN(n)) return null
  const str = String(n)
  if (str.length === 4) return str.slice(2)
  if (str.length === 2) return str
  return null
}

/** 전체 판매량.xlsx
 *  형식 A (연간 합산): 품목코드, 품목명, 년, 수량
 *  형식 B (판매현황 분석): 거래처명, 품목코드, 품목명, 년, 월, 수량, 공급가액 — 자동 감지 후 연간 합산
 */
export function parseSalesFile(buffer: ArrayBuffer, logs: string[]): SalesRow[] {
  const wb  = XLSX.read(buffer, { type: 'array' })
  const allRaw: Record<string, unknown>[] = []
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name]
    if (!ws || !ws['!ref']) continue
    allRaw.push(...XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null }))
  }
  logs.push(`전체 판매량: ${allRaw.length.toLocaleString()}행 로드`)

  // 형식 감지 — '월' 컬럼 또는 '거래처명' 컬럼이 있으면 판매현황 분석 파일
  const isDetailed = allRaw.length > 0 && (
    allRaw[0]['월'] !== undefined || allRaw[0]['거래처명'] !== undefined || allRaw[0]['거래처'] !== undefined
  )
  if (isDetailed) logs.push('  → 판매현황 분석 파일 감지 — 월별 데이터를 연간 합산합니다')

  const agg: Record<string, { code: string; name: string; qty: number }> = {}

  for (const row of allRaw) {
    const code = String(row['품목코드'] ?? '').trim()
    const name = String(row['품목명'] ?? row['제품명'] ?? '').trim()
    if (!name) continue

    const qty = parseInt(String(row['수량'] ?? '0')) || 0
    if (qty <= 0) continue

    const yr = toYY(row['년'] ?? row['연도'] ?? row['year'])
    if (!yr) continue

    const key = `${code}||${name}||${yr}`
    if (!agg[key]) agg[key] = { code, name, qty: 0 }
    agg[key].qty += qty
  }

  const rows: SalesRow[] = Object.entries(agg).map(([key, v]) => {
    const [,, yr] = key.split('||')
    return { code: v.code, name: v.name, year: yr, qty: v.qty }
  })

  logs.push(`  → OJC 포함 전체 ${rows.length.toLocaleString()}건 (${isDetailed ? '월별→연간 합산' : '연간 직접'})`)
  return rows
}

/** 구매관리(맥산).xlsx — 컬럼: 품목코드, 품목명, 입고일자, 수량 */
export function parseProductionFile(buffer: ArrayBuffer, logs: string[]): SalesRow[] {
  const wb  = XLSX.read(buffer, { type: 'array' })
  const ws  = wb.Sheets[wb.SheetNames[0]]
  const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })
  logs.push(`구매관리(맥산): ${raw.length.toLocaleString()}행 로드`)

  const rows: SalesRow[] = []
  for (const row of raw) {
    const code = String(row['품목코드'] ?? '').trim()
    const name = String(row['품목명']  ?? '').trim()
    if (!code || !name) continue

    const qty = parseInt(String(row['수량'] ?? '0')) || 0
    if (qty <= 0) continue

    // 입고일자: 'YYYY-MM-DD' or 'YYYYMMDD' or 'YY/MM/DD'
    const dateRaw = String(row['입고일자'] ?? '').replace(/\s*-\d+\s*$/, '').trim()
    const m = dateRaw.match(/^(\d{2,4})/)
    if (!m) continue
    const yr = m[1].length === 4 ? m[1].slice(2) : m[1]
    if (!/^\d{2}$/.test(yr)) continue

    rows.push({ code, name, year: yr, qty })
  }
  logs.push(`  → 생산입고 ${rows.length.toLocaleString()}건`)
  return rows
}
