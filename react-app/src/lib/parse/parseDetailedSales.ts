import * as XLSX from 'xlsx'

export interface DetailedSalesRow {
  customer: string   // 거래처명
  code:     string   // 품목코드
  name:     string   // 품목명
  year:     string   // 'YY' e.g. '23'
  month:    string   // '01'..'12'
  qty:      number
}

function toYY(raw: unknown): string {
  const s = String(raw ?? '').trim()
  if (/^\d{4}$/.test(s)) return s.slice(2)  // 2023 → 23
  if (/^\d{2}$/.test(s)) return s
  return ''
}

function toMM(raw: unknown): string {
  const n = parseInt(String(raw ?? '0'))
  if (n < 1 || n > 12) return ''
  return String(n).padStart(2, '0')
}

// 날짜 문자열에서 year/month 추출 (YYYY-MM-DD, YYYYMMDD, YY/MM/DD 등)
function dateToYearMonth(raw: unknown): { year: string; month: string } | null {
  const s = String(raw ?? '').trim()
  const m1 = s.match(/^(\d{4})[/-](\d{1,2})/)
  if (m1) return { year: m1[1].slice(2), month: String(parseInt(m1[2])).padStart(2, '0') }
  const m2 = s.match(/^(\d{6})/)
  if (m2) return { year: m2[1].slice(0, 4).slice(2), month: m2[1].slice(4, 6) }
  return null
}

function findCol(row: Record<string, unknown>, ...keys: string[]): unknown {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null) return row[k]
  }
  return null
}

export function parseDetailedSalesFile(
  buffer: ArrayBuffer,
  logs: string[],
): DetailedSalesRow[] {
  const wb = XLSX.read(buffer, { type: 'array' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })

  const rows: DetailedSalesRow[] = []
  let skipped = 0

  for (const row of raw) {
    const customer = String(
      findCol(row, '거래처명', '거래처', '고객사', '고객사명', '업체명') ?? ''
    ).trim()
    const code = String(findCol(row, '품목코드', '코드') ?? '').trim()
    const name = String(findCol(row, '품목명', '제품명', '상품명') ?? '').trim()
    const qty  = parseInt(String(findCol(row, '수량', 'qty', '판매수량', '출고수량') ?? '0')) || 0

    if (!name || qty <= 0) { skipped++; continue }

    // year + month: 직접 컬럼 우선, 날짜 컬럼으로 대체
    let year = toYY(findCol(row, '년', '연도', 'year'))
    let month = toMM(findCol(row, '월', 'month'))

    if (!year || !month) {
      const fromDate = dateToYearMonth(findCol(row, '날짜', '일자', '출고일자', '판매일자', '거래일자'))
      if (fromDate) { year = fromDate.year; month = fromDate.month }
    }

    if (!year || !month) { skipped++; continue }

    rows.push({ customer, code, name, year, month, qty })
  }

  const years = [...new Set(rows.map(r => r.year))].sort()
  const customers = [...new Set(rows.map(r => r.customer).filter(Boolean))]
  logs.push(`판매 데이터: ${rows.length.toLocaleString()}행 파싱 / ${skipped}행 제외`)
  logs.push(`연도: ${years.map(y => '20' + y + '년').join(', ')} | 거래처: ${customers.length}개`)

  return rows
}
