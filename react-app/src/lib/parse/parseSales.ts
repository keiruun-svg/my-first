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

/** 컬럼 헤더에서 연도 추출 — "23년\n판매(EA)", "24년 판매(EA)" 등 */
function extractYearFromHeader(h: string): string | null {
  const m = h.replace(/\r\n/g, '\n').match(/^(\d{2})년[\s\S]*판매\(EA\)/)
  return m ? m[1] : null
}

/**
 * 전체 판매량.xlsx 지원 형식
 *  형식 A (연간 합산):  품목코드, 품목명, 년, 수량
 *  형식 B (월별 상세):  거래처명, 품목코드, 품목명, 년, 월, 수량, 공급가액 — 연간 합산
 *  형식 C (가로 집계):  판매현황분석 출력 파일 — 타이틀 행 + "23년\n판매(EA)" 형태 컬럼 — 자동 감지
 */
export function parseSalesFile(buffer: ArrayBuffer, logs: string[]): SalesRow[] {
  const wb = XLSX.read(buffer, { type: 'array' })
  const allRows: SalesRow[] = []
  logs.push(`시트 목록: ${wb.SheetNames.join(', ')}`)

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName]
    if (!ws || !ws['!ref']) { logs.push(`  ${sheetName}: 빈 시트`); continue }

    // 원시 배열로 읽어서 헤더 행 위치 찾기
    const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]

    // '품목명' 셀이 있는 행을 헤더 행으로 인식
    const hdrIdx = raw.findIndex(r => r.some(c => String(c ?? '').trim() === '품목명'))
    if (hdrIdx < 0) {
      // 첫 행 컬럼 샘플 로그 (진단용)
      const firstCols = (raw[0] ?? []).slice(0, 5).map(c => String(c ?? '').slice(0, 20)).join(' | ')
      logs.push(`  ${sheetName}: 헤더없음 — 1행: [${firstCols}]`)
      continue
    }

    const headers = raw[hdrIdx].map(h => String(h ?? '').trim())

    // 형식 C 판별 — 헤더에 "XX년 판매(EA)" 패턴 컬럼이 있는 경우
    const yearCols: Array<{ yr: string; idx: number }> = []
    headers.forEach((h, idx) => {
      const yr = extractYearFromHeader(h)
      if (yr) yearCols.push({ yr, idx })
    })

    if (yearCols.length > 0) {
      // 형식 C: 가로 집계형
      const nameIdx = headers.findIndex(h => h === '품목명')
      const codeIdx = headers.findIndex(h => h === '품목코드')
      let cnt = 0
      for (let i = hdrIdx + 1; i < raw.length; i++) {
        const row = raw[i]
        const name = String(row[nameIdx] ?? '').trim()
        const code = codeIdx >= 0 ? String(row[codeIdx] ?? '').trim() : ''
        if (!name) continue
        for (const { yr, idx } of yearCols) {
          const qty = parseInt(String(row[idx] ?? '0')) || 0
          if (qty <= 0) continue
          allRows.push({ code, name, year: yr, qty })
          cnt++
        }
      }
      logs.push(`  ${sheetName}: 가로집계형 감지 (${yearCols.map(c => '20' + c.yr + '년').join('/')}) → ${cnt.toLocaleString()}건`)
      continue
    }

    // 형식 A / B: 세로형 (년/연도 컬럼 + 수량 컬럼)
    const dataRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
      defval: null,
      range: hdrIdx,  // 헤더 행부터 읽기
    })

    // 시트명에서 연도 추출 — 예: "4_판매 20230101_20231231" → '23'
    const sheetYrMatch = sheetName.match(/20(\d{2})\d{4}/)
    const sheetYr = sheetYrMatch ? sheetYrMatch[1] : null

    const isDetailed = dataRows.length > 0 && (
      dataRows[0]['월'] !== undefined ||
      dataRows[0]['거래처명'] !== undefined ||
      dataRows[0]['거래처'] !== undefined
    )

    const agg: Record<string, { code: string; name: string; qty: number }> = {}
    for (const row of dataRows) {
      const code = String(row['품목코드'] ?? '').trim()
      const name = String(row['품목명'] ?? row['제품명'] ?? '').trim()
      if (!name) continue

      const qty = parseInt(String(row['수량'] ?? row['합계'] ?? '0')) || 0
      if (qty <= 0) continue

      // 년 컬럼 → 일자 앞 4자리 → 시트명 순서로 연도 추출
      const dateStr = String(row['일자'] ?? row['일'] ?? '')
      const dateYr = dateStr.length >= 4 ? toYY(dateStr.slice(0, 4)) : null
      const yr = toYY(row['년'] ?? row['연도'] ?? row['year']) ?? dateYr ?? sheetYr
      if (!yr) continue

      const key = `${code}||${name}||${yr}`
      if (!agg[key]) agg[key] = { code, name, qty: 0 }
      agg[key].qty += qty
    }

    const sheetRows = Object.entries(agg).map(([key, v]) => {
      const [,, yr] = key.split('||')
      return { code: v.code, name: v.name, year: yr, qty: v.qty }
    })

    if (sheetRows.length > 0) {
      allRows.push(...sheetRows)
      logs.push(`  ${sheetName}: ${isDetailed ? '월별→연간 합산' : '연간 직접'} ${sheetRows.length.toLocaleString()}건`)
    }
  }

  logs.push(`전체 판매량 합산: ${allRows.length.toLocaleString()}건`)
  return allRows
}

/** 구매관리(맥산).xlsx — 컬럼: 품목코드, 품목명, 입고일자, 수량 */
export function parseProductionFile(buffer: ArrayBuffer, logs: string[]): SalesRow[] {
  const wb  = XLSX.read(buffer, { type: 'array' })
  const ws  = wb.Sheets[wb.SheetNames[0]]
  const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })
  logs.push(`구매관리(맥산): ${raw.length.toLocaleString()}행 로드`)

  // 진단: 첫 행 컬럼 확인
  if (raw.length > 0) {
    const firstKeys = Object.keys(raw[0]).slice(0, 6).join(', ')
    logs.push(`  컬럼: [${firstKeys}]`)
  }

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
