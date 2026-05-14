import * as XLSX from 'xlsx'

export interface ErpRow {
  year: string   // '2023', '2024', ...
  month: number  // 1–12
  code: string
  name: string
  spec: string
  qty: number
}

function parseBuyDate(val: unknown): [string, number] | null {
  if (!val) return null
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val)
    if (d) return [String(d.y), d.m]
  }
  const s = String(val).trim()
  let m: RegExpMatchArray | null
  m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/); if (m) return [m[1], parseInt(m[2])]
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);   if (m) return [m[1], parseInt(m[2])]
  m = s.match(/^(\d{4})(\d{2})(\d{2})/);           if (m) return [m[1], parseInt(m[2])]
  m = s.match(/^(\d{2})\/(\d{1,2})\/(\d{1,2})/);  if (m) return ['20' + m[1], parseInt(m[2])]
  return null
}

type ColFormat = 'ojc' | 'purchase'

function detectColFormat(ws: XLSX.WorkSheet): ColFormat {
  const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]
  const h0 = String(raw[1]?.[0] ?? '').trim()
  return ['거래처', '구매처', '공급처', '납품처'].some(kw => h0.includes(kw)) ? 'ojc' : 'purchase'
}

function parseSheet(ws: XLSX.WorkSheet, fmt: ColFormat, logs: string[]): ErpRow[] {
  const raw    = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null, raw: false }) as unknown[][]
  const rawNum = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null, raw: true  }) as unknown[][]
  const rows: ErpRow[] = []
  let skippedNoKey = 0, skippedDate = 0

  for (let i = 2; i < raw.length; i++) {
    const cells    = raw[i]    as unknown[]
    const cellsNum = rawNum[i] as unknown[]
    if (!cells || !cellsNum) { skippedNoKey++; continue }

    let buyNo: unknown, code: unknown, nameRaw: unknown, specRaw: unknown, qty: unknown
    if (fmt === 'ojc') {
      buyNo = cellsNum[2]; code = cells[3]; nameRaw = cells[4]; specRaw = cells[6]; qty = cells[7]
    } else {
      buyNo = cellsNum[0]; code = cells[4]; nameRaw = cells[5]; specRaw = null;     qty = cells[6]
    }

    if (!buyNo || !nameRaw || !qty) { skippedNoKey++; continue }
    const parsed = parseBuyDate(buyNo)
    if (!parsed) { skippedDate++; continue }
    const [year, month] = parsed

    const s  = String(nameRaw).trim().replace(/\s*외\s*\d+건\s*$/, '')
    const bm = s.match(/^(.+?)\s*\[(.+)\]\s*$/)
    let name: string, spec: string
    if (bm) { name = bm[1].trim(); spec = bm[2].trim() }
    else     { name = s;           spec = specRaw ? String(specRaw).trim() : '' }

    const q = parseInt(String(qty))
    if (isNaN(q) || q <= 0) continue

    rows.push({ year, month, code: code ? String(code).trim() : '', name, spec, qty: q })
  }

  if (!rows.length) {
    logs.push(`[진단] 필수값 없음: ${skippedNoKey}행 / 날짜 불일치: ${skippedDate}행`)
    logs.push(`[진단] 샘플: ${JSON.stringify(rawNum.slice(2, 4).map(r => (r as unknown[]).slice(0, 8)))}`)
  }
  return rows
}

export function parseERP(buffer: ArrayBuffer, logs: string[]): ErpRow[] {
  const wb     = XLSX.read(buffer, { type: 'array', cellDates: false, raw: false })
  const sheets = wb.SheetNames

  if (sheets.includes('구매조회')) {
    logs.push('구매조회 형식 감지')
    return parseSheet(wb.Sheets['구매조회'], 'ojc', logs)
  }
  if (sheets.includes('구매현황')) {
    const ws  = wb.Sheets['구매현황']
    const fmt = detectColFormat(ws)
    logs.push(fmt === 'ojc' ? '구매현황(맥산납품) 형식 감지' : '구매현황 형식 감지')
    return parseSheet(ws, fmt, logs)
  }
  if (sheets.includes('생산현황')) {
    logs.push('생산현황 형식 감지')
    return parseSheet(wb.Sheets['생산현황'], 'ojc', logs)
  }
  throw new Error(`지원하지 않는 파일 형식입니다. 시트: ${sheets.join(', ')}`)
}

/** 파일에서 연도 목록만 빠르게 추출 — UI 연도 선택용 */
export function detectYears(buffer: ArrayBuffer): string[] {
  const rows = parseERP(buffer, [])
  return [...new Set(rows.map(r => r.year))].sort()
}
