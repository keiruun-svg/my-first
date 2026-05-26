import * as XLSX from 'xlsx'

// EMP 재고현황 또는 이카운트 수불부를 자동 감지하여 code → stock 반환
export async function parseStockFile(file: File): Promise<Map<string, number>> {
  const buf  = await file.arrayBuffer()
  const wb   = XLSX.read(buf, { type: 'array' })
  const ws   = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]

  // EMP 형식 감지 — 헤더에 "재고수량" 포함
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i] as unknown[]
    const hasCode = row.some(c => String(c ?? '').includes('상품코드') || String(c ?? '').includes('품목코드'))
    const hasQty  = row.some(c => String(c ?? '').includes('재고수량'))
    if (hasCode && hasQty) return parseEmpFormat(rows, i)
  }

  // 이카운트 수불부 형식 — "합계" 행 존재
  return parseEcountFormat(rows)
}

function parseEmpFormat(rows: unknown[][], headerRow: number): Map<string, number> {
  const result = new Map<string, number>()
  const header = (rows[headerRow] as unknown[]).map(h => String(h ?? '').trim())
  const codeIdx = header.findIndex(h => h.includes('상품코드') || h.includes('품목코드'))
  const qtyIdx  = header.findIndex(h => h.includes('재고수량'))
  if (codeIdx < 0 || qtyIdx < 0) return result
  for (const row of rows.slice(headerRow + 1)) {
    const r    = row as unknown[]
    const code = String(r[codeIdx] ?? '').trim()
    const qty  = Number(r[qtyIdx]) || 0
    if (!code) continue
    result.set(code, (result.get(code) ?? 0) + qty)
  }
  return result
}

function parseEcountFormat(rows: unknown[][]): Map<string, number> {
  const result = new Map<string, number>()
  let lastSumRow: unknown[] | null = null
  let code = ''
  for (const row of rows) {
    const r  = row as unknown[]
    const c1 = String(r[1] ?? '').trim()
    if (c1 && !code) code = c1
    const c0 = String(r[0] ?? '').trim()
    if (c0 === '합계' || c0.includes('합계')) lastSumRow = r
  }
  if (code && lastSumRow) {
    const stock = Number(lastSumRow[lastSumRow.length - 1]) || 0
    result.set(code, stock)
  }
  return result
}
