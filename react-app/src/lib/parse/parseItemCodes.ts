import * as XLSX from 'xlsx'

export interface ItemCode {
  code: string
  name: string
  spec: string
}

export async function parseItemCodes(file: File): Promise<ItemCode[]> {
  const buf  = await file.arrayBuffer()
  const wb   = XLSX.read(buf, { type: 'array' })
  const ws   = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][]

  // 헤더 행 탐색 — 이카운트(품목코드) 또는 EMP(상품코드) 자동 감지
  let headerIdx = -1
  let codeCol   = -1
  let nameCol   = -1
  let specCol   = -1
  for (let r = 0; r < Math.min(rows.length, 10); r++) {
    const row = rows[r] as unknown[]
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] ?? '').trim()
      if (cell === '품목코드' || cell === '상품코드') { headerIdx = r; codeCol = c }
      if (cell === '품목명'   || cell === '상품명')   { nameCol = c }
      if (cell === '규격')                             { specCol = c }
    }
    if (headerIdx === r && codeCol >= 0 && nameCol >= 0) break
  }

  if (headerIdx < 0 || codeCol < 0 || nameCol < 0) {
    throw new Error('"품목코드(상품코드)" / "품목명(상품명)" 컬럼을 찾을 수 없습니다. 이카운트 또는 EMP 품목 리스트 파일인지 확인하세요.')
  }

  const result: ItemCode[] = []
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row  = rows[r] as unknown[]
    const code = String(row[codeCol] ?? '').trim()
    const name = String(row[nameCol] ?? '').trim()
    const spec = specCol >= 0 ? String(row[specCol] ?? '').trim() : ''
    if (code) result.push({ code, name, spec })
  }
  return result
}
