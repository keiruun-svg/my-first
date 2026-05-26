import * as XLSX from 'xlsx'

export interface ItemCode {
  code: string
  name: string
}

export async function parseItemCodes(file: File): Promise<ItemCode[]> {
  const buf  = await file.arrayBuffer()
  const wb   = XLSX.read(buf, { type: 'array' })
  const ws   = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][]

  // 헤더 행 탐색: "품목코드" 포함 행
  let headerIdx = -1
  let codeCol   = -1
  let nameCol   = -1
  for (let r = 0; r < Math.min(rows.length, 10); r++) {
    const row = rows[r] as unknown[]
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] ?? '')
      if (cell === '품목코드') { headerIdx = r; codeCol = c }
      if (cell === '품목명')   { nameCol = c }
    }
    if (headerIdx === r && codeCol >= 0 && nameCol >= 0) break
  }

  if (headerIdx < 0 || codeCol < 0 || nameCol < 0) {
    throw new Error('"품목코드" / "품목명" 컬럼을 찾을 수 없습니다. 이카운트 품목 리스트 파일인지 확인하세요.')
  }

  const result: ItemCode[] = []
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row  = rows[r] as unknown[]
    const code = String(row[codeCol] ?? '').trim()
    const name = String(row[nameCol] ?? '').trim()
    if (code) result.push({ code, name })
  }
  return result
}
