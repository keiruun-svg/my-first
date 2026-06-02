import * as XLSX from 'xlsx'

export interface MaterialStockItem {
  품명:   string
  공급사: string
  구분:   string   // '현행' | '단종'
  총재고: number
}

// 의왕_생산자재_재고현황 형식 파싱
// 헤더: 행2~3(병합), 데이터: 행4~
// D열(3)=품번, E열(4)=품명, C열(2)=구분, L열(11)=공급사, O열(14)=총재고
export function parseMaterialStock(buf: ArrayBuffer): Record<string, MaterialStockItem> {
  const wb = XLSX.read(buf, { type: 'array' })
  const sheetName = wb.SheetNames.find(s => s.includes('주간사용량')) ?? wb.SheetNames[0]
  const ws = wb.Sheets[sheetName]
  if (!ws) return {}

  const range = XLSX.utils.decode_range(ws['!ref']!)
  const result: Record<string, MaterialStockItem> = {}

  for (let r = 4; r <= range.e.r; r++) {
    const 품번  = ws[XLSX.utils.encode_cell({ r, c: 3 })]?.v?.toString().trim() ?? ''
    if (!품번) continue
    const 품명   = ws[XLSX.utils.encode_cell({ r, c: 4  })]?.v?.toString().trim() ?? ''
    const 구분   = ws[XLSX.utils.encode_cell({ r, c: 2  })]?.v?.toString().trim() ?? ''
    const 공급사 = ws[XLSX.utils.encode_cell({ r, c: 11 })]?.v?.toString().trim() ?? ''
    const 총재고 = Number(ws[XLSX.utils.encode_cell({ r, c: 14 })]?.v ?? 0)
    result[품번] = { 품명, 공급사, 구분, 총재고 }
  }

  return result
}
