import * as XLSX from 'xlsx'

// XLSX 헤더 공백 무시 정규화 조회
function makeNorm(row: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  for (const [k, v] of Object.entries(row)) out[k.replace(/\s+/g, '')] = v
  return out
}

function numVal(v: unknown): number {
  if (typeof v === 'number') return v
  return parseFloat(String(v ?? '0').replace(/,/g, '')) || 0
}

function str(norm: Record<string, unknown>, ...keys: string[]): string {
  for (const k of keys) {
    const v = norm[k]
    if (v !== null && v !== undefined) {
      const s = String(v).trim()
      if (s) return s
    }
  }
  return ''
}

function num(norm: Record<string, unknown>, ...keys: string[]): number {
  for (const k of keys) {
    const v = norm[k]
    if (v !== null && v !== undefined && String(v).trim()) return numVal(v)
  }
  return 0
}

export interface CostEntry {
  code: string
  name: string
  생산원가: number
  표준원가: number
}

export function parseMaeksanCost(buffer: ArrayBuffer): Map<string, CostEntry> {
  const wb = XLSX.read(buffer, { type: 'array' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })
  const result = new Map<string, CostEntry>()
  for (const row of rows) {
    const n = makeNorm(row)
    const code = str(n, '품번')
    if (!code) continue
    const name = str(n, '품명')
    // 맥산원가 컬럼이 생산원가에 해당
    const 생산원가 = num(n, '맥산원가', '생산원가')
    const 표준원가 = num(n, '표준원가')
    if (생산원가 <= 0 && 표준원가 <= 0) continue
    result.set(code, { code, name, 생산원가, 표준원가 })
  }
  return result
}

export function parseFlcCost(buffer: ArrayBuffer): Map<string, CostEntry> {
  const wb = XLSX.read(buffer, { type: 'array' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })
  const result = new Map<string, CostEntry>()
  for (const row of rows) {
    const n = makeNorm(row)
    const code = str(n, '품번')
    if (!code) continue
    const name = str(n, '품명')
    const 생산원가 = num(n, '생산원가')
    const 표준원가 = num(n, '표준원가')
    if (생산원가 <= 0 && 표준원가 <= 0) continue
    result.set(code, { code, name, 생산원가, 표준원가 })
  }
  return result
}

export interface ContractItem {
  code: string
  name: string
  spec: string
  단가: number
}

export function parseContractItems(buffer: ArrayBuffer): Map<string, ContractItem> {
  const wb = XLSX.read(buffer, { type: 'array' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })
  const result = new Map<string, ContractItem>()
  for (const row of rows) {
    const n = makeNorm(row)
    const code = str(n, '품목코드', '품번')
    if (!code) continue
    const name = str(n, '품명')
    const spec = str(n, '규격')
    const 단가 = num(n, '단가')
    if (단가 <= 0) continue
    result.set(code, { code, name, spec, 단가 })
  }
  return result
}
