import * as XLSX from 'xlsx'
import type { Metadata, Inventory } from './types'

const PAI_ORDER: Record<string, number> = { '2.0mm': 0, '3.0mm': 1, '0.9mm': 2 }

export interface Step3Row {
  type:         'cable' | 'housing'
  key:          string   // metadata lookup key e.g. "2.0mm|A1-SP"
  label:        string   // display label
  pai:          string
  unit:         string   // 'm' | '개'
  byYear:       Record<string, { annual: number; peak: number }>
  years:        string[] // sorted years that have data
  latestAnnual: number
  latestPeak:   number
  품번:          string
  품명:          string
  구매처:        string
  리드타임:      number  // days
  안전재고:      number  // round(latestPeak × LT / 30)
  현재고:        number
  기발주:        number
  발주필요량:    number  // latestAnnual + 안전재고 - 현재고 - 기발주, min 0
}

function ltDays(lt: unknown, def: number): number {
  const m = String(lt ?? '').match(/(\d+)/)
  return m ? parseInt(m[1]) : def
}

function mkRow(type: 'cable' | 'housing', key: string, label: string, pai: string, unit: string): Step3Row {
  return { type, key, label, pai, unit, byYear: {}, years: [], latestAnnual: 0, latestPeak: 0, 품번: '', 품명: '', 구매처: '', 리드타임: 60, 안전재고: 0, 현재고: 0, 기발주: 0, 발주필요량: 0 }
}

export function buildStep3Plan(
  fileBuffer: ArrayBuffer,
  metadata:  Metadata,
  inventory: Inventory,
  ltDefault: number,
): { rows: Step3Row[]; years: string[]; logs: string[] } {
  const logs: string[] = []
  const wb    = XLSX.read(fileBuffer, { type: 'array' })
  const names = new Set(wb.SheetNames)

  // 감지: {YY}년_케이블_집계
  const years = wb.SheetNames
    .map(s => s.match(/^(\d{2})년_케이블_집계/)?.[1])
    .filter(Boolean)
    .sort() as string[]

  if (!years.length) {
    logs.push('⚠ 집계 시트 없음 — STEP 1을 최신 버전으로 다시 실행하세요.')
    return { rows: [], years: [], logs }
  }
  logs.push(`감지된 연도: ${years.map(y => '20'+y+'년').join(', ')}`)

  const cableMap:   Record<string, Step3Row> = {}
  const housingMap: Record<string, Step3Row> = {}

  for (const yr of years) {
    // ── 케이블 집계 ─────────────────────────────────────────
    // cols: 0=타입, 1=파이, 2=단위, 3~14=월별, 15=연간합계, 16=월최대, 17=최대발생월
    const cSheet = `${yr}년_케이블_집계`
    if (names.has(cSheet)) {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[cSheet], { header: 1, defval: null }) as unknown[][]
      for (const row of rows.slice(1)) {
        const label  = String(row[0] ?? '').trim()
        const pai    = String(row[1] ?? '').trim()
        const unit   = String(row[2] ?? '').trim()
        const annual = Number(row[15]) || 0
        const peak   = Number(row[16]) || 0
        if (!label || !pai || annual <= 0) continue
        const key = `${pai}|${label}`
        if (!cableMap[key]) cableMap[key] = mkRow('cable', key, label, pai, unit)
        cableMap[key].byYear[yr] = { annual, peak }
      }
    }

    // ── 하우징 집계 ─────────────────────────────────────────
    // cols: 0=타입명(e.g. "2.0mm - LC/PC 청색"), 1~12=월별, 13=연간합계, 14=월최대
    const hSheet = `${yr}년_하우징_집계`
    if (names.has(hSheet)) {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[hSheet], { header: 1, defval: null }) as unknown[][]
      for (const row of rows.slice(1)) {
        const fullLabel = String(row[0] ?? '').trim()
        const annual    = Number(row[13]) || 0
        const peak      = Number(row[14]) || 0
        if (!fullLabel || annual <= 0) continue
        // "2.0mm - LC/PC 청색" → pai="2.0mm", type="LC/PC 청색" → key="2.0mm|LC/PC 청색"
        const sep  = fullLabel.indexOf(' - ')
        const pai  = sep > 0 ? fullLabel.slice(0, sep) : ''
        const type = sep > 0 ? fullLabel.slice(sep + 3) : fullLabel
        if (!pai) continue
        const key = `${pai}|${type}`
        if (!housingMap[key]) housingMap[key] = mkRow('housing', key, fullLabel, pai, 'EA')
        housingMap[key].byYear[yr] = { annual, peak }
      }
    }
  }

  // ── 메타/재고 연결 + 계산 ────────────────────────────────
  const latestYr = years[years.length - 1]
  const allRows: Step3Row[] = []

  for (const row of Object.values(cableMap)) {
    row.years        = years.filter(yr => row.byYear[yr])
    row.latestAnnual = row.byYear[latestYr]?.annual ?? 0
    row.latestPeak   = row.byYear[latestYr]?.peak   ?? 0
    const meta = metadata.cable?.[row.key] as Record<string, unknown> | undefined ?? {}
    const inv  = inventory.cable?.[row.key] as Record<string, number> | undefined ?? {}
    row.품번     = String(meta['품번'] ?? '')
    row.품명     = String(meta['품명'] ?? '')
    row.구매처   = String(meta['구매처'] ?? '')
    row.리드타임 = ltDays(meta['리드타임'], ltDefault)
    row.안전재고 = Math.round(row.latestPeak * row.리드타임 / 30)
    row.현재고   = Number(inv['현재고'] ?? 0)
    row.기발주   = 0
    row.발주필요량 = Math.max(0, row.latestAnnual + row.안전재고 - row.현재고)
    allRows.push(row)
  }

  for (const row of Object.values(housingMap)) {
    row.years        = years.filter(yr => row.byYear[yr])
    row.latestAnnual = row.byYear[latestYr]?.annual ?? 0
    row.latestPeak   = row.byYear[latestYr]?.peak   ?? 0
    const metaRaw = metadata.housing?.[row.key]
    const meta = (metaRaw ? (Array.isArray(metaRaw) ? metaRaw[0] : metaRaw) : {}) as Record<string, unknown>
    const invRaw = inventory.housing?.[row.key]
    const inv  = (invRaw ? (Array.isArray(invRaw) ? invRaw[0] : invRaw) : {}) as Record<string, number>
    row.품번     = String(meta['품번'] ?? '')
    row.품명     = String(meta['품명'] ?? '')
    row.구매처   = String(meta['구매처'] ?? '')
    row.리드타임 = ltDays(meta['리드타임'], ltDefault)
    row.안전재고 = Math.round(row.latestPeak * row.리드타임 / 30)
    row.현재고   = Number(inv['현재고'] ?? 0)
    row.기발주   = Number(inv['기발주'] ?? 0)
    row.발주필요량 = Math.max(0, row.latestAnnual + row.안전재고 - row.현재고 - row.기발주)
    allRows.push(row)
  }

  // ── 정렬: 케이블 우선, 파이 순, 타입 알파벳 ──────────────
  allRows.sort((a, b) => {
    if (a.type !== b.type) return a.type === 'cable' ? -1 : 1
    return (PAI_ORDER[a.pai] ?? 9) - (PAI_ORDER[b.pai] ?? 9) || a.label.localeCompare(b.label)
  })

  const nC = allRows.filter(r => r.type === 'cable').length
  const nH = allRows.filter(r => r.type === 'housing').length
  const nMissing = allRows.filter(r => !r.품번).length
  logs.push(`집계 완료 — 케이블 ${nC}타입 / 하우징 ${nH}타입`)
  if (nMissing) logs.push(`⚠ 품번 미등록 ${nMissing}건 — 품번 관리 탭에서 입력 필요`)

  return { rows: allRows, years, logs }
}
