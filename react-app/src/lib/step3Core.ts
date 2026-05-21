import * as XLSX from 'xlsx'
import type { Metadata, Inventory } from './types'

const PAI_ORDER: Record<string, number> = { '2.0mm': 0, '3.0mm': 1, '0.9mm': 2 }

export interface CodeCableEntry {
  key:    string  // pai|label key matching Step3Row
  length: number  // meters per unit
}

export interface Step3Row {
  type:           'cable' | 'housing'
  key:            string
  label:          string
  pai:            string
  unit:           string
  byYear:         Record<string, { annual: number; peak: number; monthly: number[] }>
  years:          string[]
  latestAnnual:   number
  latestPeak:     number
  forecastAnnual: number  // latestAnnual × (1 + CAGR); equals latestAnnual if no CAGR applied
  appliedCagr:    number  // 0 if not applied
  품번:            string
  품명:            string
  구매처:          string
  리드타임:        number
  안전재고:        number
  현재고:          number
  기발주:          number
  발주필요량:      number
  isSubRow:       boolean  // 다중 부품 하우징 타입의 2번째+ 행
}

// kind + core → 집계 레이블 (writeGaong.ts kindLabel과 동일)
function kindToLabel(kind: string, core: number): string {
  const sd = core === 1 ? 'SP' : core === 2 ? 'DP' : `${core}C`
  const map: Record<string, string> = {
    'a1':'A1','b3':'B3','om1':'OM1','om3':'OM3',
    'a1-청':'A1_청','a1-녹':'A1_녹','a1-적':'A1_적','a1-자':'A1_자',
  }
  if (map[kind])  return `${map[kind]}-${sd}`
  if (kind === 'drop') return `DROP-${sd}`
  if (kind === 'pigtail' || kind === 'om1-pigtail') return ''  // 수요분석 제외
  if (kind === 'a2') return 'Optical cable'
  return kind.toUpperCase()
}

// 집계 시트 레이블("A1-SP", "B3-DP", ...) → salesAgg kind 키
function labelToKind(label: string): string {
  if (/^A1_청/i.test(label))         return 'a1-청'
  if (/^A1_녹/i.test(label))         return 'a1-녹'
  if (/^A1_적/i.test(label))         return 'a1-적'
  if (/^A1_자/i.test(label))         return 'a1-자'
  if (/^A1/i.test(label))            return 'a1'
  if (/^B3/i.test(label))            return 'b3'
  if (/^OM3/i.test(label))           return 'om3'
  if (/^OM1/i.test(label))           return 'om1'
  if (/^DROP/i.test(label))          return 'drop'
  if (/^PIGTAIL/i.test(label))       return 'pigtail'
  if (/^Optical cable/i.test(label)) return 'a2'
  return ''
}

// STEP 2 Excel(판매분석.xlsx)의 "타입별_분석" 시트에서 kind → 판매CAGR 맵 추출
export function parseSalesAggExcel(buffer: ArrayBuffer): Record<string, number> {
  try {
    const wb = XLSX.read(buffer, { type: 'array' })
    const ws = wb.Sheets['타입별_분석']
    if (!ws) return {}
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]
    if (rows.length < 2) return {}
    const header  = rows[0] as string[]
    const cagrIdx = header.indexOf('판매CAGR')
    if (cagrIdx < 0) return {}
    const result: Record<string, number> = {}
    for (const row of rows.slice(1)) {
      const kind = String(row[0] ?? '').trim()
      const cagr = Number(row[cagrIdx]) || 0
      if (kind) result[kind] = cagr
    }
    return result
  } catch { return {} }
}

function ltDays(lt: unknown, def: number): number {
  const m = String(lt ?? '').match(/(\d+)/)
  return m ? parseInt(m[1]) : def
}

function mkRow(type: 'cable' | 'housing', key: string, label: string, pai: string, unit: string): Step3Row {
  return { type, key, label, pai, unit, byYear: {}, years: [], latestAnnual: 0, latestPeak: 0, forecastAnnual: 0, appliedCagr: 0, 품번: '', 품명: '', 구매처: '', 리드타임: 60, 안전재고: 0, 현재고: 0, 기발주: 0, 발주필요량: 0, isSubRow: false }
}

export function buildStep3Plan(
  fileBuffer: ArrayBuffer,
  metadata:  Metadata,
  inventory: Inventory,
  ltDefault: number,
  safetyK = 1.5,
  kindCagr?: Record<string, number>,
): { rows: Step3Row[]; years: string[]; logs: string[]; codeToCable: Record<string, CodeCableEntry> } {
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
    return { rows: [], years: [], logs, codeToCable: {} }
  }
  logs.push(`감지된 연도: ${years.map(y => '20'+y+'년').join(', ')}`)

  // ── 품목코드 → 자재 키 매핑: 코드매핑 시트 우선, 없으면 원본 케이블 시트 ──
  // 코드매핑 cols: 0=품목코드, 1=케이블종류, 2=파이, 3=코어수, 4=길이(m), 5=집계키
  // 원본 케이블 cols: 0=품목코드, 3=케이블종류, 4=파이, 5=코어수, 6=케이블길이
  const codeToCable: Record<string, CodeCableEntry> = {}
  if (names.has('코드매핑')) {
    const mapRows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets['코드매핑'], { header: 1, defval: null }) as unknown[][]
    for (const row of mapRows.slice(1)) {
      const code = String(row[0] ?? '').trim()
      const key  = String(row[5] ?? '').trim()
      const len  = Number(row[4]) || 0
      if (!code || !key || len <= 0) continue
      codeToCable[code] = { key, length: len }
    }
  } else {
    for (const yr of years) {
      const rawSheet = `${yr}년_케이블`
      if (names.has(rawSheet)) {
        const rawRows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[rawSheet], { header: 1, defval: null }) as unknown[][]
        for (const row of rawRows.slice(1)) {
          const code = String(row[0] ?? '').trim()
          if (!code || codeToCable[code]) continue
          const kind = String(row[3] ?? '').trim()
          const pai  = String(row[4] ?? '').trim()
          const core = Number(row[5]) || 1
          const len  = Number(row[6]) || 0
          if (!kind || !pai || len <= 0) continue
          const label = kindToLabel(kind, core)
          if (!label) continue
          codeToCable[code] = { key: `${pai}|${label}`, length: len }
        }
      }
    }
  }

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
        const monthly = Array.from({ length: 12 }, (_, i) => Number(row[3 + i]) || 0)
        const annual = Number(row[15]) || 0
        const peak   = Number(row[16]) || 0
        if (!label || !pai || annual <= 0) continue
        const key = `${pai}|${label}`
        if (!cableMap[key]) cableMap[key] = mkRow('cable', key, label, pai, unit)
        cableMap[key].byYear[yr] = { annual, peak, monthly }
      }
    }

    // ── 하우징 집계 ─────────────────────────────────────────
    // 신형: col[0]=타입, col[1]=파이, col[2]=단위, col[3~14]=월별, col[15]=연간합계, col[16]=월최대
    // 구형: col[0]=전체명("2.0mm - LC/PC 청색"), col[1~12]=월별, col[13]=연간합계, col[14]=월최대
    const hSheet = `${yr}년_하우징_집계`
    if (names.has(hSheet)) {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[hSheet], { header: 1, defval: null }) as unknown[][]
      const isNew = rows.length > 0 && String((rows[0] as unknown[])[1] ?? '').trim() === '파이'
      for (const row of rows.slice(1)) {
        let pai: string, type: string, monthly: number[], annual: number, peak: number
        if (isNew) {
          type    = String(row[0] ?? '').trim()
          pai     = String(row[1] ?? '').trim()
          monthly = Array.from({ length: 12 }, (_, i) => Number(row[3 + i]) || 0)
          annual  = Number(row[15]) || 0
          peak    = Number(row[16]) || 0
        } else {
          const fullLabel = String(row[0] ?? '').trim()
          const sep = fullLabel.indexOf(' - ')
          pai  = sep > 0 ? fullLabel.slice(0, sep) : ''
          type = sep > 0 ? fullLabel.slice(sep + 3) : fullLabel
          monthly = Array.from({ length: 12 }, (_, i) => Number(row[1 + i]) || 0)
          annual  = Number(row[13]) || 0
          peak    = Number(row[14]) || 0
        }
        if (!type || !pai || annual <= 0) continue
        const key       = `${pai}|${type}`
        const fullLabel = `${pai} - ${type}`
        if (!housingMap[key]) housingMap[key] = mkRow('housing', key, fullLabel, pai, 'EA')
        housingMap[key].byYear[yr] = { annual, peak, monthly }
      }
    }
  }

  // ── 가공파일에 없지만 자재 관리에 품번 등록된 케이블 추가 (사용량 0) ──
  for (const [key, metaVal] of Object.entries(metadata.cable ?? {})) {
    if (cableMap[key]) continue
    const m = metaVal as Record<string, unknown>
    if (!String(m['품번'] ?? '').trim()) continue
    const [pai, label] = key.split('|', 2)
    if (!pai || !label) continue
    const unit = pai === '0.9mm' ? '개' : 'm'
    cableMap[key] = mkRow('cable', key, label, pai, unit)
  }

  // ── 메타/재고 연결 + 계산 ────────────────────────────────
  const latestYr = years[years.length - 1]
  const allRows: Step3Row[] = []

  for (const row of Object.values(cableMap)) {
    row.years        = years.filter(yr => row.byYear[yr])
    row.latestAnnual = row.byYear[latestYr]?.annual ?? 0
    row.latestPeak   = row.byYear[latestYr]?.peak   ?? 0
    const meta = (metadata.cable?.[row.key] ?? {}) as unknown as Record<string, unknown>
    const inv  = (inventory.cable?.[row.key] ?? {}) as unknown as Record<string, number>
    row.품번     = String(meta['품번'] ?? '')
    row.품명     = String(meta['품명'] ?? '')
    row.구매처   = String(meta['구매처'] ?? '')
    row.리드타임 = ltDays(meta['리드타임'], ltDefault)
    const cKind       = labelToKind(row.label)
    const cCagr       = (kindCagr && cKind) ? (kindCagr[cKind] ?? 0) : 0
    row.appliedCagr    = cCagr
    row.forecastAnnual = cCagr !== 0 ? Math.max(0, Math.round(row.latestAnnual * (1 + cCagr))) : row.latestAnnual
    const cPeakVals    = row.years.map(yr => row.byYear[yr]?.peak ?? 0)
    const cPeakAvg     = cPeakVals.length > 0 ? cPeakVals.reduce((s, v) => s + v, 0) / cPeakVals.length : 0
    const cForecastPeak = row.latestAnnual > 0 ? cPeakAvg * (row.forecastAnnual / row.latestAnnual) : cPeakAvg
    row.안전재고 = Math.round(cForecastPeak * (row.리드타임 / 30) * safetyK)
    row.현재고   = Number(inv['현재고'] ?? 0)
    row.기발주   = Number(inv['기발주'] ?? 0)
    row.발주필요량 = Math.max(0, row.forecastAnnual + row.안전재고 - row.현재고 - row.기발주)
    allRows.push(row)
  }

  for (const baseRow of Object.values(housingMap)) {
    baseRow.years        = years.filter(yr => baseRow.byYear[yr])
    baseRow.latestAnnual = baseRow.byYear[latestYr]?.annual ?? 0
    baseRow.latestPeak   = baseRow.byYear[latestYr]?.peak   ?? 0
    baseRow.appliedCagr    = 0
    baseRow.forecastAnnual = baseRow.latestAnnual

    const metaRaw = metadata.housing?.[baseRow.key]
    const comps   = metaRaw ? (Array.isArray(metaRaw) ? metaRaw : [metaRaw]) : [{}]
    const invRaw  = inventory.housing?.[baseRow.key]
    const invs    = invRaw  ? (Array.isArray(invRaw)  ? invRaw  : [invRaw])  : [{}]

    for (let ci = 0; ci < comps.length; ci++) {
      const comp = comps[ci] as Record<string, unknown>
      const inv  = ((ci < invs.length ? invs[ci] : invs[0]) ?? {}) as Record<string, number>
      const row: Step3Row = ci === 0 ? baseRow : { ...baseRow }
      row.품번     = String(comp['품번']   ?? '')
      row.품명     = String(comp['품명']   ?? '')
      row.구매처   = String(comp['구매처'] ?? '')
      row.리드타임 = ltDays(comp['리드타임'], ltDefault)
      const hPeakVals = baseRow.years.map(yr => baseRow.byYear[yr]?.peak ?? 0)
      const hPeakAvg  = hPeakVals.length > 0 ? hPeakVals.reduce((s, v) => s + v, 0) / hPeakVals.length : 0
      row.안전재고 = Math.round(hPeakAvg * (row.리드타임 / 30) * safetyK)
      row.현재고   = Number(inv['현재고'] ?? 0)
      row.기발주   = Number(inv['기발주'] ?? 0)
      row.발주필요량 = Math.max(0, baseRow.forecastAnnual + row.안전재고 - row.현재고 - row.기발주)
      row.isSubRow = ci > 0
      allRows.push(row)
    }
  }

  // ── 정렬: 케이블 우선, 파이 순, 타입 알파벳 ──────────────
  allRows.sort((a, b) => {
    if (a.type !== b.type) return a.type === 'cable' ? -1 : 1
    return (PAI_ORDER[a.pai] ?? 9) - (PAI_ORDER[b.pai] ?? 9) || a.label.localeCompare(b.label)
  })

  const nC = allRows.filter(r => r.type === 'cable').length
  const nH = allRows.filter(r => r.type === 'housing' && !r.isSubRow).length
  const nHComp = allRows.filter(r => r.type === 'housing').length
  const nMissing = allRows.filter(r => !r.품번).length
  logs.push(`집계 완료 — 케이블 ${nC}타입 / 하우징 ${nH}타입 (부품 ${nHComp}항목)`)
  logs.push(`안전재고 계수 k=${safetyK} (안전재고 = 피크평균 × LT/30 × ${safetyK})`)
  if (kindCagr && Object.keys(kindCagr).length) {
    const nApplied = allRows.filter(r => r.appliedCagr !== 0).length
    logs.push(`판매CAGR 적용: 케이블 ${nApplied}타입에 반영 (하우징은 최근년 실적 기준)`)
  }
  if (nMissing) logs.push(`⚠ 품번 미등록 ${nMissing}건 — 품번 관리 탭에서 입력 필요`)

  return { rows: allRows, years, logs, codeToCable }
}
