/**
 * convert_core.py → TypeScript 포팅
 * 구매조회 / 구매현황 ERP 파일 → 연도별 시트(가공파일) 변환
 */
import * as XLSX from 'xlsx'

// ── 타입 ─────────────────────────────────────────────────────
interface ErpRow {
  연도: string; 월: number; 품목코드: string
  품목명: string; 규격명: string; 수량: number
}

interface PivotRow {
  품목코드: string; 품목명: string; 규격명: string
  monthly: number[]  // 0-indexed, months 1-12
  코어수: number; 케이블종류: string; 파이: string
  케이블길이: number | null; 타입1: string; 타입2: string
}

// ── 파싱 함수 ────────────────────────────────────────────────
export function extractTypes(p: string): [string, string] {
  p = p.trim()
  if (p.startsWith('DROP')) {
    const m = p.match(/\((.+?)\)/)
    if (m) {
      const t = m[1].match(/(?:SC|LC|FC)\/(?:PC|APC)/g) ?? []
      if (t.length >= 2) return [t[0]!, t[1]!]
      if (t.length === 1) return [t[0]!, t[0]!]
    }
    return ['', '']
  }
  if (p.startsWith('PIGTAIL')) {
    const t = p.match(/(?:SC|LC|FC)\/(?:PC|APC)/g) ?? []
    return t.length ? [t[0]!, ''] : ['', '']
  }
  if (p.startsWith('OJC-')) {
    const parts = p.split('-')
    const conn = parts.find(x => /^(SC|LC|FC)\/(SC|LC|FC)$/.test(x))
    const ferr = parts.find(x => /^(PC|APC)\/(PC|APC)$/.test(x))
    if (conn && ferr) {
      const [cA, cB] = conn.split('/'); const [fA, fB] = ferr.split('/')
      return [`${cA}/${fA}`, `${cB}/${fB}`]
    }
    return ['', '']
  }
  if (/^(?:SOJC|DOJC|MOJC|Optical Cable Parts)/.test(p)) {
    const t = p.match(/(?:SC|LC|FC)\/(?:PC|APC)/g) ?? []
    if (t.length >= 2) return [t[0]!, t[1]!]
    if (t.length === 1) return [t[0]!, t[0]!]
    return ['', '']
  }
  return ['', '']
}

export function deriveKind(p: string, g: string, core: number): string {
  if (p.startsWith('PIGTAIL'))
    return /-MM\b|-MM-|MM\(OM3\)/.test(p) ? 'om1-pigtail' : 'pigtail'
  if (p.startsWith('DROP')) return 'drop'
  if (p.startsWith('Optical Cable Parts')) return 'a2'
  if (p.startsWith('MOJC')) return `b3-${core}c`
  if (p.includes('OJC-C2')) return core === 4 ? `b3-${core}c` : `a1-${core}c`
  if (/\bB3\b/i.test(g)) return 'b3'
  if (p.includes('MM(OM3)') || /\bOM3\b/i.test(g) || p.includes('-OM3')) return 'om3'
  if (/-MM\b|-MM-/.test(p)) return 'om1'
  for (const [c, l] of [['청','a1-청'],['적','a1-적'],['녹','a1-녹'],['자','a1-자']])
    if (g.includes(c)) return l
  for (const [c, l] of [['청','a1-청'],['적','a1-적'],['녹','a1-녹'],['자','a1-자']])
    if (p.includes(c)) return l
  return 'a1'
}

export function derivePai(p: string, g: string): string {
  if (p.startsWith('PIGTAIL')) {
    if (p.includes('(0.9mm)') || (p.includes('0.9mm') && !p.includes('2.0mm'))) return '0.9mm'
    if (p.includes('(2.0mm)') || /-MM\b|-MM-/.test(p) || p.includes('MM(OM3)')) return '2.0mm'
    return '0.9mm'
  }
  if (p.startsWith('DROP')) return '3.0mm'
  if (/-MM\b|-MM-/.test(p) || p.includes('MM(OM3)') || p.includes('OM3')) return '2.0mm'
  if (/^(?:MOJC|Optical Cable Parts)/.test(p) || p.includes('OJC-C2') || p.includes('OJC-A1')) return '2.0mm'
  if (/^(?:SOJC|DOJC)/.test(p)) {
    if (/3\.0mm|3\.0MM|3MM/i.test(g)) return '3.0mm'
    return '2.0mm'
  }
  if (/3\.0mm|3\.0MM/i.test(g)) return '3.0mm'
  return '2.0mm'
}

export function deriveCore(p: string): number {
  let m: RegExpMatchArray | null
  m = p.match(/MOJC-(?:SM|MM)-(\d+)C/); if (m) return parseInt(m[1])
  m = p.match(/PIGTAIL-[A-Z/()A-Z-]+-(\d+)C\b/); if (m) return parseInt(m[1])
  if (p.startsWith('PIGTAIL')) { m = p.match(/-(\d+)C\b/); return m ? parseInt(m[1]) : 1 }
  m = p.match(/OJC-C2-.*-(\d+)C/); if (m) return parseInt(m[1])
  if (p.startsWith('Optical Cable Parts')) { m = p.match(/(\d+)Core/); return m ? parseInt(m[1]) : 1 }
  if (p.startsWith('DROP') && p.includes('2C')) return 2
  if (p.startsWith('DROP')) return 1
  if (p.startsWith('DOJC')) return 2
  if (p.startsWith('SOJC')) return 1
  if (p.endsWith('-SP')) return 1
  if (p.endsWith('-DP')) return 2
  m = p.match(/-(\d+)C$/); if (m) return parseInt(m[1])
  return 1
}

export function deriveLength(p: string, g: string): number | null {
  let m: RegExpMatchArray | null
  if (p.startsWith('Optical Cable Parts')) {
    m = p.match(/-(\d+(?:\.\d+)?)m-/i); if (m) return parseFloat(m[1])
  }
  if (p.includes('OJC-A1') || p.includes('OJC-C2')) {
    m = p.match(/-(?:SM|MM)(?:\(OM3\))?-(\d+(?:\.\d+)?)-(?:PC|APC)/); if (m) return parseFloat(m[1])
  }
  m = g.match(/(\d+(?:\.\d+)?)\s*[mM]\b/); if (m) return parseFloat(m[1])
  m = g.match(/(\d+(?:\.\d+)?)\s*[mM]/); if (m) return parseFloat(m[1])
  m = p.match(/\[(\d+(?:\.\d+)?)M/i); if (m) return parseFloat(m[1])
  return null
}

// ── 하우징 계산 ──────────────────────────────────────────────
const MM_KINDS = new Set(['om1', 'om1-pigtail', 'om3'])
const PRIMARY_20:   Record<string, number> = { 'LC/PC':23,'LC/APC':25,'SC/PC':27,'SC/APC':29,'FC/PC':31,'FC/APC':33 }
const SECONDARY_20: Record<string, number> = { 'LC/PC':24,'LC/APC':26,'SC/PC':28,'SC/APC':30,'FC/PC':32,'FC/APC':34 }
const BEIGE_COL = 35
const PRIMARY_30:   Record<string, number> = { 'LC/PC':36,'LC/APC':38,'SC/PC':40,'SC/APC':42,'FC/PC':44,'FC/APC':46 }
const SECONDARY_30: Record<string, number> = { 'LC/PC':37,'LC/APC':39,'SC/PC':41,'SC/APC':43,'FC/PC':45,'FC/APC':47 }

function isMulticore(kind: string): boolean { return /^(b3|a1)-\d+c$/i.test(kind) }

function calcHousing(row: PivotRow): Record<number, number> {
  const { 타입1: t1, 타입2: t2, 파이: pai, 케이블종류: kind, 코어수: core, monthly } = row
  const qty = monthly.reduce((a, b) => a + b, 0)
  if (qty === 0 || !t1) return {}
  const isMm = MM_KINDS.has(kind), useRed = core >= 2 && !isMulticore(kind)
  const is20 = pai === '2.0mm', is30 = pai === '3.0mm', cps = core
  const res: Record<number, number> = {}
  const add = (col: number | undefined, val: number) => { if (col && val) res[col] = (res[col] ?? 0) + val }
  const addPrimary = (t: string, amount: number) => {
    if (t === 'LC/PC' && isMm && is20) add(BEIGE_COL, amount)
    else if (is20 && PRIMARY_20[t])   add(PRIMARY_20[t], amount)
    else if (is30 && PRIMARY_30[t])   add(PRIMARY_30[t], amount)
  }
  if (t1) addPrimary(t1, cps * qty)
  if (t2) {
    if (useRed) {
      if (is20 && SECONDARY_20[t2])   add(SECONDARY_20[t2], cps * qty)
      else if (is30 && SECONDARY_30[t2]) add(SECONDARY_30[t2], cps * qty)
    } else {
      addPrimary(t2, cps * qty)
    }
  }
  return res
}

// ── 날짜 파싱 ────────────────────────────────────────────────
function parseBuyDate(val: unknown): [string | null, number | null] {
  if (!val) return [null, null]
  if (val instanceof Date) return [String(val.getFullYear()), val.getMonth() + 1]
  // xlsx serial number → date
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
  return [null, null]
}

// ── 행 파싱 ──────────────────────────────────────────────────
function parseRows(ws: XLSX.WorkSheet, sheetType: 'ojc' | 'purchase', logs: string[]): ErpRow[] {
  const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null, raw: false }) as unknown[][]
  const rawNum = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null, raw: true }) as unknown[][]
  const rows: ErpRow[] = []
  let skippedNoKey = 0, skippedDate = 0

  for (let i = 2; i < raw.length; i++) {
    const cells = raw[i] as unknown[]
    const cellsNum = rawNum[i] as unknown[]
    let buyNo: unknown, pc: unknown, pr: unknown, gyRaw: unknown, qty: unknown

    if (sheetType === 'ojc') {
      buyNo = cellsNum[2]; pc = cells[3]; pr = cells[4]; gyRaw = cells[6]; qty = cells[7]
    } else {
      buyNo = cellsNum[0]; pc = cells[4]; pr = cells[5]; gyRaw = null; qty = cells[6]
    }

    if (!buyNo || !pr || !qty) { skippedNoKey++; continue }
    const [year, month] = parseBuyDate(buyNo)
    if (!year || !month) { skippedDate++; continue }

    const s = String(pr).trim().replace(/\s*외\s*\d+건\s*$/, '')
    const bm = s.match(/^(.+?)\s*\[(.+)\]\s*$/)
    let pn: string, gy: string
    if (bm) { pn = bm[1].trim(); gy = bm[2].trim() }
    else { pn = s; gy = gyRaw ? String(gyRaw).trim() : '' }

    const q = parseInt(String(qty))
    if (isNaN(q) || q <= 0) continue

    rows.push({ 연도: year, 월: month, 품목코드: pc ? String(pc).trim() : '', 품목명: pn, 규격명: gy, 수량: q })
  }

  if (!rows.length) {
    const samples = rawNum.slice(2, 5).map(r => (r as unknown[]).slice(0, 8))
    logs.push(`[진단] 샘플 행(첫 3행): ${JSON.stringify(samples)}`)
    logs.push(`[진단] 필수 값 없음: ${skippedNoKey}행 / 날짜 형식 불일치: ${skippedDate}행`)
  }
  return rows
}

function detectColFormat(ws: XLSX.WorkSheet): 'ojc' | 'purchase' {
  const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]
  const h0 = String(raw[1]?.[0] ?? '').trim()
  return ['거래처','구매처','공급처','납품처'].some(kw => h0.includes(kw)) ? 'ojc' : 'purchase'
}

// ── pivot (pandas pivot_table 대체) ─────────────────────────
function buildMerged(rows: ErpRow[]): PivotRow[] {
  const map = new Map<string, PivotRow>()
  for (const row of rows) {
    const key = `${row.품목코드}||${row.품목명}||${row.규격명}`
    if (!map.has(key)) {
      const core = deriveCore(row.품목명)
      const kind = deriveKind(row.품목명, row.규격명, core)
      const pai  = derivePai(row.품목명, row.규격명)
      const [t1, t2] = extractTypes(row.품목명)
      map.set(key, {
        품목코드: row.품목코드, 품목명: row.품목명, 규격명: row.규격명,
        monthly: new Array(12).fill(0),
        코어수: core, 케이블종류: kind, 파이: pai,
        케이블길이: deriveLength(row.품목명, row.규격명),
        타입1: t1, 타입2: t2,
      })
    }
    const pr = map.get(key)!
    if (row.월 >= 1 && row.월 <= 12) pr.monthly[row.월 - 1] += row.수량
  }
  return [...map.values()].sort((a, b) => a.품목코드.localeCompare(b.품목코드))
}

// ── 시트 작성 (SheetJS) ──────────────────────────────────────
function writeSheets(wb_out: XLSX.WorkBook, merged: PivotRow[], YY: string) {
  const cH = ['품목코드','품목명','규격명','케이블 종류','파이','코어수','케이블 길이','타입 1','타입2',
    ...Array.from({length:12}, (_,i) => `${YY}년${String(i+1).padStart(2,'0')}월`),
    '케이블 사용량','최고제작량','최고판매 케이블 소요량']

  const hH = ['품목코드','품목명','규격명','케이블 종류','파이','코어수','케이블 길이','타입 1','타입2',
    ...Array.from({length:12}, (_,i) => `${YY}년${String(i+1).padStart(2,'0')}월`), '합계',
    '2.0MM - LC/PC(청색)','2.0MM - LC/PC(적색)','2.0MM - LC/APC(녹색)','2.0MM - LC/APC(적색)',
    '2.0MM - SC/PC(청색)','2.0MM - SC/PC(적색)','2.0MM - SC/APC(녹색)','2.0MM - SC/APC(적색)',
    '2.0MM - FC/PC(흑색)','2.0MM - FC/PC(적색)','2.0MM - FC/APC(녹색)','2.0MM - FC/APC(적색)',
    '2.0MM - BEIGE(OM1·OM3)',
    '3.0MM - LC/PC(청색)','3.0MM - LC/PC(적색)','3.0MM - LC/APC(녹색)','3.0MM - LC/APC(적색)',
    '3.0MM - SC/PC(청색)','3.0MM - SC/PC(적색)','3.0MM - SC/APC(녹색)','3.0MM - SC/APC(적색)',
    '3.0MM - FC/PC(흑색)','3.0MM - FC/PC(적색)','3.0MM - FC/APC(녹색)','3.0MM - FC/APC(적색)',
    '검증','계산수량']

  const cData: (string | number | null)[][] = [cH]
  const hData: (string | number | null)[][] = [hH]

  for (const row of merged) {
    const total  = row.monthly.reduce((a, b) => a + b, 0)
    const peak   = Math.max(...row.monthly)
    const len    = row.케이블길이 ?? 0
    const usage  = Math.round(total * len)
    const peakUsage = Math.round(peak * len)
    const core   = row.코어수

    cData.push([
      row.품목코드, row.품목명, row.규격명,
      row.케이블종류, row.파이, core, row.케이블길이 ?? null,
      row.타입1 || null, row.타입2 || null,
      ...row.monthly.map(v => v || null),
      usage || null, peak || null, peakUsage || null,
    ])

    const hRow: (string | number | null)[] = [
      row.품목코드, row.품목명, row.규격명,
      row.케이블종류, row.파이, core, row.케이블길이 ?? null,
      row.타입1 || null, row.타입2 || null,
      ...row.monthly.map(v => v || null),
      total || null,
      ...new Array(25).fill(null),
      null, null,
    ]
    const housing = calcHousing(row)
    for (const [col1, val] of Object.entries(housing))
      hRow[parseInt(col1) - 1] = val || null

    hRow[47] = core === 1 ? total * 2 : total * core * 2
    hRow[48] = Object.values(housing).reduce((a, b) => a + b, 0) || null

    hData.push(hRow)
  }

  const wsc = XLSX.utils.aoa_to_sheet(cData)
  const wsh = XLSX.utils.aoa_to_sheet(hData)
  XLSX.utils.book_append_sheet(wb_out, wsc, `${YY}년_케이블`)
  XLSX.utils.book_append_sheet(wb_out, wsh, `${YY}년 하우징`)
}

// ── 메인 함수 ────────────────────────────────────────────────
export function preprocessERP(fileBuffer: ArrayBuffer, logs: string[]): ArrayBuffer {
  const wb_in = XLSX.read(fileBuffer, { type: 'array', cellDates: false, raw: false })
  const sheets = wb_in.SheetNames

  let rows: ErpRow[] = []

  if (sheets.includes('구매조회')) {
    logs.push('구매조회 형식 감지 → 자동 변환 시작')
    rows = parseRows(wb_in.Sheets['구매조회'], 'ojc', logs)
  } else if (sheets.includes('구매현황')) {
    const ws = wb_in.Sheets['구매현황']
    const fmt = detectColFormat(ws)
    logs.push(fmt === 'ojc'
      ? '구매현황 파일 (맥산 납품 형식) 감지 → 자동 변환 시작'
      : '구매현황 형식 감지 → 자동 변환 시작')
    rows = parseRows(ws, fmt, logs)
  } else if (sheets.includes('생산현황')) {
    logs.push('생산현황 형식 감지 → 자동 변환 시작')
    rows = parseRows(wb_in.Sheets['생산현황'], 'ojc', logs)
  } else {
    throw new Error(`지원하지 않는 파일 형식입니다. 시트 목록: ${sheets.join(', ')}`)
  }

  if (!rows.length) {
    const diag = logs.filter(l => l.includes('[진단]')).join(' | ')
    throw new Error(`파싱된 데이터가 없습니다. 파일 형식을 확인해주세요.${diag ? '\n' + diag : ''}`)
  }

  const yearMap = new Map<string, ErpRow[]>()
  for (const row of rows) {
    if (!yearMap.has(row.연도)) yearMap.set(row.연도, [])
    yearMap.get(row.연도)!.push(row)
  }

  const wb_out = XLSX.utils.book_new()

  for (const [year, yRows] of [...yearMap.entries()].sort()) {
    const YY = year.slice(-2)
    const merged = buildMerged(yRows)
    writeSheets(wb_out, merged, YY)
    logs.push(`  ${year}년: ${merged.length}개 품목 변환 완료`)
  }

  const buf = XLSX.write(wb_out, { type: 'array', bookType: 'xlsx' }) as Uint8Array
  return buf.buffer as ArrayBuffer
}
