import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import { classifyOjcDetailed, extractFiberType, COLOR_MAP_OJC_DETAILED } from '../lib/ojcFilter'
import { today } from '../lib/download'

// Excel 15자리 정밀도 버그 보정:
// 16자리 이상 숫자를 Excel이 읽으면 마지막 자리가 0으로 잘림
// 예) 2026052990228841 → 2026052990228840 으로 저장됨
// 매칭 시 끝자리 0인 경우 1~9도 시도
function resolveKey<T>(map: Map<string, T>, 주문no: string, 품목코드: string): T | undefined {
  const exact = map.get(`${주문no}||${품목코드}`)
  if (exact !== undefined) return exact
  if (/^\d{15,}$/.test(주문no) && 주문no.endsWith('0')) {
    const base = 주문no.slice(0, -1)
    for (let d = 1; d <= 9; d++) {
      const v = map.get(`${base}${d}||${품목코드}`)
      if (v !== undefined) return v
    }
  }
  return undefined
}

function parse생산구분(val: unknown, qty: number): [number, number] {
  if (val === null || val === undefined) return [0, qty]
  const s = String(val).trim()
  if (s.includes('수입') || s.includes('FLC')) return [0, qty]
  if (s === '생' || s === '셍') return [qty, 0]
  const m = s.match(/^생(\d+)$/)
  if (m) { const n = parseInt(m[1]); return [n, Math.max(0, qty - n)] }
  return [0, qty]
}

type Status = 'match' | 'mismatch' | 'adj_only' | 'fallback' | 'none'

interface RowEntry {
  excelRow: number        // Excel 워크시트 행 번호
  주문NO: string; 품목코드: string; 품목명: string; 수량: number
  맥산수량: number; 수입수량: number
  조정_맥산: number; 조정_수입: number
  출하_맥산: number; 출하_수입: number
  status: Status
}

interface RunResult { counts: Record<Status, number>; rows: RowEntry[] }
type FileKey = '조정' | '출하' | '수익'
type ViewTab = 'run' | 'review'

const STATUS_META: Record<Status, { label: string; badge: string; row: string; text: string }> = {
  match:    { label: '✅ 두 파일 일치',           badge: 'bg-green-50  border-green-200  text-green-700',  row: '',             text: 'text-green-700'  },
  mismatch: { label: '⚠️ 불일치 (수동 확인 필요)', badge: 'bg-red-50    border-red-200    text-red-600',    row: 'bg-red-50',    text: 'text-red-600'    },
  adj_only: { label: '📋 조정 파일만',             badge: 'bg-blue-50   border-blue-200   text-blue-700',   row: 'bg-blue-50',   text: 'text-blue-700'   },
  fallback: { label: '🔄 출하현황 fallback',       badge: 'bg-yellow-50 border-yellow-200 text-yellow-700', row: 'bg-yellow-50', text: 'text-yellow-700' },
  none:     { label: '❌ 미매칭',                  badge: 'bg-gray-100  border-gray-200   text-gray-500',   row: 'bg-gray-50',   text: 'text-gray-400'   },
}

const EXCEL_FILLS: Record<Status, string | null> = {
  match: null, mismatch: 'FFD7D7', adj_only: 'D6E4F7', fallback: 'FFF3CD', none: 'E8E8E8',
}

export default function MonthlyTasksTab() {
  const [files, setFiles]   = useState<Record<FileKey, File | null>>({ 조정: null, 출하: null, 수익: null })
  const [loading, setLoading] = useState(false)
  const [result, setResult]   = useState<RunResult | null>(null)
  const [error, setError]     = useState('')
  const [outBuf, setOutBuf]   = useState<ArrayBuffer | null>(null)
  const [원본Buf, set원본Buf] = useState<ArrayBuffer | null>(null)   // 재생성용 원본
  const [colMap, setColMap]   = useState<Record<string, number>>({}) // 컬럼 위치 저장
  const [edits, setEdits]     = useState<Record<number, { 맥산: number; 수입: number }>>({})
  const [viewTab, setViewTab]   = useState<ViewTab>('run')
  const [reviewFilter, setReviewFilter] = useState<Status | 'all'>('all')
  const [saving, setSaving]   = useState(false)

  const [dragging, setDragging] = useState<FileKey | null>(null)
  const refs = {
    조정: useRef<HTMLInputElement>(null),
    출하: useRef<HTMLInputElement>(null),
    수익: useRef<HTMLInputElement>(null),
  }

  function setFile(key: FileKey, f: File) {
    setFiles(prev => ({ ...prev, [key]: f }))
    setResult(null); setOutBuf(null); set원본Buf(null); setEdits({}); setError('')
  }

  // ── 메인 실행 ──
  async function handleRun() {
    if (!files.조정 || !files.출하 || !files.수익) return
    setLoading(true); setResult(null); setError(''); setOutBuf(null); setEdits({})
    try {
      // 1. 조정 파일
      const adjWb = XLSX.read(await files.조정.arrayBuffer(), { type: 'array', raw: true })
      const adjWS = adjWb.Sheets['미출하현황 1']
      if (!adjWS) throw new Error('주문건 조정 파일에 "미출하현황 1" 시트가 없습니다.')
      const adjRows = XLSX.utils.sheet_to_json<unknown[]>(adjWS, { header: 1, defval: null }) as unknown[][]
      const adjH = (adjRows[0] as string[]).map(v => String(v ?? '').replace(/\n/g, '').trim())
      const ai = {
        주문NO:   adjH.findIndex(h => h.includes('주문NO')),
        품목코드: adjH.findIndex(h => h === '품목코드'),
        품목명:   adjH.findIndex(h => h === '품목명'),
        수입출고: adjH.findIndex(h => h.includes('수입출고')),
        생산출고: adjH.findIndex(h => h === '생산출고'),
      }
      if (ai.주문NO < 0 || ai.품목코드 < 0 || ai.생산출고 < 0 || ai.수입출고 < 0)
        throw new Error('주문건 조정 파일 컬럼을 찾을 수 없습니다.')
      const adjMap = new Map<string, [number, number, string]>()
      for (let i = 1; i < adjRows.length; i++) {
        const r = adjRows[i] as unknown[]
        const 주문no = String(r[ai.주문NO] ?? '').trim()
        const 품목코드 = String(r[ai.품목코드] ?? '').trim()
        if (!주문no || !품목코드) continue
        const key = `${주문no}||${품목코드}`
        const prev = adjMap.get(key) ?? [0, 0, String(r[ai.품목명] ?? '')]
        adjMap.set(key, [prev[0] + (Number(r[ai.생산출고]) || 0), prev[1] + (Number(r[ai.수입출고]) || 0), prev[2]])
      }

      // 2. 출하현황
      const shipWb = XLSX.read(await files.출하.arrayBuffer(), { type: 'array', raw: true })
      const shipWS = shipWb.Sheets[shipWb.SheetNames[0]]
      const shipRows = XLSX.utils.sheet_to_json<unknown[]>(shipWS, { header: 1, defval: null }) as unknown[][]
      const shipH = (shipRows[1] as string[]).map(v => String(v ?? '').trim())
      const si = {
        주문NO: shipH.indexOf('주문NO(4)'), 품목코드: shipH.indexOf('품목코드'),
        품목명: shipH.indexOf('품목명'),   수량: shipH.indexOf('수량'),
        생산구분: shipH.indexOf('생산구분'),
      }
      if (si.주문NO < 0 || si.품목코드 < 0 || si.생산구분 < 0)
        throw new Error('출하현황 파일 컬럼을 찾을 수 없습니다.')
      const shipMap = new Map<string, [number, number, string]>()
      for (let i = 2; i < shipRows.length; i++) {
        const r = shipRows[i] as unknown[]
        const 주문no = String(r[si.주문NO] ?? '').trim()
        const 품목코드 = String(r[si.품목코드] ?? '').trim()
        const 수량 = Number(r[si.수량]) || 0
        if (!주문no || !품목코드 || 수량 === 0) continue
        const [맥산, 수입] = parse생산구분(r[si.생산구분], 수량)
        const key = `${주문no}||${품목코드}`
        const prev = shipMap.get(key) ?? [0, 0, String(r[si.품목명] ?? '')]
        shipMap.set(key, [prev[0] + 맥산, prev[1] + 수입, prev[2]])
      }

      // 3. 수익률 분석 파일 (ExcelJS)
      const rawBuf = await files.수익.arrayBuffer()
      set원본Buf(rawBuf)
      const { buf, rows, counts, cols } = await fillWorkbook(rawBuf, adjMap, shipMap)
      setColMap(cols)
      setOutBuf(buf)
      setResult({ counts, rows })
      setViewTab('run')
    } catch (e) {
      setError(String(e))
    } finally {
      setLoading(false)
    }
  }

  // ── 워크북 채우기 (실행 + 수정 반영 공통) ──
  async function fillWorkbook(
    rawBuf: ArrayBuffer,
    adjMap: Map<string, [number, number, string]>,
    shipMap: Map<string, [number, number, string]>,
    overrides?: Record<number, { 맥산: number; 수입: number }>,
  ): Promise<{ buf: ArrayBuffer; rows: RowEntry[]; counts: Record<Status, number>; cols: Record<string, number> }> {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(rawBuf)
    const ws = wb.worksheets[0]

    let col주문NO = -1, col품목코드 = -1, col품목명 = -1, col규격명 = -1, col수량 = -1
    let col맥산수량 = -1, col수입수량 = -1, col공급가액 = -1
    ws.getRow(1).eachCell((cell, c) => {
      const v = String(cell.value ?? '').replace(/\n/g, '').trim()
      if (v.includes('주문NO'))  col주문NO   = c
      if (v === '품목코드')       col품목코드  = c
      if (v === '품목명')         col품목명    = c
      if (v === '규격명')         col규격명    = c
      if (v === '수량')           col수량      = c
      if (v === '맥산수량')       col맥산수량  = c
      if (v === '수입수량')       col수입수량  = c
      if (v === '공급가액')       col공급가액  = c
    })
    if (col주문NO < 0 || col맥산수량 < 0 || col수입수량 < 0)
      throw new Error('수익률 분석 파일 컬럼을 찾을 수 없습니다.')

    const 비고col = ws.columnCount + 1
    ws.getRow(1).getCell(비고col).value = '검토사항'
    ws.getRow(1).getCell(비고col).font = { bold: true }

    const counts: Record<Status, number> = { match: 0, mismatch: 0, adj_only: 0, fallback: 0, none: 0 }
    const rows: RowEntry[] = []
    let entryIdx = 0

    // ── 사전 스캔: 동일 (주문NO, 품목코드) 키의 총 수량 합산 ──
    // 조정 파일의 합계값을 각 행 수량에 비례해서 분배하기 위함
    const totalQtyMap = new Map<string, number>()
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r)
      const 주문no  = String(row.getCell(col주문NO).value  ?? '').trim()
      const 품목코드 = String(row.getCell(col품목코드).value ?? '').trim()
      if (!주문no || !품목코드) continue
      const 수량 = Number(row.getCell(col수량 > 0 ? col수량 : 1).value) || 0
      const key = `${주문no}||${품목코드}`
      totalQtyMap.set(key, (totalQtyMap.get(key) ?? 0) + 수량)
    }
    // 분배된 누적량 추적 (마지막 행에서 나머지 처리)
    const assignedMap = new Map<string, [number, number]>()
    const consumedQtyMap = new Map<string, number>()

    // 제품별 집계
    interface ProdSummary { 품목명: string; 규격명: string; 카테고리: string; 건수: number; 수량: number; 맥산: number; 수입: number; 공급가액: number }
    const prodMap = new Map<string, ProdSummary>()

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r)
      const 주문no  = String(row.getCell(col주문NO).value  ?? '').trim()
      const 품목코드 = String(row.getCell(col품목코드).value ?? '').trim()
      if (!주문no || !품목코드) continue

      const a = resolveKey(adjMap, 주문no, 품목코드)
      const s = resolveKey(shipMap, 주문no, 품목코드)
      const 수량   = Number(row.getCell(col수량    > 0 ? col수량    : 1).value) || 0
      const 품목명  = String(row.getCell(col품목명  > 0 ? col품목명  : col품목코드).value ?? '')
      const 규격명  = String(row.getCell(col규격명  > 0 ? col규격명  : 1).value ?? '')
      const 공급가액 = Number(row.getCell(col공급가액 > 0 ? col공급가액 : 1).value) || 0

      // 비례 분배 계산
      const key = `${주문no}||${품목코드}`
      const totalQty   = totalQtyMap.get(key) ?? 수량
      const consumed   = consumedQtyMap.get(key) ?? 0
      const isLastRow  = consumed + 수량 >= totalQty   // 이 키의 마지막 행
      const assigned   = assignedMap.get(key) ?? [0, 0]

      function distribute(total: number): number {
        if (isLastRow) return total - assigned[total === (a?.[0] ?? (s?.[0] ?? 0)) ? 0 : 1]
        return totalQty > 0 ? Math.round(total * 수량 / totalQty) : 0
      }

      let 맥산 = 0, 수입 = 0, status: Status
      if (a && s) {
        status = a[0] === s[0] && a[1] === s[1] ? 'match' : 'mismatch'
        if (isLastRow) {
          맥산 = a[0] - assigned[0]; 수입 = a[1] - assigned[1]
        } else {
          맥산 = totalQty > 0 ? Math.round(a[0] * 수량 / totalQty) : a[0]
          수입 = totalQty > 0 ? Math.round(a[1] * 수량 / totalQty) : a[1]
        }
      } else if (a) {
        status = 'adj_only'
        if (isLastRow) { 맥산 = a[0] - assigned[0]; 수입 = a[1] - assigned[1] }
        else { 맥산 = totalQty > 0 ? Math.round(a[0] * 수량 / totalQty) : a[0]; 수입 = totalQty > 0 ? Math.round(a[1] * 수량 / totalQty) : a[1] }
      } else if (s) {
        status = 'fallback'
        if (isLastRow) { 맥산 = s[0] - assigned[0]; 수입 = s[1] - assigned[1] }
        else { 맥산 = totalQty > 0 ? Math.round(s[0] * 수량 / totalQty) : s[0]; 수입 = totalQty > 0 ? Math.round(s[1] * 수량 / totalQty) : s[1] }
      } else {
        status = 'none'
      }

      // 행별 상한: 맥산 + 수입 ≤ 판매수량
      맥산 = Math.max(0, Math.min(맥산, 수량))
      수입 = Math.max(0, Math.min(수입, 수량 - 맥산))

      consumedQtyMap.set(key, consumed + 수량)
      assignedMap.set(key, [assigned[0] + 맥산, assigned[1] + 수입])


      // 수동 수정값 덮어쓰기
      if (overrides?.[entryIdx]) {
        맥산 = overrides[entryIdx].맥산
        수입 = overrides[entryIdx].수입
      }

      counts[status]++
      row.getCell(col맥산수량).value = 맥산
      row.getCell(col수입수량).value = 수입

      const note = status === 'mismatch'
        ? `⚠ 불일치 — 조정(맥산${a?.[0]}/수입${a?.[1]}) vs 출하(맥산${s?.[0]}/수입${s?.[1]})`
        : status === 'adj_only' ? '📋 조정 파일만 (출하현황 없음)'
        : status === 'fallback' ? '🔄 출하현황 기준 (조정 파일 없음)'
        : status === 'none'    ? '❌ 미매칭 — 두 파일 모두 없음' : ''
      row.getCell(비고col).value = note

      const fill = EXCEL_FILLS[status]
      if (fill) {
        ;[col맥산수량, col수입수량, 비고col].forEach(c => {
          row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + fill } }
          if (status === 'mismatch') row.getCell(c).font = { bold: true, color: { argb: 'FFCC0000' } }
        })
      }
      row.commit()

      rows.push({
        excelRow: r, 주문NO: 주문no, 품목코드, 품목명, 수량,
        맥산수량: 맥산, 수입수량: 수입,
        조정_맥산: a?.[0] ?? 0, 조정_수입: a?.[1] ?? 0,
        출하_맥산: s?.[0] ?? 0, 출하_수입: s?.[1] ?? 0,
        status,
      })
      entryIdx++

      // 제품별 집계
      if (품목코드) {
        const prev = prodMap.get(품목코드)
        if (prev) {
          prev.건수++; prev.수량 += 수량; prev.맥산 += 맥산; prev.수입 += 수입; prev.공급가액 += 공급가액
        } else {
          const detailCat  = classifyOjcDetailed(품목명) ?? '기타'
          const fiberType  = extractFiberType(품목명)
          // SM은 생략, MM/OM3만 괄호 표기
          const 카테고리 = fiberType === 'SM' ? detailCat : `${detailCat} (${fiberType})`
          prodMap.set(품목코드, { 품목명, 규격명, 카테고리, 건수: 1, 수량, 맥산, 수입, 공급가액 })
        }
      }
    }

    // ── 제품별 요약 시트 ──
    const CAT_ORDER = [
      'KT OJC-SP', 'KT OJC-SP (MM)', 'KT OJC-SP (OM3)',
      'KT OJC-DP', 'KT OJC-DP (MM)', 'KT OJC-DP (OM3)',
      'KT OJC-MOJC', 'KT OJC-MOJC (MM)', 'KT OJC-MOJC (OM3)',
      'KT OJC', 'KT OJC (MM)', 'KT OJC (OM3)',
      'LG OJC-SP', 'LG OJC-SP (MM)', 'LG OJC-SP (OM3)',
      'LG OJC-DP', 'LG OJC-DP (MM)', 'LG OJC-DP (OM3)',
      'LG OJC-MOJC', 'LG OJC-MOJC (MM)', 'LG OJC-MOJC (OM3)',
      'DROP-1C', 'DROP-2C', 'DROP',
      '피그테일-SK 성단용', '피그테일',
      'Optical Cable Parts', 'DX-MM', '기타',
    ]
    const sorted = [...prodMap.entries()].sort((a, b) => {
      const ai = CAT_ORDER.indexOf(a[1].카테고리)
      const bi = CAT_ORDER.indexOf(b[1].카테고리)
      const ci = (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi)
      return ci !== 0 ? ci : a[0].localeCompare(b[0])
    })

    const ws2 = wb.addWorksheet('제품별 요약')
    const HDRS = ['카테고리', '품목코드', '품목명', '규격명', '판매건수', '총수량(맥산+수입)', '맥산수량', '수입수량', '공급가액 합계']
    const HDR_ROW = ws2.addRow(HDRS)
    HDR_ROW.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } }
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
    })
    ws2.getRow(1).height = 20
    ws2.columns = [
      { width: 18 }, { width: 20 }, { width: 36 }, { width: 30 },
      { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 16 },
    ]

    let prevCat = ''
    for (const [code, p] of sorted) {
      const row = ws2.addRow([p.카테고리, code, p.품목명, p.규격명, p.건수, p.수량, p.맥산, p.수입, p.공급가액])
      // "(MM)"/"(OM3)" suffix 제거 후 색상 조회
      const baseCat  = p.카테고리.replace(/ \((MM|OM3)\)$/, '')
      const catColor = (COLOR_MAP_OJC_DETAILED[baseCat] ?? 'FFF5F5F5').replace(/^FF/, '')
      // 카테고리 변경 시 상단 경계선
      if (p.카테고리 !== prevCat) {
        row.eachCell(cell => { cell.border = { top: { style: 'medium', color: { argb: 'FF1F3864' } } } })
        prevCat = p.카테고리
      }
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + catColor } }
      row.getCell(1).font = { bold: true }
      ;[5, 6, 7, 8].forEach(c => {
        row.getCell(c).alignment = { horizontal: 'right' }
        row.getCell(c).numFmt = '#,##0'
      })
      row.getCell(9).alignment = { horizontal: 'right' }
      row.getCell(9).numFmt = '#,##0'
    }

    // 카테고리별 소계
    let catStart = 2
    let curCat = sorted[0]?.[1].카테고리 ?? ''
    const addSubtotal = (catLabel: string, endRow: number) => {
      const sub = ws2.addRow([`${catLabel} 소계`, '', '', '', '', '', '', '', ''])
      ;[5, 6, 7, 8, 9].forEach(c => {
        sub.getCell(c).value = { formula: `SUM(${ws2.getColumn(c).letter}${catStart}:${ws2.getColumn(c).letter}${endRow})` }
        sub.getCell(c).numFmt = '#,##0'
        sub.getCell(c).alignment = { horizontal: 'right' }
        sub.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } }
        sub.getCell(c).font = { bold: true }
      })
      sub.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } }
      sub.getCell(1).font = { bold: true }
    }
    // 소계 삽입은 카테고리 경계에서
    let dataRowNum = 2
    for (let i = 0; i < sorted.length; i++) {
      const cat = sorted[i][1].카테고리
      if (cat !== curCat) {
        addSubtotal(curCat, dataRowNum - 1)
        catStart = dataRowNum + 1
        curCat = cat
        dataRowNum++
      }
      dataRowNum++
    }
    if (sorted.length > 0) addSubtotal(curCat, dataRowNum - 1)

    const buf = await wb.xlsx.writeBuffer()
    const cols = { 주문NO: col주문NO, 품목코드: col품목코드, 맥산수량: col맥산수량, 수입수량: col수입수량 }
    return { buf: buf as ArrayBuffer, rows, counts, cols }
  }

  // ── 수정 반영 후 재다운로드 ──
  async function handleApplyAndDownload() {
    if (!원본Buf || !result) return
    setSaving(true)
    try {
      // adjMap, shipMap을 result.rows에서 역산해서 재구성
      const adjMap = new Map<string, [number, number, string]>()
      const shipMap = new Map<string, [number, number, string]>()
      result.rows.forEach(r => {
        const key = `${r.주문NO}||${r.품목코드}`
        if (r.status === 'match' || r.status === 'mismatch' || r.status === 'adj_only')
          adjMap.set(key, [r.조정_맥산, r.조정_수입, r.품목명])
        if (r.status === 'match' || r.status === 'mismatch' || r.status === 'fallback')
          shipMap.set(key, [r.출하_맥산, r.출하_수입, r.품목명])
      })
      const { buf } = await fillWorkbook(원본Buf, adjMap, shipMap, edits)
      downloadBuf(buf)
    } finally {
      setSaving(false)
    }
  }

  function downloadBuf(buf: ArrayBuffer) {
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = `OJC_판매수익률분석_${today()}.xlsx`
    a.click()
  }

  const reviewRows = result
    ? result.rows.filter(r => reviewFilter === 'all' ? r.status !== 'match' : r.status === reviewFilter)
    : []

  const hasEdits = Object.keys(edits).length > 0

  return (
    <div className="flex flex-col gap-0">
      {/* 서브탭 */}
      <div className="flex border-b border-gray-200 mb-6">
        {([['run', '📤 실행'], ['review', '🔍 검토']] as const).map(([id, label]) => (
          <button key={id} onClick={() => setViewTab(id)}
            disabled={id === 'review' && !result}
            className={`px-5 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors whitespace-nowrap ${
              viewTab === id
                ? 'border-[#2E75B6] text-[#2E75B6]'
                : 'border-transparent text-gray-500 hover:text-gray-700 disabled:text-gray-300 disabled:cursor-not-allowed'
            }`}
          >
            {label}
            {id === 'review' && result && result.rows.filter(r => r.status !== 'match').length > 0 && (
              <span className="ml-1.5 text-xs bg-red-100 text-red-600 rounded-full px-1.5 py-0.5">
                {result.rows.filter(r => r.status !== 'match').length}
              </span>
            )}
          </button>
        ))}
      </div>

      {/* ── 실행 탭 ── */}
      {viewTab === 'run' && (
        <div className="flex flex-col gap-5 max-w-4xl">
          <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md">
            <div className="font-semibold text-[#1F3864] mb-1">📅 월간업무 — 판매 수익률 맥산/수입 자동 분류</div>
            <div className="text-sm text-gray-600">
              주문건 조정 파일(메인) + 출하현황(보조·검증)을 기반으로 수익률 분석 파일의 <b>맥산수량 / 수입수량</b>을 자동으로 채웁니다.
            </div>
          </div>

          <div className="grid grid-cols-3 gap-3">
            {([
              { key: '조정' as FileKey, num: '①', label: '주문건 조정 파일', sub: '"미출하현황 1" 시트 — 생산출고/수입출고 (메인)' },
              { key: '출하' as FileKey, num: '②', label: '출하현황 파일 (EMP)', sub: '생산구분 컬럼 포함 (보조·검증)' },
              { key: '수익' as FileKey, num: '③', label: '수익률 분석 파일 (영업팀)', sub: '맥산수량·수입수량 컬럼 포함 (출력 대상)' },
            ]).map(({ key, num, label, sub }) => (
              <div key={key}
                onClick={() => refs[key].current?.click()}
                onDragOver={e => { e.preventDefault(); setDragging(key) }}
                onDragEnter={e => { e.preventDefault(); setDragging(key) }}
                onDragLeave={() => setDragging(null)}
                onDrop={e => {
                  e.preventDefault(); setDragging(null)
                  const f = e.dataTransfer.files?.[0]
                  if (f && (f.name.endsWith('.xlsx') || f.name.endsWith('.xls'))) setFile(key, f)
                }}
                className={`cursor-pointer border-2 border-dashed rounded-xl px-4 py-4 transition-colors ${
                  dragging === key
                    ? 'border-[#2E75B6] bg-[#e8f4fd] scale-[1.02]'
                    : files[key]
                    ? 'border-[#2E75B6] bg-[#f0f4fa]'
                    : 'border-gray-300 bg-white hover:border-[#2E75B6]'
                }`}>
                <input ref={refs[key]} type="file" accept=".xlsx,.xls" className="hidden"
                  onChange={e => { const f = e.target.files?.[0]; if (f) { setFile(key, f); e.target.value = '' } }} />
                <div className="text-xs text-gray-400 mb-1">{num}</div>
                <div className="text-sm font-semibold text-gray-700 mb-1">{label}</div>
                <div className="text-xs text-gray-400 mb-2">{sub}</div>
                {dragging === key
                  ? <div className="text-xs text-[#2E75B6] font-semibold">여기에 놓으세요</div>
                  : files[key]
                  ? <div className="text-xs text-[#2E75B6] font-medium truncate">📎 {files[key]!.name}</div>
                  : <div className="text-xs text-gray-400">클릭 또는 파일을 여기로 드래그</div>
                }
              </div>
            ))}
          </div>

          <button onClick={handleRun} disabled={loading || !files.조정 || !files.출하 || !files.수익}
            className="px-6 py-2.5 bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold rounded-lg transition text-sm">
            {loading ? '⏳ 처리 중...' : '▶ 맥산/수입 자동 분류 실행'}
          </button>

          {error && <div className="bg-red-50 border border-red-200 rounded-lg px-4 py-3 text-sm text-red-600">❌ {error}</div>}

          {result && (
            <div className="flex flex-col gap-3">
              <div className="grid grid-cols-5 gap-2">
                {(Object.keys(STATUS_META) as Status[]).map(s => (
                  <div key={s} className={`border rounded-lg px-3 py-3 text-center ${STATUS_META[s].badge}`}>
                    <div className="text-xl font-bold">{result.counts[s]}</div>
                    <div className="text-xs mt-0.5 leading-tight">{STATUS_META[s].label}</div>
                  </div>
                ))}
              </div>
              <div className="flex items-center gap-3 bg-green-50 border border-green-200 rounded-lg px-4 py-3">
                <span className="text-green-700 font-medium text-sm">✅ 처리 완료</span>
                <button onClick={() => setViewTab('review')}
                  className="px-3 py-1.5 bg-white border border-gray-300 text-gray-600 text-xs rounded hover:bg-gray-50 transition">
                  🔍 검토 탭으로
                </button>
                <button onClick={() => outBuf && downloadBuf(outBuf)}
                  className="ml-auto px-4 py-1.5 bg-[#2E75B6] text-white text-xs font-medium rounded hover:bg-[#1a5a9e] transition">
                  📥 Excel 다운로드
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── 검토 탭 ── */}
      {viewTab === 'review' && result && (
        <div className="flex flex-col gap-4">
          {/* 필터 + 저장 버튼 */}
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-sm text-gray-500 font-medium">필터</span>
            {(['all', 'mismatch', 'fallback', 'adj_only', 'none'] as const).map(f => {
              const count = f === 'all'
                ? result.rows.filter(r => r.status !== 'match').length
                : result.counts[f]
              if (f !== 'all' && count === 0) return null
              return (
                <button key={f} onClick={() => setReviewFilter(f)}
                  className={`px-3 py-1 rounded-full text-xs font-semibold border transition-colors ${
                    reviewFilter === f
                      ? 'bg-[#2E75B6] text-white border-[#2E75B6]'
                      : 'bg-white text-gray-500 border-gray-300 hover:border-[#2E75B6]'
                  }`}>
                  {f === 'all' ? `전체 (${count})` : `${STATUS_META[f].label} (${count})`}
                </button>
              )
            })}

            {hasEdits && (
              <button onClick={handleApplyAndDownload} disabled={saving}
                className="ml-auto px-4 py-1.5 bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white text-xs font-semibold rounded transition">
                {saving ? '⏳ 처리 중...' : `💾 수정 반영 후 다운로드 (${Object.keys(edits).length}건)`}
              </button>
            )}
            {!hasEdits && (
              <button onClick={() => outBuf && downloadBuf(outBuf)}
                className="ml-auto px-4 py-1.5 bg-[#2E75B6] text-white text-xs font-medium rounded hover:bg-[#1a5a9e] transition">
                📥 Excel 다운로드
              </button>
            )}
          </div>

          {/* 안내 (불일치 있을 때만) */}
          {result.counts.mismatch > 0 && (
            <div className="bg-red-50 border border-red-200 rounded-lg px-4 py-2.5 text-xs text-red-600">
              ⚠️ 불일치 항목의 <b>맥산수량 / 수입수량</b> 셀을 직접 수정한 후 <b>수정 반영 후 다운로드</b>를 클릭하세요.
            </div>
          )}

          {/* 테이블 */}
          <div className="border border-gray-200 rounded-lg overflow-hidden">
            <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
              <table className="w-full text-xs">
                <thead className="bg-[#1F3864] text-white sticky top-0 z-10">
                  <tr>
                    {['상태', '주문NO', '품목코드', '품목명', '수량', '맥산수량', '수입수량', '조정_맥산', '조정_수입', '출하_맥산', '출하_수입'].map(h => (
                      <th key={h} className="px-3 py-2 text-left font-medium whitespace-nowrap">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {reviewRows.length === 0
                    ? <tr><td colSpan={11} className="px-4 py-8 text-center text-gray-400">해당 항목 없음</td></tr>
                    : reviewRows.map((r) => {
                      const idx = result.rows.indexOf(r)
                      const edited = edits[idx]
                      const 현재맥산 = edited?.맥산 ?? r.맥산수량
                      const 현재수입 = edited?.수입 ?? r.수입수량
                      const isEdited = !!edited

                      return (
                        <tr key={r.excelRow} className={`${STATUS_META[r.status].row} ${isEdited ? 'ring-1 ring-inset ring-blue-400' : ''}`}>
                          <td className="px-3 py-1.5 whitespace-nowrap">
                            <span className={`font-semibold ${STATUS_META[r.status].text}`}>
                              {STATUS_META[r.status].label}
                            </span>
                            {isEdited && <span className="ml-1 text-blue-600 font-bold">✏️</span>}
                          </td>
                          <td className="px-3 py-1.5 text-gray-600 max-w-[150px] truncate whitespace-nowrap" title={r.주문NO}>{r.주문NO}</td>
                          <td className="px-3 py-1.5 font-mono whitespace-nowrap">{r.품목코드}</td>
                          <td className="px-3 py-1.5 text-gray-600 max-w-[160px] truncate" title={r.품목명}>{r.품목명}</td>
                          <td className="px-3 py-1.5 text-right tabular-nums">{r.수량}</td>

                          {/* 맥산수량 — 불일치만 편집 가능 */}
                          <td className="px-2 py-1">
                            {r.status === 'mismatch' ? (
                              <input
                                type="number" min={0}
                                value={현재맥산}
                                onChange={e => {
                                  const val = Math.max(0, parseInt(e.target.value) || 0)
                                  setEdits(prev => ({ ...prev, [idx]: { 맥산: val, 수입: prev[idx]?.수입 ?? r.수입수량 } }))
                                }}
                                className="w-16 text-right border border-gray-300 rounded px-1 py-0.5 text-xs focus:border-blue-400 focus:outline-none tabular-nums"
                              />
                            ) : (
                              <span className="tabular-nums font-semibold">{현재맥산}</span>
                            )}
                          </td>

                          {/* 수입수량 — 불일치만 편집 가능 */}
                          <td className="px-2 py-1">
                            {r.status === 'mismatch' ? (
                              <input
                                type="number" min={0}
                                value={현재수입}
                                onChange={e => {
                                  const val = Math.max(0, parseInt(e.target.value) || 0)
                                  setEdits(prev => ({ ...prev, [idx]: { 맥산: prev[idx]?.맥산 ?? r.맥산수량, 수입: val } }))
                                }}
                                className="w-16 text-right border border-gray-300 rounded px-1 py-0.5 text-xs focus:border-blue-400 focus:outline-none tabular-nums"
                              />
                            ) : (
                              <span className="tabular-nums font-semibold">{현재수입}</span>
                            )}
                          </td>

                          <td className={`px-3 py-1.5 text-right tabular-nums ${r.status === 'mismatch' ? 'text-blue-700 font-bold' : 'text-gray-400'}`}>{r.조정_맥산 || '-'}</td>
                          <td className={`px-3 py-1.5 text-right tabular-nums ${r.status === 'mismatch' ? 'text-blue-700 font-bold' : 'text-gray-400'}`}>{r.조정_수입 || '-'}</td>
                          <td className={`px-3 py-1.5 text-right tabular-nums ${r.status === 'mismatch' ? 'text-orange-600 font-bold' : 'text-gray-400'}`}>{r.출하_맥산 || '-'}</td>
                          <td className={`px-3 py-1.5 text-right tabular-nums ${r.status === 'mismatch' ? 'text-orange-600 font-bold' : 'text-gray-400'}`}>{r.출하_수입 || '-'}</td>
                        </tr>
                      )
                    })
                  }
                </tbody>
              </table>
            </div>
          </div>
          <div className="text-xs text-gray-400">
            ※ 불일치 행: <span className="text-blue-700 font-bold">파랑</span> = 조정 파일값 / <span className="text-orange-600 font-bold">주황</span> = 출하현황값 | ✏️ = 수동 수정됨
          </div>
        </div>
      )}
    </div>
  )
}
