import { useState, useCallback } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import { downloadXlsx } from '../lib/download'

// ── 타입 ──────────────────────────────────────────────────────────────
interface LotEntry { lot: string | null; qty: number; autoLoc?: boolean }

// ── ZZ-ZZ 품목 자동 로케이션 배정 ──────────────────────────────────────
function assignLocation(name: string): string | null {
  const n = name.toLowerCase()
  if (n.includes('drop-cable') || n.includes('drop optical cable') || n.includes('drop cable'))
    return 'D1'
  if (n.includes('dx-mm'))   return 'D2-MM'
  if (n.includes('dojc-mm')) return 'DO-MM'
  if (n.includes('dojc'))    return 'LG-DO'
  if (n.includes('sojc'))    return 'LG-SO'
  if (n.includes('mojc'))    return 'LG-MO'
  if (n.includes('adapter') || n.includes('어댑터') || n.includes('감쇠기') || n.includes('attenuator'))
    return 'EX'
  if (n.includes('열수축슬리브') || n.includes('splice protection sleeve'))
    return 'EX'
  if (n.includes('ojc housing kit') || n.includes('housing kit')) return 'M1'
  if (n.includes('ferrule'))                                       return 'M1'
  if (n.includes('pigtail'))                                       return 'PI-TA'
  if (n.includes('optical cable parts') || n.includes('optical cable part')) return 'O1'
  if (n.includes('optical cable 0.9'))   return 'K9-PI'
  if (n.includes('optical cable'))       return 'K1'
  if (n.includes('ojc'))                 return 'KT'
  return null
}

interface ReconRow {
  code:       string
  name:       string
  empRaw:     number   // EMP 재고 합산
  shipDeduct: number   // 출고확인서 출고수량
  empAdj:     number   // empRaw - shipDeduct
  ecRaw:      number   // 이카운트 재고
  unshipped:  number   // 미출하 잔량
  ecAdj:      number   // ecRaw + unshipped
  diff:       number   // empAdj - ecAdj
  status:     'match' | 'diff' | 'emp_only' | 'ecount_only'
  lots:       LotEntry[]
}

// ── EMP 코드 기본코드 추출 ──────────────────────────────────────────────
function resolveBaseCode(empCode: string, ecountCodes: Set<string>): string {
  if (ecountCodes.has(empCode)) return empCode
  for (const ec of ecountCodes) {
    if (empCode.startsWith(ec + ' ') || empCode.startsWith(ec + '(')) return ec
  }
  const m = empCode.match(/^([^\s(]+)/)
  return m ? m[1] : empCode
}

// ── ① EMP 재고현황 파싱 (.xls) ─────────────────────────────────────────
// 헤더 행 자동 감지, 로케이션 컬럼 활용, ZZ-ZZ → 품명 기반 자동 배정
function parseEmpStock(buf: ArrayBuffer): Map<string, LotEntry[]> {
  const wb = XLSX.read(buf, { type: 'array' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]

  // 헤더 행 탐색 (상품코드 포함 행)
  let headerRow = -1
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const r = rows[i] as unknown[]
    if (r.some(c => String(c ?? '').includes('상품코드') || String(c ?? '').includes('품목코드'))) {
      headerRow = i; break
    }
  }
  if (headerRow < 0) return new Map()

  const header  = (rows[headerRow] as unknown[]).map(h => String(h ?? '').trim())
  const codeIdx = header.findIndex(h => h.includes('상품코드') || h.includes('품목코드'))
  const nameIdx = header.findIndex(h => h.includes('상품명') || h.includes('품목명') || h.includes('품명'))
  const qtyIdx  = header.findIndex(h => h.includes('재고수량') || h === '수량')
  const locoIdx = header.findIndex(h => h.includes('다중로케이션') || h.includes('로케이션') || h.includes('위치코드'))
  if (codeIdx < 0 || qtyIdx < 0) return new Map()

  const result = new Map<string, LotEntry[]>()
  for (const row of rows.slice(headerRow + 1)) {
    const code = String(row[codeIdx] ?? '').trim()
    const qty  = Number(row[qtyIdx]) || 0
    if (!code) continue
    const rawLoc = locoIdx >= 0 ? String(row[locoIdx] ?? '').trim() : ''
    const isZZ   = !rawLoc || rawLoc.startsWith('ZZ-') || rawLoc === '00-DK-00-00'
    let lot: string | null
    let autoLoc = false
    if (isZZ) {
      const name    = nameIdx >= 0 ? String(row[nameIdx] ?? '').trim() : ''
      const assigned = assignLocation(name)
      lot     = assigned ?? 'ZZ (미배정)'
      autoLoc = true
    } else {
      lot = rawLoc
    }
    if (!result.has(code)) result.set(code, [])
    result.get(code)!.push({ lot, qty, autoLoc })
  }
  return result
}

// ── ② 이카운트 수불부 파싱 (.xlsx 여러 개) ──────────────────────────────
// 마지막 '합계' 행의 마지막 컬럼 = 현재고, key=col[1]
function parseEcountStock(bufs: ArrayBuffer[]): Map<string, number> {
  const result = new Map<string, number>()
  for (const buf of bufs) {
    const wb = XLSX.read(buf, { type: 'array' })
    for (const sheetName of wb.SheetNames) {
      const ws   = wb.Sheets[sheetName]
      const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]
      let lastSumRow: unknown[] | null = null
      let code = ''
      for (const row of rows) {
        const c1 = String(row[1] ?? '').trim()
        if (c1 && !code) code = c1
        const c0 = String(row[0] ?? '').trim()
        if (c0 === '합계' || c0.includes('합계')) lastSumRow = row
      }
      if (!code || !lastSumRow) continue
      const stock = Number(lastSumRow[lastSumRow.length - 1]) || 0
      result.set(code, (result.get(code) ?? 0) + stock)
    }
  }
  return result
}

// ── ③ 미출하 현황 파싱 ──────────────────────────────────────────────────
// 시트: 미출하현황, key=col[6], 잔량=col[10]
function parseUnshipped(buf: ArrayBuffer): Map<string, number> {
  const result = new Map<string, number>()
  const wb = XLSX.read(buf, { type: 'array' })
  const ws = wb.Sheets['미출하현황'] ?? wb.Sheets[wb.SheetNames[0]]
  if (!ws) return result
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]
  for (const row of rows) {
    const code = String(row[6] ?? '').trim()
    const qty  = Number(row[10]) || 0
    if (!code || code === '품목코드' || code === '상품코드') continue
    result.set(code, (result.get(code) ?? 0) + qty)
  }
  return result
}

// ── ④ 일일 입출고 확인서 파싱 ──────────────────────────────────────────
// 날짜별 시트, 헤더=행7, 데이터=행8~, col[1]='출고' 행만, key=col[5], 수량=col[8]
// 반환: { shipMap, lastSheetDate }
function parseShipConfirm(buf: ArrayBuffer): { shipMap: Map<string, number>; lastDate: string } {
  const shipMap = new Map<string, number>()
  const wb = XLSX.read(buf, { type: 'array' })
  let lastDate = ''
  for (const sheetName of wb.SheetNames) {
    lastDate = sheetName
    const ws   = wb.Sheets[sheetName]
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][]
    for (const row of rows.slice(7)) {
      const direction = String(row[1] ?? '').trim()
      if (direction !== '출고' && direction !== '국내') continue
      const code = String(row[5] ?? '').trim()
      const qty  = Number(row[8]) || 0
      if (!code) continue
      shipMap.set(code, (shipMap.get(code) ?? 0) + qty)
    }
  }
  return { shipMap, lastDate }
}

// ── 대사 계산 ──────────────────────────────────────────────────────────
function buildReconRows(
  empRaw:    Map<string, LotEntry[]>,
  ecountRaw: Map<string, number>,
  unshipped: Map<string, number>,
  shipMap:   Map<string, number>,
): ReconRow[] {
  const ecountCodes = new Set(ecountRaw.keys())
  // EMP 코드를 기본코드로 집계, 로케이션별 하위 행 보존
  const empByBase = new Map<string, { lots: LotEntry[]; total: number }>()
  for (const [empCode, rows] of empRaw.entries()) {
    const base      = resolveBaseCode(empCode, ecountCodes)
    const lotSuffix = empCode !== base ? empCode.slice(base.length).trim() : null
    if (!empByBase.has(base)) empByBase.set(base, { lots: [], total: 0 })
    const entry = empByBase.get(base)!
    for (const r of rows) {
      const label = [lotSuffix, r.lot].filter(Boolean).join(' / ') || null
      entry.lots.push({ lot: label, qty: r.qty })
      entry.total += r.qty
    }
  }

  // 출고 수불도 기본코드로 집계
  const shipByBase = new Map<string, number>()
  for (const [shipCode, qty] of shipMap.entries()) {
    const base = resolveBaseCode(shipCode, ecountCodes)
    shipByBase.set(base, (shipByBase.get(base) ?? 0) + qty)
  }

  const allCodes = new Set([...empByBase.keys(), ...ecountRaw.keys()])
  const rows: ReconRow[] = []

  for (const code of allCodes) {
    const empEntry  = empByBase.get(code)
    const empRawQty = empEntry?.total ?? 0
    const shipDeduct = shipByBase.get(code) ?? 0
    const empAdj   = empRawQty - shipDeduct
    const ecRaw    = ecountRaw.get(code) ?? 0
    const unship   = unshipped.get(code) ?? 0
    const ecAdj    = ecRaw + unship
    const diff     = empAdj - ecAdj

    let status: ReconRow['status']
    if (!empEntry && ecRaw === 0) continue
    else if (!empEntry) status = 'ecount_only'
    else if (ecRaw === 0 && unship === 0) status = 'emp_only'
    else if (diff === 0) status = 'match'
    else status = 'diff'

    rows.push({
      code, name: '', empRaw: empRawQty, shipDeduct, empAdj,
      ecRaw, unshipped: unship, ecAdj, diff, status,
      lots: empEntry?.lots ?? [],
    })
  }

  rows.sort((a, b) => {
    const order = { diff: 0, emp_only: 1, ecount_only: 2, match: 3 }
    if (order[a.status] !== order[b.status]) return order[a.status] - order[b.status]
    return Math.abs(b.diff) - Math.abs(a.diff)
  })
  return rows
}

// ── Excel 출력 ─────────────────────────────────────────────────────────
async function exportReconExcel(rows: ReconRow[], date: string): Promise<void> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('대사결과')
  const sum = wb.addWorksheet('요약')

  ws.columns = [
    { header: '품목코드', key: 'code',       width: 18 },
    { header: 'EMP재고', key: 'empRaw',      width: 10 },
    { header: '출고차감', key: 'shipDeduct',  width: 10 },
    { header: 'EMP조정', key: 'empAdj',      width: 10 },
    { header: '이카운트', key: 'ecRaw',       width: 10 },
    { header: '미출하',  key: 'unshipped',   width: 10 },
    { header: 'EC조정',  key: 'ecAdj',       width: 10 },
    { header: '차이',    key: 'diff',        width: 8  },
    { header: '상태',    key: 'status',      width: 12 },
  ]
  ws.getRow(1).font = { bold: true }

  const STATUS_FILL: Record<string, ExcelJS.Fill> = {
    diff:        { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE0E0' } },
    emp_only:    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } },
    ecount_only: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3F2FD' } },
  }
  const STATUS_LABEL: Record<string, string> = {
    match: '✅ 일치', diff: '❌ 불일치', emp_only: '⚠ EMP만', ecount_only: '⚠ EC만',
  }

  for (const r of rows) {
    const row = ws.addRow({
      code: r.code, empRaw: r.empRaw, shipDeduct: r.shipDeduct, empAdj: r.empAdj,
      ecRaw: r.ecRaw, unshipped: r.unshipped, ecAdj: r.ecAdj, diff: r.diff,
      status: STATUS_LABEL[r.status],
    })
    const fill = STATUS_FILL[r.status]
    if (fill) row.eachCell(c => { c.fill = fill })
  }

  const total     = rows.length
  const match     = rows.filter(r => r.status === 'match').length
  const diff      = rows.filter(r => r.status === 'diff').length
  const empOnly   = rows.filter(r => r.status === 'emp_only').length
  const ecOnly    = rows.filter(r => r.status === 'ecount_only').length

  sum.addRow(['기준일자', date])
  sum.addRow(['총 품목', total])
  sum.addRow(['✅ 일치', match])
  sum.addRow(['❌ 불일치', diff])
  sum.addRow(['⚠ EMP만', empOnly])
  sum.addRow(['⚠ EC만', ecOnly])

  const buf = await wb.xlsx.writeBuffer()
  downloadXlsx(buf as ArrayBuffer, `재고대사_${date}.xlsx`)
}

// ── UI ────────────────────────────────────────────────────────────────
const STATUS_LABEL: Record<string, string> = {
  match: '✅ 일치', diff: '❌ 불일치', emp_only: '⚠ EMP만', ecount_only: '⚠ EC만',
}
const STATUS_BG: Record<string, string> = {
  diff:        'bg-red-50',
  emp_only:    'bg-yellow-50',
  ecount_only: 'bg-blue-50',
  match:       '',
}

export default function InventoryReconciliationTab() {
  const [empFile,      setEmpFile]      = useState<File | null>(null)
  const [ecountFiles,  setEcountFiles]  = useState<File[]>([])
  const [unshipFile,   setUnshipFile]   = useState<File | null>(null)
  const [shipFile,     setShipFile]     = useState<File | null>(null)

  const [rows,    setRows]    = useState<ReconRow[]>([])
  const [lastDate, setLastDate] = useState('')
  const [running, setRunning] = useState(false)
  const [expanded, setExpanded] = useState<Set<string>>(new Set())
  const [err, setErr] = useState('')

  const toggleExpand = useCallback((code: string) => {
    setExpanded(prev => {
      const next = new Set(prev)
      next.has(code) ? next.delete(code) : next.add(code)
      return next
    })
  }, [])

  async function runRecon() {
    if (!empFile || !shipFile) {
      setErr('EMP 재고현황과 일일 입출고 확인서는 필수입니다.')
      return
    }
    setErr('')
    setRunning(true)
    try {
      const [empBuf, shipBuf] = await Promise.all([
        empFile.arrayBuffer(),
        shipFile.arrayBuffer(),
      ])
      const ecountBufs = await Promise.all(ecountFiles.map(f => f.arrayBuffer()))
      const unshipBuf  = unshipFile ? await unshipFile.arrayBuffer() : null

      const empMap   = parseEmpStock(empBuf)
      const ecMap    = parseEcountStock(ecountBufs)
      const unship   = unshipBuf ? parseUnshipped(unshipBuf) : new Map<string, number>()
      const { shipMap, lastDate: ld } = parseShipConfirm(shipBuf)

      const result = buildReconRows(empMap, ecMap, unship, shipMap)
      setRows(result)
      setLastDate(ld)
    } catch (e) {
      setErr(String(e))
    } finally {
      setRunning(false)
    }
  }

  const total     = rows.length
  const matchCnt  = rows.filter(r => r.status === 'match').length
  const diffCnt   = rows.filter(r => r.status === 'diff').length
  const empOnly   = rows.filter(r => r.status === 'emp_only').length
  const ecOnly    = rows.filter(r => r.status === 'ecount_only').length

  return (
    <div className="space-y-5">
      <div>
        <h2 className="text-xl font-bold text-[#1F3864]">🗂 재고 대사</h2>
        <p className="text-sm text-gray-500 mt-0.5">EMP(생산관리) ↔ 이카운트(ERP) 재고 불일치 즉시 확인</p>
      </div>

      {/* 파일 업로드 패널 */}
      <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
        <FileCard
          label="① EMP 재고현황"
          accept=".xls,.xlsx"
          file={empFile}
          onChange={setEmpFile}
          hint=".xls 단일"
        />
        <FileCard
          label="② 이카운트 수불부"
          accept=".xlsx"
          file={ecountFiles[0] ?? null}
          multiple
          onChange={f => setEcountFiles(f ? [f] : [])}
          onMultipleChange={setEcountFiles}
          hint={`.xlsx ${ecountFiles.length > 1 ? `${ecountFiles.length}개 선택됨` : '여러 개 가능'}`}
        />
        <FileCard
          label="③ 미출하 현황"
          accept=".xlsx"
          file={unshipFile}
          onChange={setUnshipFile}
          hint=".xlsx 단일 (선택)"
          optional
        />
        <FileCard
          label="④ 일일 입출고 확인서"
          accept=".xlsx"
          file={shipFile}
          onChange={setShipFile}
          hint=".xlsx 단일"
        />
      </div>

      <div className="flex gap-3 items-center">
        <button
          onClick={runRecon}
          disabled={running}
          className="px-5 py-2 bg-[#1F3864] text-white text-sm font-medium rounded-lg hover:bg-[#162a4e] disabled:opacity-50 transition-colors"
        >
          {running ? '⏳ 대사 중...' : '🔍 대사 실행'}
        </button>
        {err && <span className="text-sm text-red-500">{err}</span>}
      </div>

      {/* 요약 배너 */}
      {rows.length > 0 && (
        <div className="flex flex-wrap gap-3 items-center bg-gray-50 border border-gray-200 rounded-lg px-4 py-3">
          <span className="text-sm font-medium text-gray-700">총 {total}개 품목</span>
          <span className="text-sm text-green-700">✅ 일치 {matchCnt}</span>
          <span className="text-sm text-red-600">❌ 불일치 {diffCnt}</span>
          <span className="text-sm text-yellow-700">⚠ EMP만 {empOnly}</span>
          <span className="text-sm text-blue-700">⚠ EC만 {ecOnly}</span>
          {lastDate && <span className="text-sm text-gray-400 ml-1">기준: {lastDate}</span>}
          <button
            onClick={() => exportReconExcel(rows, lastDate)}
            className="ml-auto px-4 py-1.5 text-sm bg-white border border-gray-300 rounded-lg hover:bg-gray-100 transition-colors"
          >
            📥 Excel 다운로드
          </button>
        </div>
      )}

      {/* 대사 결과 테이블 */}
      {rows.length > 0 && (
        <div className="overflow-x-auto border border-gray-200 rounded-lg">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-[#1F3864] text-white text-xs">
                <th className="px-3 py-2 text-left">품목코드</th>
                <th className="px-3 py-2 text-right">EMP재고</th>
                <th className="px-3 py-2 text-right">출고차감</th>
                <th className="px-3 py-2 text-right font-bold">EMP조정</th>
                <th className="px-3 py-2 text-right">이카운트</th>
                <th className="px-3 py-2 text-right">미출하</th>
                <th className="px-3 py-2 text-right font-bold">EC조정</th>
                <th className="px-3 py-2 text-right">차이</th>
                <th className="px-3 py-2 text-center">상태</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(r => (
                <>
                  <tr
                    key={r.code}
                    className={`border-t border-gray-100 ${STATUS_BG[r.status]} cursor-pointer hover:brightness-95`}
                    onClick={() => r.lots.length > 1 && toggleExpand(r.code)}
                  >
                    <td className="px-3 py-1.5 font-mono text-xs">
                      {r.lots.length > 1 && (
                        <span className="mr-1 text-gray-400">{expanded.has(r.code) ? '▼' : '▶'}</span>
                      )}
                      {r.code}
                    </td>
                    <td className="px-3 py-1.5 text-right">{r.empRaw.toLocaleString()}</td>
                    <td className="px-3 py-1.5 text-right text-gray-500">{r.shipDeduct > 0 ? `-${r.shipDeduct.toLocaleString()}` : '–'}</td>
                    <td className="px-3 py-1.5 text-right font-medium">{r.empAdj.toLocaleString()}</td>
                    <td className="px-3 py-1.5 text-right">{r.ecRaw.toLocaleString()}</td>
                    <td className="px-3 py-1.5 text-right text-gray-500">{r.unshipped > 0 ? `+${r.unshipped.toLocaleString()}` : '–'}</td>
                    <td className="px-3 py-1.5 text-right font-medium">{r.ecAdj.toLocaleString()}</td>
                    <td className={`px-3 py-1.5 text-right font-bold ${r.diff !== 0 ? 'text-red-600' : 'text-green-600'}`}>
                      {r.diff > 0 ? `+${r.diff}` : r.diff === 0 ? '0' : r.diff}
                    </td>
                    <td className="px-3 py-1.5 text-center text-xs">{STATUS_LABEL[r.status]}</td>
                  </tr>
                  {expanded.has(r.code) && r.lots.map((lot, i) => (
                    <tr key={`${r.code}-lot-${i}`} className={`border-t border-gray-50 ${STATUS_BG[r.status]} opacity-75`}>
                      <td className="px-3 py-1 pl-8 font-mono text-xs">
                        <span className="text-gray-400">└ </span>
                        <span className={lot.autoLoc ? 'text-orange-500 italic' : 'text-gray-500'}>
                          {lot.lot ?? '(기본)'}
                        </span>
                        {lot.autoLoc && (
                          <span className="ml-1 text-[10px] text-orange-400 bg-orange-50 rounded px-1">자동배정</span>
                        )}
                      </td>
                      <td className="px-3 py-1 text-right text-xs text-gray-500">{lot.qty.toLocaleString()}</td>
                      <td colSpan={7} />
                    </tr>
                  ))}
                </>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

// ── FileCard 컴포넌트 ──────────────────────────────────────────────────
interface FileCardProps {
  label:    string
  accept:   string
  file:     File | null
  onChange: (f: File | null) => void
  hint?:    string
  optional?: boolean
  multiple?: boolean
  onMultipleChange?: (files: File[]) => void
}

function FileCard({ label, accept, file, onChange, hint, optional, multiple, onMultipleChange }: FileCardProps) {
  return (
    <div className={`border rounded-lg p-3 flex flex-col gap-2 ${file ? 'border-green-400 bg-green-50' : 'border-gray-200 bg-white'}`}>
      <div className="text-xs font-medium text-gray-700">
        {label} {optional && <span className="text-gray-400">(선택)</span>}
      </div>
      <label className="cursor-pointer">
        <div className="text-xs px-3 py-1.5 bg-white border border-gray-300 rounded text-center hover:bg-gray-50 transition-colors">
          {file ? '✅ ' + file.name.slice(0, 18) + (file.name.length > 18 ? '…' : '') : '파일 선택'}
        </div>
        <input
          type="file"
          accept={accept}
          multiple={multiple}
          className="hidden"
          onChange={e => {
            const files = Array.from(e.target.files ?? [])
            if (multiple && onMultipleChange) {
              onMultipleChange(files)
              onChange(files[0] ?? null)
            } else {
              onChange(files[0] ?? null)
            }
          }}
        />
      </label>
      {hint && <div className="text-[10px] text-gray-400">{hint}</div>}
    </div>
  )
}
