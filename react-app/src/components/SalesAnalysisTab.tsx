import { useState, useMemo, Fragment } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import { classifyOjc, classifyOjcDetailed, COLOR_MAP_OJC_DETAILED } from '../lib/ojcFilter'
import { parseDetailedSalesFile } from '../lib/parse/parseDetailedSales'
import type { DetailedSalesRow } from '../lib/parse/parseDetailedSales'
import { downloadXlsx, pickSaveFile, writeToFileHandle, today } from '../lib/download'
import FileUploader from './FileUploader'

// ── 상수 ────────────────────────────────────────────────────────────
const MONTHS = ['01','02','03','04','05','06','07','08','09','10','11','12']

// OJC 외 품목 카테고리 분류 (ESZ018R 실제 품목 기반)
function classifyEtc(name: string): string | null {
  const n = name.toUpperCase()
  // ── 영문 prefix 우선 ──────────────────────────────────────
  if (n.startsWith('DOFC'))              return 'DOFC (돔형광통합체)'
  if (n.startsWith('BOFC'))             return 'BOFC (광접속함체)'
  if (n.startsWith('FAOC'))             return 'FAOC (현장조립커넥터)'
  if (n.startsWith('MOFD'))             return 'MOFD'
  if (n.startsWith('AJW-RN'))          return 'AJW-RN 스플리터'
  if (n.startsWith('BBX'))             return 'BBX'
  if (n.startsWith('OFD'))             return 'OFD (광분배함)'
  if (n.startsWith('SOC'))             return '생산자재'
  if (n.startsWith('IJP'))             return 'IJP'
  if (n.startsWith('DISTRIBUTION CABLE')) return 'DISTRIBUTION CABLE'
  if (n.startsWith('DROP OPTICAL CABLE')) return '생산자재'
  if (n.startsWith('RIBBON FAN-OUT'))  return '리본팬아웃'
  // ── 한/영 키워드 ─────────────────────────────────────────
  if (/ADAPTER|OFA/.test(n) || name.includes('어댑터'))           return 'ADAPTER (어댑터)'
  if (name.includes('광접속함체'))                                  return 'BOFC (광접속함체)'
  if (name.includes('광분배함'))                                    return 'OFD (광분배함)'
  if (name.includes('광통합체') || name.includes('돔형'))           return 'DOFC (돔형광통합체)'
  if (name.includes('광감쇠기'))                                    return '광감쇠기'
  if (name.includes('심선접속자'))                                  return '심선접속자'
  if (name.includes('스플리터') || /SPLITTER/.test(n))             return '스플리터'
  if (name.includes('광커플러') || /COUPLER/.test(n))              return '광커플러'
  if (name.includes('인입광케이블'))                                return '인입광케이블'
  if (name.includes('광섬유보호') || name.includes('보호튜브') || name.includes('열수축')) return '광섬유보호재'
  if (/HOUSING|HOUSING KIT/.test(n) || name.includes('하우징'))    return '생산자재'
  if (/FERRULE/.test(n) || name.includes('페롤'))                  return '생산자재'
  if (/BOOT/.test(n) || name.includes('부트'))                     return '생산자재'
  if (/UTP|모듈러/.test(n) || name.includes('UTP') || name.includes('모듈러')) return 'UTP/통신자재'
  if (/TOOL|STRIPPER|CLEAVER|CUTTER|METER|공구|측정기/.test(n))   return '공구/측정기'
  return null
}

const num      = (v: number) => v === 0 ? '—' : v.toLocaleString()
const dec1     = (v: number) => v.toFixed(1)
const fmtPrice = (v: number) => v === 0 ? '—' : v.toLocaleString() + '원'

function isTop3Excluded(name: string): boolean {
  const u = name.toUpperCase()
  return u.includes('PROTECTION SLEEVE') || u.includes('운송비')
}

type SubView = 'summary' | 'ojc' | 'customer_domestic' | 'customer_foreign' | 'ojc_customer' | 'full'

// ── 타입 ─────────────────────────────────────────────────────────────
interface ProductData {
  code:          string
  annuals:       Record<string, number>
  priceAnnuals:  Record<string, number>
  monthlyLatest: Record<string, number>
  monthlyByYear: Record<string, Record<string, number>>
}
interface CategoryData {
  products:      Record<string, ProductData>
  annuals:       Record<string, number>
  priceAnnuals:  Record<string, number>
  monthlyLatest: Record<string, number>
  monthlyByYear: Record<string, Record<string, number>>  // year → month → qty
}

// ── 메인 컴포넌트 ─────────────────────────────────────────────────────
export default function SalesAnalysisTab() {
  const [rawRows, setRawRows]   = useState<DetailedSalesRow[]>([])
  const [logs, setLogs]         = useState<string[]>([])
  const [file, setFile]         = useState<File | null>(null)
  const [invFile, setInvFile]   = useState<File | null>(null)
  const [running, setRunning]   = useState(false)
  const [dlProgress, setDlProgress] = useState<{ current: number; total: number; label: string } | null>(null)
  const [invLoaded, setInvLoaded] = useState(false)

  // ESZ018R 파싱 상태
  interface InvRow { code: string; name: string; total: number; byWh: Record<string, number> }
  const [eszRows,      setEszRows]      = useState<InvRow[] | null>(null)
  const [eszWarehouses, setEszWarehouses] = useState<string[]>([])
  const [eszSelectedWh, setEszSelectedWh] = useState('')
  const [eszSubTab,    setEszSubTab]    = useState<'total' | 'warehouse'>('total')
  const [subView, setSubView]   = useState<SubView>('summary')
  const [ojcStock, setOjcStock] = useState<Record<string, number>>(() => {
    try { return JSON.parse(localStorage.getItem('ojc_stock') ?? '{}') } catch { return {} }
  })

  const run = async () => {
    if (!file) return alert('판매현황 파일을 선택해주세요.')
    setRunning(true)
    const newLogs: string[] = []
    try {
      const rows = parseDetailedSalesFile(await file.arrayBuffer(), newLogs)
      setRawRows(rows)
    } catch (e) {
      newLogs.push(`❌ 오류: ${e}`)
    } finally {
      setLogs(newLogs)
      setRunning(false)
    }
  }

  // 파일 선택 시 ESZ018R 여부 자동 감지 후 파싱
  const handleInvFile = async (f: File) => {
    setInvFile(f)
    setEszRows(null)
    setInvLoaded(false)
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' })
      if (!wb.Sheets['재고현황']) return   // ESZ018R 아님 → 기존 방식 사용
      const ws = wb.Sheets['재고현황']
      const range = XLSX.utils.decode_range(ws['!ref']!)
      // 2행(index 1) = 헤더
      const header: string[] = []
      for (let c = 0; c <= range.e.c; c++) {
        header.push(ws[XLSX.utils.encode_cell({ r: 1, c })]?.v?.toString().trim() ?? '')
      }
      const whCols = header.slice(4).filter(h => h)
      setEszWarehouses(whCols)
      setEszSelectedWh(whCols[0] ?? '')
      // 3행(index 2)~ 파싱
      const rows: { code: string; name: string; total: number; byWh: Record<string, number> }[] = []
      for (let r = 2; r <= range.e.r; r++) {
        const code  = ws[XLSX.utils.encode_cell({ r, c: 0 })]?.v?.toString().trim() ?? ''
        const name  = ws[XLSX.utils.encode_cell({ r, c: 1 })]?.v?.toString().trim() ?? ''
        const total = Number(ws[XLSX.utils.encode_cell({ r, c: 3 })]?.v ?? 0)
        if (!code) continue
        const byWh: Record<string, number> = {}
        whCols.forEach((wh, wi) => {
          byWh[wh] = Number(ws[XLSX.utils.encode_cell({ r, c: 4 + wi })]?.v ?? 0)
        })
        rows.push({ code, name, total, byWh })
      }
      setEszRows(rows)
    } catch (e) { console.error('ESZ018R 파싱 오류', e) }
  }

  // ESZ018R → ojcStock 카테고리별 합산 적용 (OJC + 기타 포함)
  const applyEszStock = (useTotal: boolean) => {
    if (!eszRows) return
    const next: Record<string, number> = {}
    for (const row of eszRows) {
      const cat = classifyOjc(row.name) ?? classifyEtc(row.name)
      if (!cat) continue
      const qty = useTotal ? row.total : (row.byWh[eszSelectedWh] ?? 0)
      next[cat] = (next[cat] ?? 0) + qty
    }
    const merged = { ...ojcStock, ...next }
    setOjcStock(merged)
    localStorage.setItem('ojc_stock', JSON.stringify(merged))
    setInvLoaded(true)
    setTimeout(() => setInvLoaded(false), 3000)
  }

  // 기존 단순 형식 재고 반영
  const loadInventory = async () => {
    if (!invFile) return alert('재고 파일을 선택해주세요.')
    try {
      const wb  = XLSX.read(await invFile.arrayBuffer(), { type: 'array' })
      const ws  = wb.Sheets[wb.SheetNames[0]]
      const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null })
      const next = { ...ojcStock }
      for (const row of raw) {
        const qty = parseInt(String(
          row['현재고'] ?? row['재고'] ?? row['재고수량'] ?? row['수량'] ?? '0'
        )) || 0
        if (qty <= 0) continue
        const catName = String(row['카테고리'] ?? row['분류'] ?? row['구분'] ?? '').trim()
        if (catName) { next[catName] = qty; continue }
        const prodName = String(row['품목명'] ?? row['제품명'] ?? row['품목'] ?? '').trim()
        if (prodName) {
          const cat = classifyOjc(prodName)
          if (cat) next[cat] = (next[cat] ?? 0) + qty
        }
      }
      setOjcStock(next)
      localStorage.setItem('ojc_stock', JSON.stringify(next))
      setInvLoaded(true)
      setTimeout(() => setInvLoaded(false), 3000)
    } catch (e) { alert(`재고 파일 오류: ${e}`) }
  }

  const updateStock = (cat: string, val: string) => {
    const next = { ...ojcStock, [cat]: parseInt(val) || 0 }
    setOjcStock(next)
    localStorage.setItem('ojc_stock', JSON.stringify(next))
  }

  const years    = useMemo(() => [...new Set(rawRows.map(r => r.year))].sort(), [rawRows])
  const latestYr = years[years.length - 1] ?? ''

  // ── 전체 품목별 집계 (단일 소스) ──────────────────────────────
  // rawRows를 한 번만 순회해 품목별 연간·월별 데이터를 모두 담음
  const fullProducts = useMemo(() => {
    const map: Record<string, {
      name: string; code: string
      ojcCat: string | null; etcCat: string | null
      annuals: Record<string, number>
      priceAnnuals: Record<string, number>
      monthlyByYear: Record<string, Record<string, number>>
    }> = {}
    for (const row of rawRows) {
      if (!map[row.name]) map[row.name] = {
        name: row.name, code: row.code,
        ojcCat: classifyOjc(row.name), etcCat: classifyEtc(row.name),
        annuals: {}, priceAnnuals: {}, monthlyByYear: {},
      }
      const p = map[row.name]
      p.annuals[row.year]      = (p.annuals[row.year] ?? 0) + row.qty
      p.priceAnnuals[row.year] = (p.priceAnnuals[row.year] ?? 0) + row.price
      if (!p.monthlyByYear[row.year]) p.monthlyByYear[row.year] = {}
      p.monthlyByYear[row.year][row.month] = (p.monthlyByYear[row.year][row.month] ?? 0) + row.qty
    }
    return Object.values(map).sort((a, b) =>
      (b.annuals[latestYr] ?? 0) - (a.annuals[latestYr] ?? 0)
    )
  }, [rawRows, latestYr])

  // ── 카테고리별 집계 (fullProducts에서 파생) ─────────────────
  function buildCategoryMap(
    products: typeof fullProducts,
    getCat: (p: typeof fullProducts[number]) => string | null,
    latestYr: string,
  ): Record<string, CategoryData> {
    const cats: Record<string, CategoryData> = {}
    for (const p of products) {
      const cat = getCat(p)
      if (!cat) continue
      if (!cats[cat]) cats[cat] = { products: {}, annuals: {}, priceAnnuals: {}, monthlyLatest: {}, monthlyByYear: {} }
      const c = cats[cat]
      const monthlyLatest = p.monthlyByYear[latestYr] ?? {}
      c.products[p.name] = {
        code: p.code,
        annuals: p.annuals,
        priceAnnuals: p.priceAnnuals,
        monthlyLatest,
        monthlyByYear: p.monthlyByYear,
      }
      for (const [yr, qty] of Object.entries(p.annuals)) {
        c.annuals[yr] = (c.annuals[yr] ?? 0) + qty
      }
      for (const [yr, price] of Object.entries(p.priceAnnuals)) {
        c.priceAnnuals[yr] = (c.priceAnnuals[yr] ?? 0) + price
      }
      for (const [yr, months] of Object.entries(p.monthlyByYear)) {
        if (!c.monthlyByYear[yr]) c.monthlyByYear[yr] = {}
        for (const [mo, qty] of Object.entries(months)) {
          c.monthlyByYear[yr][mo] = (c.monthlyByYear[yr][mo] ?? 0) + qty
        }
      }
      for (const [mo, qty] of Object.entries(monthlyLatest)) {
        c.monthlyLatest[mo] = (c.monthlyLatest[mo] ?? 0) + qty
      }
    }
    return cats
  }

  const ojcByCategory = useMemo(
    () => buildCategoryMap(fullProducts, p => p.ojcCat, latestYr),
    [fullProducts, latestYr],
  )
  const etcByCategory = useMemo(
    () => buildCategoryMap(fullProducts, p => p.etcCat, latestYr),
    [fullProducts, latestYr],
  )

  // OJC 행 (거래처 탑3용으로만 사용)
  const ojcRows = useMemo(() => rawRows.filter(r => classifyOjc(r.name) !== null), [rawRows])

  // 거래처별 탑3 집계 헬퍼 (내자/외자 공통)
  function buildCustomerTop3(rows: typeof rawRows) {
    const custMap: Record<string, Record<string, { code: string; qty: number; price: number }>> = {}
    for (const row of rows) {
      if (isTop3Excluded(row.name)) continue
      const cust = row.customer || '(미상)'
      if (!custMap[cust]) custMap[cust] = {}
      if (!custMap[cust][row.name]) custMap[cust][row.name] = { code: row.code, qty: 0, price: 0 }
      custMap[cust][row.name].qty   += row.qty
      custMap[cust][row.name].price += row.price
    }
    const result: Record<string, Array<{ name: string; code: string; qty: number; price: number }>> = {}
    for (const [cust, products] of Object.entries(custMap)) {
      result[cust] = Object.entries(products)
        .sort((a, b) => b[1].qty - a[1].qty)
        .slice(0, 3)
        .map(([name, d]) => ({ name, code: d.code, qty: d.qty, price: d.price }))
    }
    return Object.fromEntries(
      Object.entries(result).sort((a, b) =>
        b[1].reduce((s, x) => s + x.qty, 0) - a[1].reduce((s, x) => s + x.qty, 0)
      )
    )
  }

  // 내자/외자 분리 집계
  const domesticRows = useMemo(() => rawRows.filter(r => !r.isForeign), [rawRows])
  const foreignRows  = useMemo(() => rawRows.filter(r => r.isForeign),  [rawRows])

  const customerTop3          = useMemo(() => buildCustomerTop3(rawRows),     [rawRows])
  const customerTop3Domestic  = useMemo(() => buildCustomerTop3(domesticRows), [domesticRows])
  const customerTop3Foreign   = useMemo(() => buildCustomerTop3(foreignRows),  [foreignRows])

  const hasData = rawRows.length > 0

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>판매 현황 분석</b>: OJC 판매량(연간·월간·피크·재고커버리지), 거래처별 탑3,
        전체 품목 판매량을 조회합니다.<br />
        필요 컬럼: <code className="bg-gray-100 px-1 rounded">거래처명 | 품목코드 | 품목명 | 년 | 월 | 수량</code>
      </div>

      {/* 파일 업로드 */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        {/* ① 판매현황 */}
        <div className="flex gap-2 items-end">
          <div className="flex-1">
            <FileUploader
              label="① 대외비_판매현황.xlsx — 거래처명, 품목코드, 품목명, 년, 월, 수량"
              fileName={file?.name ?? ''}
              onFile={setFile}
            />
          </div>
          <button
            onClick={run}
            disabled={!file || running}
            className="bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold px-4 py-2 rounded transition text-sm h-[42px] whitespace-nowrap"
          >
            {running ? '⏳...' : '▶ 실행'}
          </button>
        </div>

        {/* ② 재고 파일 */}
        <div className="space-y-2">
          <div className="flex gap-2 items-end">
            <div className="flex-1">
              <FileUploader
                label="② OJC 완제품 재고.xlsx — ESZ018R 또는 카테고리+현재고 형식"
                fileName={invFile?.name ?? ''}
                onFile={handleInvFile}
                optional
              />
            </div>
            {/* ESZ018R 아닌 경우에만 기존 버튼 표시 */}
            {invFile && !eszRows && (
              <button
                onClick={loadInventory}
                className="bg-[#2E75B6] hover:bg-[#1a5a9e] text-white font-semibold px-4 py-2 rounded transition text-sm h-[42px] whitespace-nowrap"
              >
                재고 반영
              </button>
            )}
          </div>

          {/* ESZ018R 감지 시 → 서브탭 */}
          {eszRows && (
            <div className="border border-[#2E75B6] rounded-lg overflow-hidden">
              {/* 서브탭 헤더 */}
              <div className="flex border-b border-gray-200 bg-gray-50">
                {([['total', '총수량 (D열)'], ['warehouse', '창고별']] as const).map(([id, lbl]) => (
                  <button
                    key={id}
                    onClick={() => setEszSubTab(id)}
                    className={`px-4 py-2 text-sm font-medium border-b-2 -mb-px transition-colors ${
                      eszSubTab === id
                        ? 'border-[#2E75B6] text-[#2E75B6] bg-white'
                        : 'border-transparent text-gray-500 hover:text-gray-700'
                    }`}
                  >
                    {lbl}
                  </button>
                ))}
                <div className="ml-auto px-3 py-2 text-xs text-gray-400 self-center">
                  {eszRows.filter(r => classifyOjc(r.name)).length}건 매칭 / 전체 {eszRows.length}건
                </div>
              </div>

              <div className="p-3 space-y-2">
                {/* 창고 선택 */}
                {eszSubTab === 'warehouse' && (
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-medium text-gray-600">창고 선택</span>
                    <select
                      value={eszSelectedWh}
                      onChange={e => setEszSelectedWh(e.target.value)}
                      className="border border-gray-300 rounded px-2 py-1 text-xs bg-white focus:outline-none focus:border-blue-400"
                    >
                      {eszWarehouses.map(wh => <option key={wh} value={wh}>{wh}</option>)}
                    </select>
                  </div>
                )}

                {/* 카테고리별 미리보기 */}
                <div className="grid grid-cols-3 gap-1.5 text-xs">
                  {Object.entries(
                    eszRows.reduce<Record<string, number>>((acc, row) => {
                      const cat = classifyOjc(row.name)
                      if (!cat) return acc
                      const qty = eszSubTab === 'total' ? row.total : (row.byWh[eszSelectedWh] ?? 0)
                      acc[cat] = (acc[cat] ?? 0) + qty
                      return acc
                    }, {})
                  ).map(([cat, qty]) => (
                    <div key={cat} className="flex justify-between bg-white border border-gray-100 rounded px-2 py-1">
                      <span className="font-medium text-gray-700">{cat}</span>
                      <span className="font-mono text-[#2E75B6] font-bold">{qty.toLocaleString()}</span>
                    </div>
                  ))}
                </div>

                <button
                  onClick={() => applyEszStock(eszSubTab === 'total')}
                  className="w-full bg-[#2E75B6] hover:bg-[#1a5a9e] text-white font-semibold py-1.5 rounded text-sm transition"
                >
                  {eszSubTab === 'total' ? '총수량으로 현재고 반영' : `${eszSelectedWh} 수량으로 현재고 반영`}
                </button>
              </div>
            </div>
          )}
        </div>
      </div>

      {invLoaded && (
        <div className="bg-[#D6F0D8] px-4 py-2 rounded text-sm font-semibold text-[#1a6a2a]">
          ✅ 재고 데이터가 반영됐습니다.
        </div>
      )}

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs leading-5 max-h-48 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {/* 서브탭 + 통합 다운로드 */}
      {hasData && (
        <div className="flex gap-2 flex-wrap items-center justify-between">
          <div className="flex gap-2 flex-wrap">
            {(['summary', 'ojc', 'customer_domestic', 'customer_foreign', 'ojc_customer', 'full'] as SubView[]).map(v => (
              <button
                key={v}
                onClick={() => setSubView(v)}
                className={`px-4 py-1.5 text-sm font-medium rounded border transition ${
                  subView === v
                    ? 'bg-yellow-400 border-yellow-500 text-black'
                    : 'bg-white border-gray-300 text-gray-600 hover:bg-gray-50'
                }`}
              >
                {v === 'summary'           ? '⓪ 요약 대시보드'
                  : v === 'ojc'            ? '① OJC 판매 현황'
                  : v === 'customer_domestic' ? '② 거래처별 탑3 (내자)'
                  : v === 'customer_foreign'  ? '③ 거래처별 탑3 (외자)'
                  : v === 'ojc_customer'   ? '④ 거래처별 OJC 탑3'
                  : '⑤ 전체 판매량'}
              </button>
            ))}
          </div>
          {dlProgress ? (
            <div className="flex flex-col gap-1 min-w-[220px]">
              <div className="flex justify-between text-xs text-gray-600">
                <span className="font-medium">{dlProgress.label}</span>
                <span className="text-gray-400">{dlProgress.current} / {dlProgress.total}</span>
              </div>
              <div className="w-full bg-gray-200 rounded-full h-2 overflow-hidden">
                <div
                  className="bg-[#2E75B6] h-2 rounded-full transition-all duration-200"
                  style={{ width: `${Math.round((dlProgress.current / dlProgress.total) * 100)}%` }}
                />
              </div>
            </div>
          ) : (
            <button
              onClick={async () => {
                const filename = `판매현황분석_${today()}.xlsx`
                // user gesture 컨텍스트가 살아있는 지금 파일 핸들 먼저 획득
                const handle = await pickSaveFile(filename)
                if (handle === 'cancelled') return
                downloadStyledExcel(
                  ojcByCategory, etcByCategory, fullProducts, years, latestYr, ojcStock, rawRows,
                  (current, total, label) => setDlProgress({ current, total, label }),
                  handle,
                ).catch(e => alert(`다운로드 오류: ${e}`))
                 .finally(() => setDlProgress(null))
              }}
              className="px-4 py-1.5 text-sm bg-[#2E75B6] hover:bg-[#1a5a9e] text-white font-semibold rounded transition whitespace-nowrap"
            >
              📥 전체 통합 다운로드 (8시트)
            </button>
          )}
        </div>
      )}

      {hasData && subView === 'summary' && (
        <SalesSummaryView
          ojcByCategory={ojcByCategory}
          etcByCategory={etcByCategory}
          customerTop3={customerTop3}
          fullProducts={fullProducts}
          years={years}
          latestYr={latestYr}
        />
      )}
      {hasData && subView === 'ojc' && (
        <OjcSalesView
          ojcByCategory={ojcByCategory}
          etcByCategory={etcByCategory}
          years={years}
          latestYr={latestYr}
          ojcStock={ojcStock}
          onStockChange={updateStock}
        />
      )}
      {hasData && subView === 'customer_domestic' && (
        <CustomerTop3View customerTop3={customerTop3Domestic} years={years} rawRows={domesticRows} label="내자" />
      )}
      {hasData && subView === 'customer_foreign' && (
        <CustomerTop3View customerTop3={customerTop3Foreign} years={years} rawRows={foreignRows} label="외자" />
      )}
      {hasData && subView === 'ojc_customer' && (
        <OjcCustomerTop3View ojcRows={ojcRows} years={years} />
      )}
      {hasData && subView === 'full' && (
        <FullSalesView products={fullProducts} years={years} latestYr={latestYr} />
      )}

      {!hasData && (
        <div className="text-center text-gray-400 py-16 text-sm">
          📊 판매현황 파일을 업로드하고 분석 실행 버튼을 누르세요.
        </div>
      )}
    </div>
  )
}

// ── 월간 카테고리 맵 빌더 (OJC 세분화·기타 공용) ──────────────────────
function buildMonthlyMap(
  rows: DetailedSalesRow[],
  classify: (name: string) => string | null,
): Record<string, CategoryData> {
  const cats: Record<string, CategoryData> = {}
  for (const row of rows) {
    const cat = classify(row.name)
    if (!cat) continue
    if (!cats[cat]) cats[cat] = { products: {}, annuals: {}, priceAnnuals: {}, monthlyLatest: {}, monthlyByYear: {} }
    const c = cats[cat]
    if (!c.products[row.name]) c.products[row.name] = { code: row.code, annuals: {}, priceAnnuals: {}, monthlyLatest: {}, monthlyByYear: {} }
    const p = c.products[row.name]
    p.annuals[row.year]      = (p.annuals[row.year] ?? 0) + row.qty
    p.priceAnnuals[row.year] = (p.priceAnnuals[row.year] ?? 0) + row.price
    if (!p.monthlyByYear[row.year]) p.monthlyByYear[row.year] = {}
    p.monthlyByYear[row.year][row.month] = (p.monthlyByYear[row.year][row.month] ?? 0) + row.qty
    c.annuals[row.year]      = (c.annuals[row.year] ?? 0) + row.qty
    c.priceAnnuals[row.year] = (c.priceAnnuals[row.year] ?? 0) + row.price
    if (!c.monthlyByYear[row.year]) c.monthlyByYear[row.year] = {}
    c.monthlyByYear[row.year][row.month] = (c.monthlyByYear[row.year][row.month] ?? 0) + row.qty
  }
  return cats
}

const OJC_DETAIL_ORDER = [
  'KT OJC-SP', 'KT OJC-DP', 'KT OJC-MOJC', 'KT OJC',
  'LG OJC-SP', 'LG OJC-DP', 'LG OJC-MOJC',
  'DROP-1C', 'DROP-2C', 'DROP',
  '피그테일-SK 성단용', '피그테일', 'Optical Cable Parts', 'DX-MM',
]

// ── 통합 판매현황 다운로드 (ExcelJS 스타일) ───────────────────────────
async function downloadStyledExcel(
  ojcByCategory: Record<string, CategoryData>,
  etcByCategory: Record<string, CategoryData>,
  fullProducts: Array<{ name: string; code: string; ojcCat: string | null; etcCat: string | null; annuals: Record<string, number>; priceAnnuals: Record<string, number> }>,
  years: string[],
  latestYr: string,
  ojcStock: Record<string, number>,
  rawRows: Array<{ customer: string; code: string; name: string; year: string; month: string; qty: number; price: number; isForeign: boolean }>,
  onProgress: (current: number, total: number, label: string) => void,
  fileHandle: unknown = null,
) {
  const TOTAL = 9  // 시트 8개 + 파일 저장 1단계
  const prog = async (step: number, label: string) => {
    onProgress(step, TOTAL, label)
    await new Promise(r => setTimeout(r, 0))
  }
  const wb = new ExcelJS.Workbook()
  wb.creator = 'AJW SCM팀'
  wb.created = new Date()

  const C = {
    darkBlue: 'FF1F3864', midBlue: 'FF2E75B6', lightBlue: 'FFD6E4F0', altBlue: 'FFE9F2F9',
    purple: 'FF7030A0', lightPurple: 'FFF3EEF8', altPurple: 'FFF8F3FC',
    orange: 'FFC55A11', green: 'FF375623', yellow: 'FF7F6000', red: 'FF9C0006',
    white: 'FFFFFFFF', gray1: 'FFF5F5F5', gray2: 'FFD9D9D9', gray3: 'FF595959', infoBg: 'FFF0F4FA',
  }
  const coverFg = (v: number) => v >= 2 ? C.green : v >= 1 ? C.yellow : C.red

  function peakOfYear(data: CategoryData, yr: string) {
    const e = Object.entries(data.monthlyByYear[yr] ?? {})
      .reduce<[string, number]>((b, [m, v]) => v > b[1] ? [m, v] : b, ['', 0])
    return { month: e[0] ? `${parseInt(e[0])}월` : '', val: e[1] }
  }

  function addTitle(ws: ExcelJS.Worksheet, title: string, sub: string, cols: number) {
    ws.addRow([title]); ws.mergeCells(1, 1, 1, cols)
    const t = ws.getCell(1, 1)
    t.value = title; t.font = { bold: true, size: 13, color: { argb: C.white } }
    t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.darkBlue } }
    t.alignment = { horizontal: 'center', vertical: 'middle' }; ws.getRow(1).height = 30
    ws.addRow([sub]); ws.mergeCells(2, 1, 2, cols)
    const s = ws.getCell(2, 1)
    s.value = sub; s.font = { size: 9, color: { argb: C.gray3 } }
    s.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.infoBg } }
    s.alignment = { horizontal: 'left', vertical: 'middle' }; ws.getRow(2).height = 14
  }

  function styleHdr(row: ExcelJS.Row, bg = C.midBlue, h = 28) {
    row.height = h
    row.eachCell({ includeEmpty: true }, c => {
      c.font = { bold: true, size: 10, color: { argb: C.white } }
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      c.border = { bottom: { style: 'medium', color: { argb: C.darkBlue } }, right: { style: 'thin', color: { argb: C.gray2 } } }
    })
  }

  function addSep(ws: ExcelJS.Worksheet, text: string, cols: number, fg = C.purple, bg = C.lightPurple) {
    const row = ws.addRow([text]); ws.mergeCells(row.number, 1, row.number, cols)
    const c = ws.getCell(row.number, 1)
    c.value = text; c.font = { bold: true, size: 10, color: { argb: fg } }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
    c.alignment = { horizontal: 'left', vertical: 'middle' }
    c.border = { top: { style: 'medium', color: { argb: fg } } }; row.height = 20
  }


  await prog(1, '① 판매현황_요약')
  // ── Sheet 1: 판매현황_요약 대시보드 ──────────────────────────────────────
  const domesticRowsForExcel = rawRows.filter(r => !r.isForeign)
  const foreignRowsForExcel  = rawRows.filter(r => r.isForeign)

  const ws1 = wb.addWorksheet('① 판매현황_요약')
  ws1.views = [{ state: 'frozen', ySplit: 2 }]

  // cat(1) + count(1) + years(N) + yoy(1) + cagr(1) + stock(1) + cover_pk(1) + cover_avg(1) = N+7
  const S1     = 2 + years.length + 5
  const prevYr = years[years.length - 2] ?? ''

  const sumQty   = (rows: typeof rawRows, yr: string) => rows.filter(r => r.year === yr).reduce((s, r) => s + r.qty, 0)
  const yoyRate  = (cur: number, pre: number): number | null => pre > 0 ? (cur - pre) / pre : null
  const cagrRate = (first: number, last: number, n: number): number | null =>
    n >= 2 && first > 0 && last > 0 ? Math.pow(last / first, 1 / (n - 1)) - 1 : null

  const allCusts  = [...new Set(rawRows.map(r => r.customer).filter(Boolean))].length
  const domCusts  = [...new Set(domesticRowsForExcel.map(r => r.customer).filter(Boolean))].length
  const forCusts  = [...new Set(foreignRowsForExcel.map(r => r.customer).filter(Boolean))].length
  addTitle(ws1,
    `AJW 판매현황 요약 대시보드  ·  20${years[0]}년 ~ 20${latestYr}년`,
    `생성: ${today()}  |  ${fullProducts.length}종 품목  |  거래처 ${allCusts}개 (내자 ${domCusts} / 외자 ${forCusts})`, S1)

  const addSecHdr1 = (label: string) => {
    const r = ws1.addRow([label])
    ws1.mergeCells(r.number, 1, r.number, S1)
    const c = ws1.getCell(r.number, 1)
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.darkBlue } }
    c.font = { bold: true, size: 10, color: { argb: C.white } }
    c.alignment = { horizontal: 'left', vertical: 'middle' }
    r.height = 22
  }

  // ── [A] KPI 블록 ─────────────────────────────────────────────────────────
  addSecHdr1('  ▶ 전체 판매 현황 요약')

  // 구분(1) + 연도별(N) + 전년비(0 or 1) + 거래처수(1) + 품목수(1)
  const kpiYoyCols  = prevYr ? 1 : 0
  const kpiFixedCols = 1 + years.length + kpiYoyCols + 2
  const kpiHdr = ws1.addRow([
    '구분',
    ...years.map(y => `20${y}년\n판매(EA)`),
    ...(prevYr ? [`20${latestYr}년/\n20${prevYr}년 전년비`] : []),
    '거래처 수', '품목 수',
    ...Array(Math.max(0, S1 - kpiFixedCols)).fill(null),
  ])
  styleHdr(kpiHdr, C.midBlue, 22)
  if (S1 > kpiFixedCols) ws1.mergeCells(kpiHdr.number, kpiFixedCols, kpiHdr.number, S1)

  const kpiYoyCol = prevYr ? 2 + years.length : -1
  ;[
    { label: '전체',      rows: rawRows,                custs: allCusts, items: fullProducts.length },
    { label: '내자 🏠',  rows: domesticRowsForExcel,   custs: domCusts,
      items: [...new Set(domesticRowsForExcel.map(r => r.name))].length },
    { label: '외자 ✈️', rows: foreignRowsForExcel,    custs: forCusts,
      items: [...new Set(foreignRowsForExcel.map(r => r.name))].length },
  ].forEach((kd, ki) => {
    const yqts  = years.map(yr => sumQty(kd.rows, yr))
    const latQ  = yqts[yqts.length - 1] ?? 0
    const preQ  = yqts[yqts.length - 2] ?? 0
    const yoyV  = yoyRate(latQ, preQ)
    const row   = ws1.addRow([
      kd.label,
      ...yqts.map(q => q || null),
      ...(prevYr ? [yoyV] : []),
      kd.custs || null, kd.items || null,
      ...Array(Math.max(0, S1 - kpiFixedCols)).fill(null),
    ])
    if (S1 > kpiFixedCols) ws1.mergeCells(row.number, kpiFixedCols, row.number, S1)
    row.height = 20
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ki === 0 ? C.altBlue : C.white } }
      c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ci === 1) {
        c.font = { bold: true, color: { argb: C.darkBlue } }; c.alignment = { horizontal: 'left', vertical: 'middle' }
      } else if (ci === kpiYoyCol) {
        c.alignment = { horizontal: 'right', vertical: 'middle' }
        if (typeof c.value === 'number') {
          c.numFmt = '+0.0%;-0.0%;"-"'
          c.font = { bold: true, color: { argb: c.value > 0 ? C.green : c.value < 0 ? C.red : C.gray3 } }
        }
      } else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
    })
  })

  // ── [B] OJC 카테고리별 판매 추이 ─────────────────────────────────────────
  addSecHdr1('  ▶ OJC 판매 현황')
  styleHdr(ws1.addRow([
    '카테고리', '품목수', ...years.map(y => `20${y}년\n판매(EA)`),
    prevYr ? `20${latestYr}년/\n20${prevYr}년 전년비` : '전년비',
    '연평균\n성장률(CAGR)', '현재고\n(EA)', '재고여유\n(최고월)', '재고여유\n(평균)',
  ]), C.midBlue, 32)

  const applyTrendRow = (row: ExcelJS.Row, bg: string, fontColor: string) => {
    const yoyCol  = 3 + years.length
    const cagrCol = 4 + years.length
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ci === 1) {
        c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { bold: true, color: { argb: fontColor } }
      } else if (ci === 2) {
        c.alignment = { horizontal: 'center', vertical: 'middle' }
      } else if (ci === yoyCol || ci === cagrCol) {
        c.alignment = { horizontal: 'right', vertical: 'middle' }
        if (typeof c.value === 'number') {
          c.numFmt = '+0.0%;-0.0%;"-"'
          c.font = { bold: true, color: { argb: c.value > 0 ? C.green : c.value < 0 ? C.red : C.gray3 } }
        }
      } else if (ci === S1 - 1 || ci === S1) {
        c.alignment = { horizontal: 'right', vertical: 'middle' }
        if (typeof c.value === 'number') {
          c.numFmt = '0.0"개월"'
          c.font = { bold: true, color: { argb: coverFg(c.value as number) } }
        }
      } else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
    })
  }

  let ri1 = 0
  for (const [cat, data] of Object.entries(ojcByCategory)) {
    const firstQty  = data.annuals[years[0]] ?? 0
    const latestQty = data.annuals[latestYr]  ?? 0
    const prevQty   = data.annuals[prevYr]    ?? 0
    const mv        = MONTHS.map(m => data.monthlyLatest[m] ?? 0)
    const pk        = Math.max(0, ...mv)
    const avg       = latestQty / 12
    const st        = ojcStock[cat] ?? 0
    const cpk       = st > 0 && pk > 0 ? parseFloat((st / pk).toFixed(1)) : null
    const cav       = st > 0 && avg > 0 ? parseFloat((st / avg).toFixed(1)) : null
    const row = ws1.addRow([
      cat, Object.keys(data.products).length,
      ...years.map(yr => data.annuals[yr] || null),
      yoyRate(latestQty, prevQty), cagrRate(firstQty, latestQty, years.length),
      st || null, cpk, cav,
    ])
    row.height = 18
    applyTrendRow(row, ri1 % 2 === 0 ? C.white : C.altBlue, C.midBlue)
    ri1++
  }

  // ── [C] 기타 품목 판매 현황 ───────────────────────────────────────────────
  addSecHdr1('  ▶ 기타 품목 판매 현황')
  styleHdr(ws1.addRow([
    '카테고리', '품목수', ...years.map(y => `20${y}년\n판매(EA)`),
    prevYr ? `20${latestYr}년/\n20${prevYr}년 전년비` : '전년비',
    '연평균\n성장률(CAGR)', '현재고\n(EA)', '재고여유\n(최고월)', '재고여유\n(평균)',
  ]), C.purple, 32)

  let ri2 = 0
  for (const [cat, data] of Object.entries(etcByCategory)) {
    const firstQty  = data.annuals[years[0]] ?? 0
    const latestQty = data.annuals[latestYr]  ?? 0
    const prevQty   = data.annuals[prevYr]    ?? 0
    const mv        = MONTHS.map(m => data.monthlyLatest[m] ?? 0)
    const pk        = Math.max(0, ...mv)
    const avg       = latestQty / 12
    const st        = ojcStock[cat] ?? 0
    const cpk       = st > 0 && pk > 0 ? parseFloat((st / pk).toFixed(1)) : null
    const cav       = st > 0 && avg > 0 ? parseFloat((st / avg).toFixed(1)) : null
    const row = ws1.addRow([
      cat, Object.keys(data.products).length,
      ...years.map(yr => data.annuals[yr] || null),
      yoyRate(latestQty, prevQty), cagrRate(firstQty, latestQty, years.length),
      st || null, cpk, cav,
    ])
    row.height = 18
    applyTrendRow(row, ri2 % 2 === 0 ? C.white : C.altPurple, C.purple)
    ri2++
  }

  // ── [D] 거래처 TOP5 (내자 / 외자) ───────────────────────────────────────
  addSecHdr1('  ▶ 거래처 TOP5')

  const writeCustTop5 = (
    hdrLabel: string | null, hdrColor: string,
    subLabel: string, subBg: string, subFontColor: string,
    filteredRows: typeof rawRows,
    sortBy: 'qty' | 'price',
  ) => {
    if (hdrLabel) {
      const hd = ws1.addRow([
        '순위', '거래처명',
        sortBy === 'price' ? '누적 금액(원)' : '누적 구매(EA)',
        sortBy === 'price' ? `20${latestYr}년(원)` : `20${latestYr}년(EA)`,
        '전년비', ...Array(Math.max(0, S1 - 5)).fill(null),
      ])
      styleHdr(hd, hdrColor, 22)
      if (S1 > 5) ws1.mergeCells(hd.number, 5, hd.number, S1)
    }
    const sub = ws1.addRow([subLabel, ...Array(S1 - 1).fill(null)])
    ws1.mergeCells(sub.number, 1, sub.number, S1)
    const sc = ws1.getCell(sub.number, 1)
    sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: subBg } }
    sc.font = { bold: true, size: 9, color: { argb: subFontColor } }
    sc.alignment = { horizontal: 'left', vertical: 'middle' }
    sub.height = 16

    Object.entries(buildTop3(filteredRows, 'cumul', sortBy)).slice(0, 5).forEach(([cust, prods], ci) => {
      const isPrice  = sortBy === 'price'
      const cumul    = prods.reduce((s, x) => s + (isPrice ? x.price : x.qty), 0)
      const custLat  = filteredRows.filter(r => r.customer === cust && r.year === latestYr).reduce((s, r) => s + (isPrice ? r.price : r.qty), 0)
      const custPrev = filteredRows.filter(r => r.customer === cust && r.year === prevYr).reduce((s, r) => s + (isPrice ? r.price : r.qty), 0)
      const row = ws1.addRow([ci + 1, cust, cumul || null, custLat || null, yoyRate(custLat, custPrev), ...Array(Math.max(0, S1 - 5)).fill(null)])
      if (S1 > 5) ws1.mergeCells(row.number, 5, row.number, S1)
      row.height = 18
      const bgRow = ci % 2 === 0 ? C.white : (subBg === C.altPurple ? C.altPurple : C.altBlue)
      row.eachCell({ includeEmpty: true }, (c, col) => {
        c.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } }
        c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
        if (col === 1) {
          c.alignment = { horizontal: 'center', vertical: 'middle' }
          c.font = { bold: true, color: { argb: ci === 0 ? 'FFCC0000' : ci === 1 ? 'FF833C00' : C.darkBlue } }
        } else if (col === 2) {
          c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { bold: true }
        } else if (col === 5) {
          c.alignment = { horizontal: 'right', vertical: 'middle' }
          if (typeof c.value === 'number') {
            c.numFmt = '+0.0%;-0.0%;"-"'
            c.font = { bold: true, color: { argb: c.value > 0 ? C.green : c.value < 0 ? C.red : C.gray3 } }
          }
        } else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
      })
    })
  }

  writeCustTop5('  📦 구매량 기준', C.midBlue, '  🏠 내자 거래처', C.lightBlue, C.darkBlue, domesticRowsForExcel, 'qty')
  writeCustTop5(null, C.midBlue, '  ✈️  외자 거래처', C.altPurple, C.purple, foreignRowsForExcel, 'qty')
  writeCustTop5('  💰 판매금액 기준', C.purple, '  🏠 내자 거래처', C.lightBlue, C.darkBlue, domesticRowsForExcel, 'price')
  writeCustTop5(null, C.purple, '  ✈️  외자 거래처', C.altPurple, C.purple, foreignRowsForExcel, 'price')

  ws1.columns = [
    { width: 26 }, { width: 10 },
    ...years.map(() => ({ width: 14 })),
    { width: 11 }, { width: 11 }, { width: 12 }, { width: 12 }, { width: 12 },
  ]

  await prog(2, '② 연도별_피크분석')
  // ── Sheet 2: 연도별_피크분석 ──────────────────────────────────────────
  const ws2 = wb.addWorksheet('② 연도별_피크분석')
  ws2.views = [{ state: 'frozen', ySplit: 5 }]
  const S2 = 2 + years.length * 5 + 3
  addTitle(ws2, '연도별 최고월 분석 (OJC + 기타)',
    `각 연도별 최대 판매 월·수량·집중도·전년비  |  커버리지는 ${latestYr}년 기준`, S2)

  // ── 지표 설명 행 ──
  const noteStyle = (row: ExcelJS.Row, label: string, desc: string) => {
    row.height = 15
    const c1 = row.getCell(1); c1.value = label
    c1.font = { bold: true, size: 9, color: { argb: C.darkBlue } }
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF1FA' } }
    c1.alignment = { horizontal: 'left', vertical: 'middle' }
    ws2.mergeCells(row.number, 1, row.number, S2)
    const mc = ws2.getCell(row.number, 1); mc.value = `${label}  ${desc}`
    mc.font = { size: 9, color: { argb: C.gray3 } }
    mc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF1FA' } }
    mc.alignment = { horizontal: 'left', vertical: 'middle' }
  }
  noteStyle(ws2.addRow([]), '📌 집중도',
    '= 피크(EA) ÷ 월평균  |  1.0배 = 완전 균등 분포 / 2.5배 이상(주황) = 특정 월 집중 → 재고 추가 확보 필요')
  noteStyle(ws2.addRow([]), '📌 전년비',
    '= (올해 피크 − 전년 피크) ÷ 전년 피크  |  상승 초록 / 하락 빨강 / 첫 연도는 비교 대상 없으므로 — 표시')

  styleHdr(ws2.addRow(['카테고리', '품목수',
    ...years.flatMap(y => [`${y}년\n판매(EA)`, `${y}년\n최고월`, `${y}년\n최고월판매량`, `${y}년\n집중도`, `${y}년\n전년비`]),
    '현재고\n(EA)', '최고월\n재고여유', '평균\n재고여유']), C.darkBlue, 36)

  function addPeakRow(ws: ExcelJS.Worksheet, cat: string, data: CategoryData, isEtc: boolean, rn: number) {
    const la = data.annuals[latestYr] ?? 0; const mv = MONTHS.map(m => data.monthlyLatest[m] ?? 0)
    const pk = Math.max(0, ...mv); const avg = la / 12; const st = ojcStock[cat] ?? 0
    const cpk = st > 0 && pk > 0 ? parseFloat((st / pk).toFixed(1)) : null
    const cav = st > 0 && avg > 0 ? parseFloat((st / avg).toFixed(1)) : null
    const bg = rn % 2 === 0 ? C.white : (isEtc ? C.altPurple : C.altBlue)
    const yearCols = years.flatMap((yr, idx) => {
      const p    = peakOfYear(data, yr)
      const ann  = data.annuals[yr] ?? 0
      const mAvg = ann / 12
      const conc = mAvg > 0 && p.val > 0 ? parseFloat((p.val / mAvg).toFixed(2)) : null
      const prev = idx > 0 ? peakOfYear(data, years[idx - 1]).val : 0
      const yoy  = prev > 0 && p.val > 0 ? parseFloat(((p.val - prev) / prev).toFixed(3)) : null
      return [ann || null, p.month || null, p.val || null, conc, yoy]
    })
    const row = ws.addRow([cat, Object.keys(data.products).length, ...yearCols, st || null, cpk, cav])
    row.height = 18
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ci === 1) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { bold: true, color: { argb: isEtc ? C.purple : C.midBlue } } }
      else if (ci === 2) { c.alignment = { horizontal: 'center', vertical: 'middle' } }
      else if (ci <= 2 + years.length * 5) {
        const rel = (ci - 3) % 5
        if (rel === 1) {
          c.alignment = { horizontal: 'center', vertical: 'middle' }
          if (c.value) c.font = { bold: true, color: { argb: C.orange } }
        } else if (rel === 3) {
          c.alignment = { horizontal: 'right', vertical: 'middle' }
          if (typeof c.value === 'number') {
            c.numFmt = '0.0"배"'
            if (c.value >= 2.5) c.font = { bold: true, color: { argb: C.orange } }
          }
        } else if (rel === 4) {
          c.alignment = { horizontal: 'right', vertical: 'middle' }
          if (typeof c.value === 'number') {
            c.numFmt = '+0.0%;-0.0%;"-"'
            c.font = { color: { argb: c.value > 0 ? C.green : C.red } }
          }
        } else {
          c.alignment = { horizontal: 'right', vertical: 'middle' }
          if (typeof c.value === 'number') c.numFmt = '#,##0'
        }
      } else {
        c.alignment = { horizontal: 'right', vertical: 'middle' }
        if (typeof c.value === 'number') c.numFmt = '#,##0'
      }
    })
    if (cpk !== null) row.getCell(S2 - 1).font = { bold: true, color: { argb: coverFg(cpk) } }
    if (cav !== null) row.getCell(S2).font = { bold: true, color: { argb: coverFg(cav) } }
  }

  let ri = 0
  for (const [c, d] of Object.entries(ojcByCategory)) addPeakRow(ws2, c, d, false, ri++)
  addSep(ws2, '  기타 품목', S2); ri = 0
  for (const [c, d] of Object.entries(etcByCategory)) addPeakRow(ws2, c, d, true, ri++)
  ws2.columns = [{ width: 28 }, { width: 7 }, ...years.flatMap(() => [{ width: 13 }, { width: 9 }, { width: 12 }, { width: 9 }, { width: 9 }]), { width: 11 }, { width: 11 }, { width: 11 }]

  await prog(3, '③ OJC_월간상세')
  // ── Sheet 3: OJC 품목별_월간상세 (KT-SP/DP/다심 · LG-S/D/M 세분화) ────
  const ojcDetailedMap = buildMonthlyMap(rawRows, classifyOjcDetailed)
  const ojcDetailEntries = Object.entries(ojcDetailedMap)
    .sort(([a], [b]) => {
      const ai = OJC_DETAIL_ORDER.indexOf(a); const bi = OJC_DETAIL_ORDER.indexOf(b)
      return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi)
    })
  const S3 = 18  // 카테고리+코드+품목명+연도+12월+합계+공급가액
  const ws3 = wb.addWorksheet('③ OJC_월간상세')
  ws3.properties.outlineProperties = { summaryBelow: false, summaryRight: false }
  ws3.views = [{ state: 'frozen', xSplit: 3, ySplit: 3 }]
  addTitle(ws3, 'OJC 품목별 월간 판매 상세', 'KT-SP/DP/다심 · LG-S/D/M · 피그테일-성단용 세분화  |  주황 = 최고월  |  카테고리 클릭으로 접기/펼치기', S3)
  styleHdr(ws3.addRow(['카테고리', '품목코드', '품목명', '연도', ...MONTHS.map(m => `${parseInt(m)}월`), '합계', '공급가액\n(연간)']), C.midBlue, 28)

  for (const [cat, catData] of ojcDetailEntries) {
    const catBg = COLOR_MAP_OJC_DETAILED[cat] ?? 'FFDEEAF1'
    for (const yr of years) {
      const byYr  = catData.monthlyByYear[yr] ?? {}
      const mv    = MONTHS.map(m => byYr[m] ?? 0)
      const pk    = Math.max(0, ...mv)
      const total = mv.reduce((s, v) => s + v, 0)
      const catRow = ws3.addRow([cat, '', '(합계)', `20${yr}년`, ...MONTHS.map(m => byYr[m] || null), total || null, catData.priceAnnuals[yr] || null])
      catRow.outlineLevel = 0; catRow.height = 20
      catRow.eachCell({ includeEmpty: true }, (c, ci) => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: catBg } }
        c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
        c.font = { bold: true }
        if (ci <= 4) c.alignment = { horizontal: 'left', vertical: 'middle' }
        else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
      })
      MONTHS.forEach((m, mi) => { if ((byYr[m] ?? 0) === pk && pk > 0) { const pc = catRow.getCell(5 + mi); pc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } }; pc.font = { bold: true, color: { argb: C.orange } } } })
    }
    for (const [name, prod] of Object.entries(catData.products).sort((a, b) => (b[1].annuals[latestYr] ?? 0) - (a[1].annuals[latestYr] ?? 0))) {
      for (const yr of years) {
        const byYr  = prod.monthlyByYear[yr] ?? {}
        const mv    = MONTHS.map(m => byYr[m] ?? 0)
        const pk    = Math.max(0, ...mv)
        const total = mv.reduce((s, v) => s + v, 0)
        const row   = ws3.addRow([cat, prod.code || null, name, `20${yr}년`, ...MONTHS.map(m => byYr[m] || null), total || null, prod.priceAnnuals[yr] || null])
        row.outlineLevel = 1; row.height = 16
        row.eachCell({ includeEmpty: true }, (c, ci) => {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.white } }
          c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
          if (ci === 1) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { color: { argb: C.midBlue } } }
          else if (ci <= 4) c.alignment = { horizontal: 'left', vertical: 'middle' }
          else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
        })
        MONTHS.forEach((m, mi) => { if ((byYr[m] ?? 0) === pk && pk > 0) { const pc = row.getCell(5 + mi); pc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } }; pc.font = { bold: true, color: { argb: C.orange } } } })
      }
    }
  }
  ws3.columns = [{ width: 22 }, { width: 16 }, { width: 40 }, { width: 9 }, ...MONTHS.map(() => ({ width: 8 })), { width: 9 }, { width: 16 }]

  await prog(4, '④ OJC외_월간상세')
  // ── Sheet 4: OJC 외 품목별_월간상세 ─────────────────────────────────────
  const etcMonthlyMap = buildMonthlyMap(rawRows, classifyEtc)
  const ws4 = wb.addWorksheet('④ OJC외_월간상세')
  ws4.properties.outlineProperties = { summaryBelow: false, summaryRight: false }
  ws4.views = [{ state: 'frozen', xSplit: 3, ySplit: 3 }]
  addTitle(ws4, 'OJC 외 품목별 월간 판매 상세', '카테고리별 접기/펼치기  |  주황 = 최고월', S3)
  styleHdr(ws4.addRow(['카테고리', '품목코드', '품목명', '연도', ...MONTHS.map(m => `${parseInt(m)}월`), '합계', '공급가액\n(연간)']), C.midBlue, 28)

  for (const [cat, catData] of Object.entries(etcMonthlyMap)) {
    for (const yr of years) {
      const byYr  = catData.monthlyByYear[yr] ?? {}
      const mv    = MONTHS.map(m => byYr[m] ?? 0)
      const pk    = Math.max(0, ...mv)
      const total = mv.reduce((s, v) => s + v, 0)
      const catRow = ws4.addRow([cat, '', '(합계)', `20${yr}년`, ...MONTHS.map(m => byYr[m] || null), total || null, catData.priceAnnuals[yr] || null])
      catRow.outlineLevel = 0; catRow.height = 20
      catRow.eachCell({ includeEmpty: true }, (c, ci) => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lightPurple } }
        c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
        c.font = { bold: true, color: { argb: C.purple } }
        if (ci <= 4) c.alignment = { horizontal: 'left', vertical: 'middle' }
        else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
      })
      MONTHS.forEach((m, mi) => { if ((byYr[m] ?? 0) === pk && pk > 0) { const pc = catRow.getCell(5 + mi); pc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } }; pc.font = { bold: true, color: { argb: C.orange } } } })
    }
    for (const [name, prod] of Object.entries(catData.products).sort((a, b) => (b[1].annuals[latestYr] ?? 0) - (a[1].annuals[latestYr] ?? 0))) {
      for (const yr of years) {
        const byYr  = prod.monthlyByYear[yr] ?? {}
        const mv    = MONTHS.map(m => byYr[m] ?? 0)
        const pk    = Math.max(0, ...mv)
        const total = mv.reduce((s, v) => s + v, 0)
        const row   = ws4.addRow([cat, prod.code || null, name, `20${yr}년`, ...MONTHS.map(m => byYr[m] || null), total || null, prod.priceAnnuals[yr] || null])
        row.outlineLevel = 1; row.height = 16
        row.eachCell({ includeEmpty: true }, (c, ci) => {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.white } }
          c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
          if (ci === 1) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { color: { argb: C.purple } } }
          else if (ci <= 4) c.alignment = { horizontal: 'left', vertical: 'middle' }
          else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
        })
        MONTHS.forEach((m, mi) => { if ((byYr[m] ?? 0) === pk && pk > 0) { const pc = row.getCell(5 + mi); pc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } }; pc.font = { bold: true, color: { argb: C.orange } } } })
      }
    }
  }
  ws4.columns = [{ width: 22 }, { width: 16 }, { width: 40 }, { width: 9 }, ...MONTHS.map(() => ({ width: 8 })), { width: 9 }, { width: 16 }]

  await prog(5, '⑤⑥ 거래처별탑3 내자·외자')
  // ── Sheet 5/6: 거래처별탑3 내자·외자 ─────────────────────────────────
  const s5TopColors = [C.midBlue, 'FF1A5A96', 'FF155480'] as const
  const s5ColWidths = [
    { width: 6 }, { width: 26 }, { width: 10 }, { width: 13 },
    { width: 14 }, { width: 34 }, { width: 11 }, { width: 9 }, { width: 14 }, { width: 9 },
    { width: 14 }, { width: 34 }, { width: 11 }, { width: 9 }, { width: 14 }, { width: 9 },
    { width: 14 }, { width: 34 }, { width: 11 }, { width: 9 }, { width: 14 }, { width: 9 },
  ]

  const writeCustTop3Body = (
    ws: ExcelJS.Worksheet,
    custTop3: Record<string, Array<{ name: string; code: string; qty: number; price: number }>>,
    filteredRows: typeof rawRows,
  ) => {
    const sh1 = ws.addRow(['순위', '거래처명', '연도', '총구매(EA)',
      'TOP 1', null, null, null, null, null, 'TOP 2', null, null, null, null, null, 'TOP 3', null, null, null, null, null])
    const sh2 = ws.addRow([null, null, null, null,
      '품목코드', '품목명', '수량(EA)', '수량점유율', '공급가액', '가액점유율',
      '품목코드', '품목명', '수량(EA)', '수량점유율', '공급가액', '가액점유율',
      '품목코드', '품목명', '수량(EA)', '수량점유율', '공급가액', '가액점유율'])
    styleHdr(sh1, C.darkBlue, 26); styleHdr(sh2, C.darkBlue, 20)
    ;[1, 2, 3, 4].forEach(ci => ws.mergeCells(sh1.number, ci, sh2.number, ci))
    ws.mergeCells(sh1.number, 5,  sh1.number, 10)
    ws.mergeCells(sh1.number, 11, sh1.number, 16)
    ws.mergeCells(sh1.number, 17, sh1.number, 22)
    ;[5, 11, 17].forEach((ci, gi) => {
      ws.getCell(sh1.number, ci).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s5TopColors[gi] } }
    })
    const applyStyle = (row: ExcelJS.Row, bg: string, isCumul: boolean) => {
      row.eachCell({ includeEmpty: true }, (c, ci) => {
        c.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
        c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
        if (ci === 1) { c.alignment = { horizontal: 'center', vertical: 'middle' }; c.font = { color: { argb: C.gray3 } } }
        else if (ci === 2) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { bold: true } }
        else if (ci === 3) { c.alignment = { horizontal: 'center', vertical: 'middle' }; if (isCumul) c.font = { bold: true, color: { argb: C.darkBlue } } }
        else if (ci === 4) { c.alignment = { horizontal: 'right', vertical: 'middle' }; c.font = { bold: true }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
        else {
          const pos = (ci - 5) % 6
          if (pos === 0) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { size: 9, color: { argb: C.gray3 } } }
          else if (pos === 1) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { size: 9 } }
          else if (pos === 2) { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
          else if (pos === 3 || pos === 5) {
            c.alignment = { horizontal: 'right', vertical: 'middle' }
            if (typeof c.value === 'number') {
              c.numFmt = '0%'
              c.font = { bold: true, color: { argb: c.value >= 0.6 ? 'FFE26B0A' : c.value >= 0.3 ? 'FF2E75B6' : 'FF808080' } }
            }
          } else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
        }
      })
    }
    Object.entries(custTop3).forEach(([cust, cumulTop], custIdx) => {
      const cumulTotal      = cumulTop.reduce((s, x) => s + x.qty, 0)
      const cumulPriceTotal = cumulTop.reduce((s, x) => s + x.price, 0)
      const startRow        = ws.rowCount + 1
      const rowBg           = custIdx % 2 === 0 ? C.white : C.altBlue
      years.forEach(yr => {
        const yrMap      = buildTop3(filteredRows, yr)
        const top3       = yrMap[cust] ?? []
        const total      = top3.reduce((s, x) => s + x.qty, 0)
        const priceTotal = top3.reduce((s, x) => s + x.price, 0)
        const qp  = (i: number) => top3[i] && total      > 0 ? top3[i].qty   / total      : null
        const pp  = (i: number) => top3[i] && priceTotal > 0 ? top3[i].price / priceTotal : null
        const row = ws.addRow([
          null, null, `20${yr}년`, total || null,
          top3[0]?.code || null, top3[0]?.name ?? null, top3[0]?.qty ?? null, qp(0), top3[0]?.price || null, pp(0),
          top3[1]?.code || null, top3[1]?.name ?? null, top3[1]?.qty ?? null, qp(1), top3[1]?.price || null, pp(1),
          top3[2]?.code || null, top3[2]?.name ?? null, top3[2]?.qty ?? null, qp(2), top3[2]?.price || null, pp(2),
        ])
        row.height = 17; applyStyle(row, rowBg, false)
      })
      const cqp = (i: number) => cumulTop[i] && cumulTotal      > 0 ? cumulTop[i].qty   / cumulTotal      : null
      const cpp = (i: number) => cumulTop[i] && cumulPriceTotal > 0 ? cumulTop[i].price / cumulPriceTotal : null
      const cumulRow = ws.addRow([
        null, null, '전체 누적', cumulTotal || null,
        cumulTop[0]?.code || null, cumulTop[0]?.name ?? null, cumulTop[0]?.qty ?? null, cqp(0), cumulTop[0]?.price || null, cpp(0),
        cumulTop[1]?.code || null, cumulTop[1]?.name ?? null, cumulTop[1]?.qty ?? null, cqp(1), cumulTop[1]?.price || null, cpp(1),
        cumulTop[2]?.code || null, cumulTop[2]?.name ?? null, cumulTop[2]?.qty ?? null, cqp(2), cumulTop[2]?.price || null, cpp(2),
      ])
      cumulRow.height = 17; applyStyle(cumulRow, C.lightBlue, true)
      const endRow = ws.rowCount
      if (endRow > startRow) { ws.mergeCells(startRow, 1, endRow, 1); ws.mergeCells(startRow, 2, endRow, 2) }
      const rkCell = ws.getCell(startRow, 1)
      rkCell.value = custIdx + 1; rkCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } }
      rkCell.alignment = { horizontal: 'center', vertical: 'middle' }; rkCell.font = { color: { argb: C.gray3 } }
      const cnCell = ws.getCell(startRow, 2)
      cnCell.value = cust; cnCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } }
      cnCell.alignment = { horizontal: 'left', vertical: 'middle' }
      cnCell.font = { bold: true, color: { argb: custIdx === 0 ? 'FFCC0000' : custIdx === 1 ? 'FF833C00' : C.darkBlue } }
    })
    ws.columns = s5ColWidths
  }

  const ws5 = wb.addWorksheet('⑤ 거래처별탑3_내자')
  ws5.views = [{ state: 'frozen', xSplit: 2, ySplit: 4 }]
  addTitle(ws5, '거래처별 TOP3 구매 품목 (내자 🏠)',
    '내자 거래처 전용 (환율=0)  |  총구매량 내림차순  |  수량점유율·가액점유율: 해당 품목의 거래처 TOP3 내 비중', 22)
  writeCustTop3Body(ws5, buildTop3(domesticRowsForExcel, 'cumul'), domesticRowsForExcel)

  const ws6 = wb.addWorksheet('⑥ 거래처별탑3_외자')
  ws6.views = [{ state: 'frozen', xSplit: 2, ySplit: 4 }]
  addTitle(ws6, '거래처별 TOP3 구매 품목 (외자 ✈️)',
    '외자 거래처 전용 (환율>0)  |  총구매량 내림차순  |  수량점유율·가액점유율: 해당 품목의 거래처 TOP3 내 비중', 22)
  writeCustTop3Body(ws6, buildTop3(foreignRowsForExcel, 'cumul'), foreignRowsForExcel)

  await prog(7, '⑦ 전체판매량')
  // ── Sheet 7: 전체판매량 ───────────────────────────────────────────────
  const ws7 = wb.addWorksheet('⑦ 전체판매량')
  ws7.views = [{ state: 'frozen', ySplit: 3 }]
  // col 구조: 코드(1)+품목명(1)+분류(1) + yr0:판매+공급(2) + yr1+:(판매+전년비+공급)*N + 합계(1)
  const S6 = 3 + 2 + (years.length - 1) * 3 + 1
  const s6ColTypes = ['code', 'name', 'cat',
    ...years.flatMap((_, idx) => idx === 0 ? ['annual', 'price'] : ['annual', 'growth', 'price']),
    'total',
  ]
  addTitle(ws7, '전체 품목 판매량', `${latestYr}년 판매량 기준 내림차순  |  전년비: 전년 대비 판매수량 증감률`, S6)
  styleHdr(ws7.addRow([
    '품목코드', '품목명', '분류',
    ...years.flatMap((y, idx) => idx === 0
      ? [`${y}년\n판매(EA)`, `${y}년\n공급가액`]
      : [`${y}년\n판매(EA)`, '전년비', `${y}년\n공급가액`]
    ),
    '합계\n(EA)',
  ]), C.darkBlue, 28)
  fullProducts.forEach((p, i) => {
    const total = years.reduce((s, yr) => s + (p.annuals[yr] ?? 0), 0)
    const bg = i % 2 === 0 ? C.white : C.gray1
    const row = ws7.addRow([
      p.code || null, p.name, p.ojcCat ?? p.etcCat ?? '기타',
      ...years.flatMap((yr, idx) => {
        const curr = p.annuals[yr] ?? 0
        const prev = idx > 0 ? (p.annuals[years[idx - 1]] ?? 0) : 0
        const g = idx > 0 && prev > 0 && curr > 0 ? (curr - prev) / prev : null
        return idx === 0
          ? [curr || null, p.priceAnnuals[yr] || null]
          : [curr || null, g, p.priceAnnuals[yr] || null]
      }),
      total || null,
    ])
    row.height = 16
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      const ct = s6ColTypes[ci - 1]
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ct === 'code') { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { size: 9, color: { argb: C.gray3 } } }
      else if (ct === 'name') { c.alignment = { horizontal: 'left', vertical: 'middle' } }
      else if (ct === 'cat') { c.alignment = { horizontal: 'left', vertical: 'middle' }; if (p.ojcCat) c.font = { color: { argb: C.midBlue } } }
      else if (ct === 'growth') {
        c.alignment = { horizontal: 'right', vertical: 'middle' }
        if (typeof c.value === 'number') {
          c.numFmt = '+0.0%;-0.0%;"-"'
          c.font = { color: { argb: c.value > 0 ? 'FF375623' : 'FF9C0006' } }
        }
      }
      else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
    })
  })
  if (fullProducts.length > 0) {
    const tots = years.map(yr => fullProducts.reduce((s, p) => s + (p.annuals[yr] ?? 0), 0))
    const srVals = [
      '합계', `${fullProducts.length}종`, '',
      ...years.flatMap((_, idx) => idx === 0 ? [tots[0], null] : [tots[idx], null, null]),
      tots.reduce((a, b) => a + b, 0),
    ]
    const sr = ws7.addRow(srVals); sr.height = 22
    sr.eachCell({ includeEmpty: true }, (c, ci) => {
      const ct = s6ColTypes[ci - 1]
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.darkBlue } }; c.font = { bold: true, color: { argb: C.white } }
      c.border = { top: { style: 'medium', color: { argb: C.midBlue } } }
      c.alignment = { horizontal: ct === 'code' || ct === 'name' || ct === 'cat' ? 'left' : 'right', vertical: 'middle' }
      if (typeof c.value === 'number') c.numFmt = '#,##0'
    })
  }
  ws7.columns = [
    { width: 16 }, { width: 40 }, { width: 18 },
    ...years.flatMap((_, idx) => idx === 0
      ? [{ width: 13 }, { width: 16 }]
      : [{ width: 13 }, { width: 10 }, { width: 16 }]
    ),
    { width: 13 },
  ]

  await prog(8, '⑧ 거래처별OJC탑3')
  // ── Sheet 8: 거래처별 OJC 탑3 (연도 행 구조, 거래처별 머지) ────────────
  const ws8 = wb.addWorksheet('⑧ 거래처별OJC탑3')
  ws8.views = [{ state: 'frozen', xSplit: 2, ySplit: 4 }]
  addTitle(ws8, '거래처별 OJC TOP3 구매 품목',
    '거래처별 연도 행 구조  |  OJC 총구매량 내림차순  |  수량점유율·가액점유율: 해당 품목의 거래처 TOP3 내 비중', 25)
  const ojcRowsForExcel = rawRows.filter(r => classifyOjc(r.name) !== null)
  const ojcCumulTop3   = buildTop3(ojcRowsForExcel, 'cumul')

  const s8h1 = ws8.addRow(['순위', '거래처명', '연도', 'OJC 총구매(EA)',
    'TOP 1', null, null, null, null, null, null,
    'TOP 2', null, null, null, null, null, null,
    'TOP 3', null, null, null, null, null, null])
  const s8h2 = ws8.addRow([null, null, null, null,
    '품목코드', '품목명', 'OJC분류', '수량(EA)', '수량점유율', '공급가액', '가액점유율',
    '품목코드', '품목명', 'OJC분류', '수량(EA)', '수량점유율', '공급가액', '가액점유율',
    '품목코드', '품목명', 'OJC분류', '수량(EA)', '수량점유율', '공급가액', '가액점유율'])
  styleHdr(s8h1, C.darkBlue, 26); styleHdr(s8h2, C.darkBlue, 20)
  ;[1, 2, 3, 4].forEach(ci => ws8.mergeCells(s8h1.number, ci, s8h2.number, ci))
  ws8.mergeCells(s8h1.number, 5,  s8h1.number, 11)
  ws8.mergeCells(s8h1.number, 12, s8h1.number, 18)
  ws8.mergeCells(s8h1.number, 19, s8h1.number, 25)
  ;[5, 12, 19].forEach((ci, gi) => {
    ws8.getCell(s8h1.number, ci).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s5TopColors[gi] } }
  })

  const applyS7Style = (row: ExcelJS.Row, bg: string, isCumul: boolean) => {
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ci === 1) {
        c.alignment = { horizontal: 'center', vertical: 'middle' }; c.font = { color: { argb: C.gray3 } }
      } else if (ci === 2) {
        c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { bold: true }
      } else if (ci === 3) {
        c.alignment = { horizontal: 'center', vertical: 'middle' }
        if (isCumul) c.font = { bold: true, color: { argb: C.darkBlue } }
      } else if (ci === 4) {
        c.alignment = { horizontal: 'right', vertical: 'middle' }; c.font = { bold: true, color: { argb: C.midBlue } }
        if (typeof c.value === 'number') c.numFmt = '#,##0'
      } else {
        const pos = (ci - 5) % 7
        if (pos === 0) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { size: 9, color: { argb: C.gray3 } } }
        else if (pos === 1) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { size: 9 } }
        else if (pos === 2) { c.alignment = { horizontal: 'center', vertical: 'middle' }; c.font = { size: 9, color: { argb: C.midBlue } } }
        else if (pos === 3) { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
        else if (pos === 4 || pos === 6) {
          c.alignment = { horizontal: 'right', vertical: 'middle' }
          if (typeof c.value === 'number') {
            c.numFmt = '0%'
            c.font = { bold: true, color: { argb: c.value >= 0.6 ? 'FFE26B0A' : c.value >= 0.3 ? 'FF2E75B6' : 'FF808080' } }
          }
        } else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
      }
    })
  }

  Object.entries(ojcCumulTop3).forEach(([cust, cumulTop], custIdx) => {
    const cumulTotal      = cumulTop.reduce((s, x) => s + x.qty, 0)
    const cumulPriceTotal = cumulTop.reduce((s, x) => s + x.price, 0)
    const startRow        = ws8.rowCount + 1
    const rowBg           = custIdx % 2 === 0 ? C.white : C.altBlue
    years.forEach(yr => {
      const yrMap      = buildTop3(ojcRowsForExcel, yr)
      const top3       = yrMap[cust] ?? []
      const total      = top3.reduce((s, x) => s + x.qty, 0)
      const priceTotal = top3.reduce((s, x) => s + x.price, 0)
      const qp  = (i: number) => top3[i] && total      > 0 ? top3[i].qty   / total      : null
      const pp  = (i: number) => top3[i] && priceTotal > 0 ? top3[i].price / priceTotal : null
      const ojcC = (i: number) => top3[i] ? classifyOjc(top3[i].name) ?? null : null
      const row  = ws8.addRow([
        null, null, `20${yr}년`, total || null,
        top3[0]?.code || null, top3[0]?.name ?? null, ojcC(0), top3[0]?.qty ?? null, qp(0), top3[0]?.price || null, pp(0),
        top3[1]?.code || null, top3[1]?.name ?? null, ojcC(1), top3[1]?.qty ?? null, qp(1), top3[1]?.price || null, pp(1),
        top3[2]?.code || null, top3[2]?.name ?? null, ojcC(2), top3[2]?.qty ?? null, qp(2), top3[2]?.price || null, pp(2),
      ])
      row.height = 17; applyS7Style(row, rowBg, false)
    })
    const cqp  = (i: number) => cumulTop[i] && cumulTotal      > 0 ? cumulTop[i].qty   / cumulTotal      : null
    const cpp  = (i: number) => cumulTop[i] && cumulPriceTotal > 0 ? cumulTop[i].price / cumulPriceTotal : null
    const cojcC = (i: number) => cumulTop[i] ? classifyOjc(cumulTop[i].name) ?? null : null
    const cumulRow = ws8.addRow([
      null, null, '전체 누적', cumulTotal || null,
      cumulTop[0]?.code || null, cumulTop[0]?.name ?? null, cojcC(0), cumulTop[0]?.qty ?? null, cqp(0), cumulTop[0]?.price || null, cpp(0),
      cumulTop[1]?.code || null, cumulTop[1]?.name ?? null, cojcC(1), cumulTop[1]?.qty ?? null, cqp(1), cumulTop[1]?.price || null, cpp(1),
      cumulTop[2]?.code || null, cumulTop[2]?.name ?? null, cojcC(2), cumulTop[2]?.qty ?? null, cqp(2), cumulTop[2]?.price || null, cpp(2),
    ])
    cumulRow.height = 17; applyS7Style(cumulRow, C.lightBlue, true)
    const endRow = ws8.rowCount
    if (endRow > startRow) { ws8.mergeCells(startRow, 1, endRow, 1); ws8.mergeCells(startRow, 2, endRow, 2) }
    const rkCell8 = ws8.getCell(startRow, 1)
    rkCell8.value = custIdx + 1; rkCell8.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } }
    rkCell8.alignment = { horizontal: 'center', vertical: 'middle' }; rkCell8.font = { color: { argb: C.gray3 } }
    const cnCell8 = ws8.getCell(startRow, 2)
    cnCell8.value = cust; cnCell8.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } }
    cnCell8.alignment = { horizontal: 'left', vertical: 'middle' }
    cnCell8.font = { bold: true, color: { argb: custIdx === 0 ? 'FFCC0000' : custIdx === 1 ? 'FF833C00' : C.darkBlue } }
  })

  ws8.columns = [
    { width: 6 }, { width: 26 }, { width: 10 }, { width: 14 },
    { width: 14 }, { width: 32 }, { width: 10 }, { width: 11 }, { width: 9 }, { width: 14 },
    { width: 14 }, { width: 32 }, { width: 10 }, { width: 11 }, { width: 9 }, { width: 14 },
    { width: 14 }, { width: 32 }, { width: 10 }, { width: 11 }, { width: 9 }, { width: 14 },
  ]

  await prog(9, '파일 저장 중...')
  const buffer = await wb.xlsx.writeBuffer()
  if (fileHandle) {
    await writeToFileHandle(fileHandle, buffer as ArrayBuffer)
  } else {
    downloadXlsx(buffer as ArrayBuffer, `판매현황분석_${today()}.xlsx`)
  }
}

function OjcSalesView({
  ojcByCategory, etcByCategory, years, latestYr, ojcStock, onStockChange,
}: {
  ojcByCategory:  Record<string, CategoryData>
  etcByCategory:  Record<string, CategoryData>
  years:          string[]
  latestYr:       string
  ojcStock:       Record<string, number>
  onStockChange:  (cat: string, val: string) => void
}) {
  const [expanded, setExpanded]         = useState<string | null>(null)
  const [expandedYear, setExpandedYear] = useState<string>(latestYr)

  const handleExpand = (key: string) => {
    setExpanded(prev => {
      if (prev === key) return null
      setExpandedYear(latestYr)
      return key
    })
  }

  const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'
  const coverFgClass = (v: number) => v >= 2 ? 'text-green-600' : v >= 1 ? 'text-yellow-600' : 'text-red-500'

  // 재고는 있지만 판매 데이터가 전혀 없는 카테고리 (OJC/ETC 모두 아닌 것)
  const ojcCats      = new Set(Object.keys(ojcByCategory))
  const etcCats      = new Set(Object.keys(etcByCategory))
  const stockOnlyCats = Object.keys(ojcStock).filter(k => !ojcCats.has(k) && !etcCats.has(k) && ojcStock[k] > 0)

  // OJC 전체 합계
  const ojcTotalAnnuals      = years.reduce<Record<string, number>>((acc, yr) => { acc[yr] = Object.values(ojcByCategory).reduce((s, d) => s + (d.annuals[yr] ?? 0), 0); return acc }, {})
  const ojcTotalPriceAnnuals = years.reduce<Record<string, number>>((acc, yr) => { acc[yr] = Object.values(ojcByCategory).reduce((s, d) => s + (d.priceAnnuals[yr] ?? 0), 0); return acc }, {})
  const ojcTotalMonthly      = MONTHS.reduce<Record<string, number>>((acc, m) => { acc[m] = Object.values(ojcByCategory).reduce((s, d) => s + (d.monthlyLatest[m] ?? 0), 0); return acc }, {})
  const ojcTotalPeak         = Math.max(0, ...MONTHS.map(m => ojcTotalMonthly[m]))
  const ojcTotalPeakMonth    = ojcTotalPeak > 0 ? (MONTHS.find(m => ojcTotalMonthly[m] === ojcTotalPeak) ?? '') : ''
  const ojcTotalLatest       = ojcTotalAnnuals[latestYr] ?? 0
  const ojcTotalAvg          = ojcTotalLatest / 12
  const ojcTotalStock        = Object.keys(ojcByCategory).reduce((s, cat) => s + (ojcStock[cat] ?? 0), 0)
  const ojcTotalCoverPeak    = ojcTotalPeak > 0 ? ojcTotalStock / ojcTotalPeak : 0
  const ojcTotalCoverAvg     = ojcTotalAvg  > 0 ? ojcTotalStock / ojcTotalAvg  : 0
  const ojcTotalProducts     = Object.values(ojcByCategory).reduce((s, d) => s + Object.keys(d.products).length, 0)

  return (
    <div className="space-y-3">
      <div className="overflow-x-auto border border-gray-200 rounded-lg">
        <table className="text-sm w-full border-collapse bg-white">
          <thead>
            <tr>
              <th className={`${thBase} text-left sticky left-0 z-20`}>카테고리</th>
              <th className={`${thBase} text-right`}>품목 수</th>
              {years.flatMap(yr => [
                <th key={yr}    className={`${thBase} text-right`}>{yr}년 판매(EA)</th>,
                <th key={`p${yr}`} className={`${thBase} text-right`}>{yr}년 공급가액</th>,
              ])}
              <th className={`${thBase} text-right`}>{latestYr}년<br/>최고월 판매량</th>
              <th className={`${thBase} text-right`}>월평균(EA)</th>
              <th className={`${thBase} text-right`}>현재고(EA)</th>
              <th className={`${thBase} text-right`}>최고월<br/>재고여유</th>
              <th className={`${thBase} text-right`}>평균<br/>재고여유</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {Object.entries(ojcByCategory).map(([cat, data], i) => {
              const latestAnnual  = data.annuals[latestYr] ?? 0
              const monthlyVals   = MONTHS.map(m => data.monthlyLatest[m] ?? 0)
              const peakMonthly   = Math.max(0, ...monthlyVals)
              const peakMonth     = peakMonthly > 0 ? MONTHS.find(m => (data.monthlyLatest[m] ?? 0) === peakMonthly) ?? '' : ''
              const avgMonthly    = latestAnnual / 12
              const stock         = ojcStock[cat] ?? 0
              const coverPeak     = peakMonthly > 0 ? stock / peakMonthly : 0
              const coverAvg      = avgMonthly  > 0 ? stock / avgMonthly  : 0
              const isExpanded    = expanded === cat
              const prodEntries   = Object.entries(data.products)

              const rowBg = i % 2 === 0 ? 'bg-white' : 'bg-gray-50'
              return (
                <Fragment key={cat}>
                  <tr className={rowBg}>
                    <td className={`px-3 py-2 font-semibold text-gray-800 sticky left-0 z-10 ${rowBg}`}>
                      <button
                        onClick={() => handleExpand(cat)}
                        className="text-left hover:text-blue-600 transition flex items-center gap-1"
                      >
                        <span className="text-xs">{isExpanded ? '▼' : '▶'}</span> {cat}
                      </button>
                    </td>
                    <td className="px-3 py-2 text-right text-gray-500">{prodEntries.length}</td>
                    {years.flatMap(yr => [
                      <td key={yr}    className="px-3 py-2 text-right font-mono">{num(data.annuals[yr] ?? 0)}</td>,
                      <td key={`p${yr}`} className="px-3 py-2 text-right text-xs text-gray-600">{fmtPrice(data.priceAnnuals[yr] ?? 0)}</td>,
                    ])}
                    <td className="px-3 py-2 text-right font-mono font-semibold text-orange-700">
                      {peakMonthly > 0 ? <>{peakMonthly.toLocaleString()}<br/><span className="text-xs font-normal text-orange-500">({parseInt(peakMonth)}월)</span></> : '—'}
                    </td>
                    <td className="px-3 py-2 text-right font-mono">
                      {latestAnnual > 0 ? Math.round(avgMonthly).toLocaleString() : '—'}
                    </td>
                    <td className="px-2 py-1 text-right">
                      <input
                        type="number" min="0"
                        value={stock || ''}
                        placeholder="입력"
                        onChange={e => onStockChange(cat, e.target.value)}
                        className="w-24 border border-gray-300 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-blue-400"
                      />
                    </td>
                    <td className={`px-3 py-2 text-right font-semibold ${
                      stock > 0 && peakMonthly > 0
                        ? coverPeak >= 2 ? 'text-green-600' : coverPeak >= 1 ? 'text-yellow-600' : 'text-red-500'
                        : 'text-gray-400'
                    }`}>
                      {stock > 0 && peakMonthly > 0 ? `${dec1(coverPeak)}개월` : '—'}
                    </td>
                    <td className={`px-3 py-2 text-right font-semibold ${
                      stock > 0 && avgMonthly > 0
                        ? coverAvg >= 2 ? 'text-green-600' : coverAvg >= 1 ? 'text-yellow-600' : 'text-red-500'
                        : 'text-gray-400'
                    }`}>
                      {stock > 0 && avgMonthly > 0 ? `${dec1(coverAvg)}개월` : '—'}
                    </td>
                  </tr>

                  {/* 확장 시 월간 + 품목별 상세 */}
                  {isExpanded && (
                    <Fragment key={`${cat}-detail`}>
                      {/* 월간 현황 행 */}
                      <tr className="bg-blue-50/60">
                        <td colSpan={3 + years.length * 2 + 5} className="px-4 py-2">
                          <div className="flex items-center gap-2 mb-1.5">
                            <div className="text-xs font-semibold text-blue-700">월간 판매 현황</div>
                            <div className="flex gap-1">
                              {years.map(yr => (
                                <button
                                  key={yr}
                                  onClick={() => setExpandedYear(yr)}
                                  className={`text-xs px-2 py-0.5 rounded ${expandedYear === yr ? 'bg-blue-600 text-white' : 'bg-white border border-blue-200 text-blue-600 hover:bg-blue-50'}`}
                                >{yr}년</button>
                              ))}
                            </div>
                          </div>
                          <div className="flex gap-1 flex-wrap">
                            {(() => {
                              const byYr = data.monthlyByYear[expandedYear] ?? {}
                              const peakVal = Math.max(0, ...MONTHS.map(m => byYr[m] ?? 0))
                              return MONTHS.map(m => {
                                const v = byYr[m] ?? 0
                                const isPeak = v === peakVal && v > 0
                                return (
                                  <div key={m} className={`text-center rounded px-2 py-1 min-w-[52px] ${isPeak ? 'bg-orange-200 font-bold text-orange-800' : 'bg-white border border-gray-200 text-gray-700'}`}>
                                    <div className="text-[10px] text-gray-400">{parseInt(m)}월</div>
                                    <div className="text-xs">{v > 0 ? v.toLocaleString() : '—'}</div>
                                  </div>
                                )
                              })
                            })()}
                          </div>
                        </td>
                      </tr>
                      {/* 품목별 상세 행들 */}
                      {prodEntries
                        .sort((a, b) => (b[1].annuals[latestYr] ?? 0) - (a[1].annuals[latestYr] ?? 0))
                        .map(([name, prod]) => {
                          const pLatest   = prod.annuals[latestYr] ?? 0
                          const pMonthly  = MONTHS.map(m => prod.monthlyLatest[m] ?? 0)
                          const pPeak     = Math.max(0, ...pMonthly)
                          return (
                            <tr key={name} className="bg-gray-50/80 hover:bg-blue-50/40 transition">
                              <td className="pl-7 pr-3 py-1.5 text-xs text-gray-600 max-w-xs sticky left-0 bg-gray-50" title={name}>
                                <div className="text-[10px] text-blue-400 font-mono">{prod.code || '—'}</div>
                                <div className="truncate">↳ {name}</div>
                              </td>
                              <td className="px-3 py-1.5 text-right text-xs text-gray-400">—</td>
                              {years.flatMap(yr => [
                                <td key={yr}    className="px-3 py-1.5 text-right text-xs font-mono text-gray-600">{num(prod.annuals[yr] ?? 0)}</td>,
                                <td key={`p${yr}`} className="px-3 py-1.5 text-right text-xs text-gray-500">{fmtPrice(prod.priceAnnuals[yr] ?? 0)}</td>,
                              ])}
                              <td className="px-3 py-1.5 text-right text-xs font-mono text-orange-600">{num(pPeak)}</td>
                              <td className="px-3 py-1.5 text-right text-xs font-mono text-gray-500">
                                {pLatest > 0 ? Math.round(pLatest / 12).toLocaleString() : '—'}
                              </td>
                              <td colSpan={3} className="px-3 py-1.5 text-xs text-gray-300 text-right">—</td>
                            </tr>
                          )
                        })}
                    </Fragment>
                  )}
                </Fragment>
              )
            })}

            {Object.keys(ojcByCategory).length > 0 && (
              <tr className="bg-[#dce9f5] font-semibold border-t-2 border-[#2E75B6]">
                <td className="px-3 py-2 text-[#1F3864] sticky left-0 bg-[#dce9f5] text-sm">OJC 합계</td>
                <td className="px-3 py-2 text-right text-gray-600 text-xs">{ojcTotalProducts}</td>
                {years.flatMap(yr => [
                  <td key={yr}    className="px-3 py-2 text-right font-mono text-[#1F3864]">{num(ojcTotalAnnuals[yr] ?? 0)}</td>,
                  <td key={`p${yr}`} className="px-3 py-2 text-right text-xs text-gray-600">{fmtPrice(ojcTotalPriceAnnuals[yr] ?? 0)}</td>,
                ])}
                <td className="px-3 py-2 text-right font-mono font-semibold text-orange-700">
                  {ojcTotalPeak > 0 ? <>{ojcTotalPeak.toLocaleString()}<br/><span className="text-xs font-normal text-orange-500">({parseInt(ojcTotalPeakMonth)}월)</span></> : '—'}
                </td>
                <td className="px-3 py-2 text-right font-mono">
                  {ojcTotalLatest > 0 ? Math.round(ojcTotalAvg).toLocaleString() : '—'}
                </td>
                <td className="px-3 py-2 text-right font-mono text-[#1F3864]">{ojcTotalStock > 0 ? ojcTotalStock.toLocaleString() : '—'}</td>
                <td className={`px-3 py-2 text-right font-semibold ${ojcTotalStock > 0 && ojcTotalPeak > 0 ? coverFgClass(ojcTotalCoverPeak) : 'text-gray-400'}`}>
                  {ojcTotalStock > 0 && ojcTotalPeak > 0 ? `${dec1(ojcTotalCoverPeak)}개월` : '—'}
                </td>
                <td className={`px-3 py-2 text-right font-semibold ${ojcTotalStock > 0 && ojcTotalAvg > 0 ? coverFgClass(ojcTotalCoverAvg) : 'text-gray-400'}`}>
                  {ojcTotalStock > 0 && ojcTotalAvg > 0 ? `${dec1(ojcTotalCoverAvg)}개월` : '—'}
                </td>
              </tr>
            )}

            {Object.keys(ojcByCategory).length === 0 && (
              <tr>
                <td colSpan={5 + years.length * 2 + 3} className="text-center text-gray-400 py-10">
                  OJC 품목이 감지되지 않았습니다. (품목명 앞부분: OJC-, SOJC-, DOJC-, MOJC-, DROP-CABLE, PIGTAIL- 등)
                </td>
              </tr>
            )}

            {/* ── 기타 품목 판매 현황 (ETC 분류, 실제 판매 데이터 있음) ── */}
            {Object.keys(etcByCategory).length > 0 && (
              <>
                <tr>
                  <td colSpan={5 + years.length * 2 + 3}
                    className="px-3 py-1.5 text-xs font-semibold text-purple-700 bg-purple-50 border-t-2 border-purple-200">
                    기타 품목 판매 현황
                  </td>
                </tr>
                {Object.entries(etcByCategory).map(([cat, data], i) => {
                  const latestAnnual = data.annuals[latestYr] ?? 0
                  const monthlyVals  = MONTHS.map(m => data.monthlyLatest[m] ?? 0)
                  const peakMonthly  = Math.max(0, ...monthlyVals)
                  const peakMonth    = peakMonthly > 0 ? MONTHS.find(m => (data.monthlyLatest[m] ?? 0) === peakMonthly) ?? '' : ''
                  const avgMonthly   = latestAnnual / 12
                  const stock        = ojcStock[cat] ?? 0
                  const coverPeak    = peakMonthly > 0 ? stock / peakMonthly : 0
                  const coverAvg     = avgMonthly  > 0 ? stock / avgMonthly  : 0
                  const isExpanded   = expanded === `etc-${cat}`
                  const prodEntries  = Object.entries(data.products)
                  const etcRowBg = i % 2 === 0 ? 'bg-white' : 'bg-purple-50'
                  return (
                    <Fragment key={`etc-${cat}`}>
                      <tr className={etcRowBg}>
                        <td className={`px-3 py-2 font-semibold text-purple-800 sticky left-0 z-10 ${etcRowBg}`}>
                          <button
                            onClick={() => handleExpand(`etc-${cat}`)}
                            className="text-left hover:text-purple-600 transition flex items-center gap-1"
                          >
                            <span className="text-xs">{isExpanded ? '▼' : '▶'}</span> {cat}
                          </button>
                        </td>
                        <td className="px-3 py-2 text-right text-gray-500">{prodEntries.length}</td>
                        {years.flatMap(yr => [
                          <td key={yr}    className="px-3 py-2 text-right font-mono">{num(data.annuals[yr] ?? 0)}</td>,
                          <td key={`p${yr}`} className="px-3 py-2 text-right text-xs text-gray-600">{fmtPrice(data.priceAnnuals[yr] ?? 0)}</td>,
                        ])}
                        <td className="px-3 py-2 text-right font-mono font-semibold text-orange-700">
                          {peakMonthly > 0 ? <>{peakMonthly.toLocaleString()}<br/><span className="text-xs font-normal text-orange-500">({parseInt(peakMonth)}월)</span></> : '—'}
                        </td>
                        <td className="px-3 py-2 text-right font-mono">
                          {latestAnnual > 0 ? Math.round(avgMonthly).toLocaleString() : '—'}
                        </td>
                        <td className="px-2 py-1 text-right">
                          <input
                            type="number" min="0"
                            value={stock || ''}
                            placeholder="입력"
                            onChange={e => onStockChange(cat, e.target.value)}
                            className="w-24 border border-purple-200 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-purple-400"
                          />
                        </td>
                        <td className={`px-3 py-2 text-right font-semibold ${
                          stock > 0 && peakMonthly > 0
                            ? coverPeak >= 2 ? 'text-green-600' : coverPeak >= 1 ? 'text-yellow-600' : 'text-red-500'
                            : 'text-gray-400'
                        }`}>
                          {stock > 0 && peakMonthly > 0 ? `${dec1(coverPeak)}개월` : '—'}
                        </td>
                        <td className={`px-3 py-2 text-right font-semibold ${
                          stock > 0 && avgMonthly > 0
                            ? coverAvg >= 2 ? 'text-green-600' : coverAvg >= 1 ? 'text-yellow-600' : 'text-red-500'
                            : 'text-gray-400'
                        }`}>
                          {stock > 0 && avgMonthly > 0 ? `${dec1(coverAvg)}개월` : '—'}
                        </td>
                      </tr>
                      {isExpanded && (
                        <Fragment key={`etc-${cat}-detail`}>
                          <tr className="bg-purple-50/40">
                            <td colSpan={3 + years.length * 2 + 5} className="px-4 py-2">
                              <div className="flex items-center gap-2 mb-1.5">
                                <div className="text-xs font-semibold text-purple-700">월간 판매 현황</div>
                                <div className="flex gap-1">
                                  {years.map(yr => (
                                    <button
                                      key={yr}
                                      onClick={() => setExpandedYear(yr)}
                                      className={`text-xs px-2 py-0.5 rounded ${expandedYear === yr ? 'bg-purple-600 text-white' : 'bg-white border border-purple-200 text-purple-600 hover:bg-purple-50'}`}
                                    >{yr}년</button>
                                  ))}
                                </div>
                              </div>
                              <div className="flex gap-1 flex-wrap">
                                {(() => {
                                  const byYr = data.monthlyByYear[expandedYear] ?? {}
                                  const peakVal = Math.max(0, ...MONTHS.map(m => byYr[m] ?? 0))
                                  return MONTHS.map(m => {
                                    const v = byYr[m] ?? 0
                                    const isPeak = v === peakVal && v > 0
                                    return (
                                      <div key={m} className={`text-center rounded px-2 py-1 min-w-[52px] ${isPeak ? 'bg-orange-200 font-bold text-orange-800' : 'bg-white border border-gray-200 text-gray-700'}`}>
                                        <div className="text-[10px] text-gray-400">{parseInt(m)}월</div>
                                        <div className="text-xs">{v > 0 ? v.toLocaleString() : '—'}</div>
                                      </div>
                                    )
                                  })
                                })()}
                              </div>
                            </td>
                          </tr>
                          {prodEntries
                            .sort((a, b) => (b[1].annuals[latestYr] ?? 0) - (a[1].annuals[latestYr] ?? 0))
                            .map(([name, prod]) => {
                              const pLatest = prod.annuals[latestYr] ?? 0
                              const pPeak   = Math.max(0, ...MONTHS.map(m => prod.monthlyLatest[m] ?? 0))
                              return (
                                <tr key={name} className="bg-purple-50/10 hover:bg-purple-50/40 transition">
                                  <td className="pl-7 pr-3 py-1.5 text-xs text-purple-700 max-w-xs sticky left-0 bg-white" title={name}>
                                    <div className="text-[10px] text-purple-400 font-mono">{prod.code || '—'}</div>
                                    <div className="truncate">↳ {name}</div>
                                  </td>
                                  <td className="px-3 py-1.5 text-right text-xs text-gray-400">—</td>
                                  {years.flatMap(yr => [
                                    <td key={yr}    className="px-3 py-1.5 text-right text-xs font-mono text-gray-600">{num(prod.annuals[yr] ?? 0)}</td>,
                                    <td key={`p${yr}`} className="px-3 py-1.5 text-right text-xs text-gray-500">{fmtPrice(prod.priceAnnuals[yr] ?? 0)}</td>,
                                  ])}
                                  <td className="px-3 py-1.5 text-right text-xs font-mono text-orange-600">{num(pPeak)}</td>
                                  <td className="px-3 py-1.5 text-right text-xs font-mono text-gray-500">
                                    {pLatest > 0 ? Math.round(pLatest / 12).toLocaleString() : '—'}
                                  </td>
                                  <td colSpan={3} className="px-3 py-1.5 text-xs text-gray-300 text-right">—</td>
                                </tr>
                              )
                            })}
                        </Fragment>
                      )}
                    </Fragment>
                  )
                })}
              </>
            )}

            {/* ── 재고만 있고 판매 데이터 없는 항목 ── */}
            {stockOnlyCats.length > 0 && (
              <>
                <tr>
                  <td colSpan={5 + years.length * 2 + 3}
                    className="px-3 py-1.5 text-xs font-semibold text-gray-500 bg-gray-50 border-t border-gray-200">
                    재고만 있음 (판매 데이터 없음)
                  </td>
                </tr>
                {stockOnlyCats.map(cat => {
                  const stock = ojcStock[cat] ?? 0
                  return (
                    <tr key={cat} className="bg-gray-50 hover:bg-gray-100">
                      <td className="px-3 py-2 font-semibold text-gray-500 text-sm sticky left-0 bg-gray-50">{cat}</td>
                      <td className="px-3 py-2 text-right text-gray-300 text-xs">—</td>
                      {years.flatMap(yr => [
                        <td key={yr}    className="px-3 py-2 text-right text-gray-300 text-xs">—</td>,
                        <td key={`p${yr}`} className="px-3 py-2 text-right text-gray-300 text-xs">—</td>,
                      ])}
                      <td className="px-3 py-2 text-right text-gray-300 text-xs">—</td>
                      <td className="px-3 py-2 text-right text-gray-300 text-xs">—</td>
                      <td className="px-2 py-1 text-right">
                        <input
                          type="number" min="0"
                          value={stock || ''}
                          placeholder="입력"
                          onChange={e => onStockChange(cat, e.target.value)}
                          className="w-24 border border-gray-200 rounded px-2 py-1 text-right text-sm focus:outline-none focus:border-gray-400"
                        />
                      </td>
                      <td className="px-3 py-2 text-right text-gray-300 text-xs">—</td>
                      <td className="px-3 py-2 text-right text-gray-300 text-xs">—</td>
                    </tr>
                  )
                })}
              </>
            )}
          </tbody>
        </table>
      </div>
      <div className="text-xs text-gray-400 space-y-0.5">
        <div>※ 최고월 판매량 / 월평균은 <b>{latestYr}년</b> 기준 | 최고월 재고여유·평균 재고여유 = 현재고 ÷ 해당 월 수요 (개월)</div>
        <div>※ 현재고는 OJC 완제품 재고 기준으로 직접 입력하세요 (입력값은 브라우저에 자동 저장됩니다)</div>
        <div>※ 카테고리 행 클릭 시 월간 분포 및 품목별 상세를 확인할 수 있습니다</div>
      </div>
    </div>
  )
}

// ── ② 거래처별 탑3 ──────────────────────────────────────────────────
function downloadCustomerTop3(
  customerTop3: Record<string, Array<{ name: string; code: string; qty: number; price: number }>>,
  years: string[],
) {
  const header = ['순위', '거래처명', '총구매(EA)', 'TOP1 품목코드', 'TOP1 품목명', 'TOP1 수량(EA)', 'TOP1 공급가액', 'TOP2 품목코드', 'TOP2 품목명', 'TOP2 수량(EA)', 'TOP2 공급가액', 'TOP3 품목코드', 'TOP3 품목명', 'TOP3 수량(EA)', 'TOP3 공급가액']
  const rows: unknown[][] = [header]
  Object.entries(customerTop3).forEach(([cust, top3], i) => {
    const total = top3.reduce((s, x) => s + x.qty, 0)
    rows.push([
      i + 1, cust, total,
      top3[0]?.code ?? null, top3[0]?.name ?? null, top3[0]?.qty ?? null, top3[0]?.price || null,
      top3[1]?.code ?? null, top3[1]?.name ?? null, top3[1]?.qty ?? null, top3[1]?.price || null,
      top3[2]?.code ?? null, top3[2]?.name ?? null, top3[2]?.qty ?? null, top3[2]?.price || null,
    ])
  })
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '거래처별탑3')
  const note = [[`전체 기간(${years.map(y => '20' + y + '년').join('~')}) 누적 기준`]]
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(note), '안내')
  downloadXlsx(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer, `거래처별탑3_${today()}.xlsx`)
}

const TOP3_DISPLAY_LIMIT = 20

function buildTop3(rawRows: import('../lib/parse/parseDetailedSales').DetailedSalesRow[], yr: string | 'cumul', sortBy: 'qty' | 'price' = 'qty') {
  const custMap: Record<string, Record<string, { code: string; qty: number; price: number }>> = {}
  for (const row of rawRows) {
    if (isTop3Excluded(row.name)) continue
    if (yr !== 'cumul' && row.year !== yr) continue
    const cust = row.customer || '(미상)'
    if (!custMap[cust]) custMap[cust] = {}
    if (!custMap[cust][row.name]) custMap[cust][row.name] = { code: row.code, qty: 0, price: 0 }
    custMap[cust][row.name].qty   += row.qty
    custMap[cust][row.name].price += row.price
  }
  const result: Record<string, Array<{ name: string; code: string; qty: number; price: number }>> = {}
  for (const [cust, products] of Object.entries(custMap)) {
    result[cust] = Object.entries(products)
      .sort((a, b) => sortBy === 'price' ? b[1].price - a[1].price : b[1].qty - a[1].qty)
      .slice(0, 3)
      .map(([name, d]) => ({ name, code: d.code, qty: d.qty, price: d.price }))
  }
  return Object.fromEntries(
    Object.entries(result).sort((a, b) =>
      sortBy === 'price'
        ? b[1].reduce((s, x) => s + x.price, 0) - a[1].reduce((s, x) => s + x.price, 0)
        : b[1].reduce((s, x) => s + x.qty, 0) - a[1].reduce((s, x) => s + x.qty, 0)
    )
  )
}

function CustomerTop3View({
  customerTop3, years, rawRows, label,
}: {
  customerTop3: Record<string, Array<{ name: string; code: string; qty: number; price: number }>>
  years:        string[]
  rawRows:      import('../lib/parse/parseDetailedSales').DetailedSalesRow[]
  label?:       string
}) {
  const latestYr = years[years.length - 1] ?? 'cumul'
  const [activeTab, setActiveTab]   = useState<string>(latestYr)
  const [search, setSearch]         = useState('')
  const [showAll, setShowAll]       = useState(false)
  const [sortMode, setSortMode]     = useState<'qty' | 'price'>('qty')

  const computedTop3 = useMemo(
    () => buildTop3(rawRows, activeTab === 'cumul' ? 'cumul' : activeTab, sortMode),
    [activeTab, rawRows, sortMode]
  )

  const customers = Object.entries(computedTop3)
  const filtered  = customers.filter(([cust]) =>
    !search || cust.toLowerCase().includes(search.toLowerCase())
  )
  const displayed = showAll ? filtered : filtered.slice(0, TOP3_DISPLAY_LIMIT)

  const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'
  const tabCls = (id: string) =>
    `px-3 py-1.5 text-xs font-semibold rounded-t border-b-2 transition whitespace-nowrap ${
      activeTab === id
        ? 'border-[#2E75B6] text-[#2E75B6] bg-white'
        : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
    }`

  if (!Object.keys(customerTop3).length) {
    return (
      <div className="text-center text-gray-400 py-10 text-sm">
        거래처 데이터가 없습니다. 파일에 <code className="bg-gray-100 px-1 rounded">거래처명</code> 컬럼이 있어야 합니다.
      </div>
    )
  }

  return (
    <div className="space-y-3">
      {label && (
        <div className="text-xs font-semibold text-white bg-[#2E75B6] rounded px-3 py-1 w-fit">
          {label === '내자' ? '🏠 내자 거래처' : '✈️ 외자 거래처'}
        </div>
      )}
      {/* 탭 + 검색/다운로드 */}
      <div className="flex justify-between items-end gap-2 flex-wrap border-b border-gray-200">
        <div className="flex gap-0.5 items-end">
          {years.map(yr => (
            <button key={yr} className={tabCls(yr)} onClick={() => { setActiveTab(yr); setShowAll(false) }}>
              20{yr}년
            </button>
          ))}
          <div className="w-px h-5 bg-gray-300 mx-1 self-center" />
          <button className={tabCls('cumul')} onClick={() => { setActiveTab('cumul'); setShowAll(false) }}>
            전체 누적
          </button>
        </div>
        <div className="flex gap-2 items-center pb-1">
          <div className="flex gap-0 border border-gray-200 rounded overflow-hidden text-xs">
            {(['qty', 'price'] as const).map(mode => (
              <button
                key={mode}
                onClick={() => { setSortMode(mode); setShowAll(false) }}
                className={`px-3 py-1.5 font-medium transition ${sortMode === mode ? 'bg-[#2E75B6] text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`}
              >
                {mode === 'qty' ? '📦 판매량순' : '💰 공급가액순'}
              </button>
            ))}
          </div>
          <span className="text-xs text-gray-500">{filtered.length}개 거래처</span>
          <input
            type="text"
            placeholder="거래처 검색..."
            value={search}
            onChange={e => setSearch(e.target.value)}
            className="border border-gray-300 rounded px-3 py-1 text-sm focus:outline-none focus:border-blue-400 w-40"
          />
          <button
            onClick={() => downloadCustomerTop3(customerTop3, years)}
            className="px-3 py-1.5 text-sm bg-[#2E75B6] hover:bg-[#1a5a9e] text-white font-semibold rounded transition whitespace-nowrap"
          >
            📥 Excel 저장
          </button>
        </div>
      </div>

      <div className="overflow-x-auto border border-gray-200 rounded-lg">
        <table className="text-sm w-full border-collapse bg-white">
          <thead>
            <tr>
              <th className={`${thBase} text-right`}>순위</th>
              <th className={`${thBase} text-left`}>거래처명</th>
              <th className={`${thBase} text-right`}>{sortMode === 'qty' ? '총구매(EA)' : '공급가액 합산'}</th>
              {[1, 2, 3].map(rank => (
                <Fragment key={rank}>
                  <th className={`${thBase} text-left`}>TOP{rank} 품목명</th>
                  {sortMode === 'qty' ? (
                    <>
                      <th className={`${thBase} text-right`}>수량(EA)</th>
                      <th className={`${thBase} text-right`}>수량점유율</th>
                      <th className={`${thBase} text-right`}>가액점유율</th>
                    </>
                  ) : (
                    <>
                      <th className={`${thBase} text-right`}>공급가액</th>
                      <th className={`${thBase} text-right`}>가액점유율</th>
                      <th className={`${thBase} text-right`}>수량점유율</th>
                    </>
                  )}
                </Fragment>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {displayed.map(([cust, top3], i) => {
              const qtyTotal   = top3.reduce((s, x) => s + x.qty, 0)
              const priceTotal = top3.reduce((s, x) => s + x.price, 0)
              return (
                <tr key={cust} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                  <td className="px-3 py-2 text-right text-gray-400 text-xs">{i + 1}</td>
                  <td className="px-3 py-2 font-semibold text-gray-800">{cust}</td>
                  <td className="px-3 py-2 text-right font-mono font-bold text-gray-900">
                    {sortMode === 'qty' ? qtyTotal.toLocaleString() : fmtPrice(priceTotal)}
                  </td>
                  {[0, 1, 2].map(rank => {
                    const item     = top3[rank]
                    const qtyPct   = item && qtyTotal   > 0 ? Math.round(item.qty   / qtyTotal   * 100) : null
                    const pricePct = item && priceTotal > 0 ? Math.round(item.price / priceTotal * 100) : null
                    const pctCls   = (v: number | null) => v === null ? 'text-gray-300' : `font-semibold ${v >= 60 ? 'text-orange-600' : v >= 30 ? 'text-blue-600' : 'text-gray-500'}`
                    return (
                      <Fragment key={rank}>
                        <td className="px-3 py-2 max-w-[200px] text-sm text-gray-700" title={item?.name ?? ''}>
                          <div className="text-[10px] text-gray-400 font-mono">{item?.code || '—'}</div>
                          <div className="truncate">{item?.name ?? <span className="text-gray-300">—</span>}</div>
                        </td>
                        {sortMode === 'qty' ? (
                          <>
                            <td className="px-3 py-2 text-right font-mono text-sm">
                              {item ? item.qty.toLocaleString() : <span className="text-gray-300">—</span>}
                            </td>
                            <td className="px-3 py-2 text-right text-xs">
                              <span className={pctCls(qtyPct)}>{qtyPct !== null ? `${qtyPct}%` : '—'}</span>
                            </td>
                            <td className="px-3 py-2 text-right text-xs">
                              <span className={pctCls(pricePct)}>{pricePct !== null ? `${pricePct}%` : '—'}</span>
                            </td>
                          </>
                        ) : (
                          <>
                            <td className="px-3 py-2 text-right text-xs text-gray-600">
                              {item ? fmtPrice(item.price) : <span className="text-gray-300">—</span>}
                            </td>
                            <td className="px-3 py-2 text-right text-xs">
                              <span className={pctCls(pricePct)}>{pricePct !== null ? `${pricePct}%` : '—'}</span>
                            </td>
                            <td className="px-3 py-2 text-right text-xs">
                              <span className={pctCls(qtyPct)}>{qtyPct !== null ? `${qtyPct}%` : '—'}</span>
                            </td>
                          </>
                        )}
                      </Fragment>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>

      {/* 더 보기 / 접기 */}
      {filtered.length > TOP3_DISPLAY_LIMIT && (
        <button
          onClick={() => setShowAll(v => !v)}
          className="w-full text-xs text-blue-600 hover:text-blue-800 py-1.5 border border-dashed border-blue-200 rounded hover:bg-blue-50 transition"
        >
          {showAll
            ? `▲ 상위 ${TOP3_DISPLAY_LIMIT}개만 보기`
            : `▼ 전체 보기 (${filtered.length - TOP3_DISPLAY_LIMIT}개 더)`}
        </button>
      )}

      <div className="text-xs text-gray-400">
        ※ {activeTab === 'cumul' ? '전체 기간 누적' : `20${activeTab}년`} 수량 기준 | 거래처당 최다 구매 품목 탑3 | 점유율: 해당 품목이 거래처 총구매량의 비중 (주황 ≥60%, 파랑 ≥30%)
      </div>
    </div>
  )
}

// ── ③ 거래처별 OJC 탑3 ──────────────────────────────────────────────
function OjcCustomerTop3View({
  ojcRows, years,
}: {
  ojcRows: import('../lib/parse/parseDetailedSales').DetailedSalesRow[]
  years:   string[]
}) {
  const latestYr = years[years.length - 1] ?? 'cumul'
  const [activeTab, setActiveTab] = useState<string>(latestYr)
  const [search, setSearch]       = useState('')
  const [showAll, setShowAll]     = useState(false)

  const computedTop3 = useMemo(
    () => buildTop3(ojcRows, activeTab === 'cumul' ? 'cumul' : activeTab),
    [activeTab, ojcRows]
  )

  const customers = Object.entries(computedTop3)
  const filtered  = customers.filter(([cust]) =>
    !search || cust.toLowerCase().includes(search.toLowerCase())
  )
  const displayed = showAll ? filtered : filtered.slice(0, TOP3_DISPLAY_LIMIT)

  const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'
  const tabCls = (id: string) =>
    `px-3 py-1.5 text-xs font-semibold rounded-t border-b-2 transition whitespace-nowrap ${
      activeTab === id
        ? 'border-[#2E75B6] text-[#2E75B6] bg-white'
        : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
    }`

  if (!ojcRows.length) {
    return <div className="text-center text-gray-400 py-10 text-sm">OJC 품목 데이터가 없습니다.</div>
  }

  return (
    <div className="space-y-3">
      {/* 탭 + 검색 */}
      <div className="flex justify-between items-end gap-2 flex-wrap border-b border-gray-200">
        <div className="flex gap-0.5 items-end">
          {years.map(yr => (
            <button key={yr} className={tabCls(yr)} onClick={() => { setActiveTab(yr); setShowAll(false) }}>
              20{yr}년
            </button>
          ))}
          <div className="w-px h-5 bg-gray-300 mx-1 self-center" />
          <button className={tabCls('cumul')} onClick={() => { setActiveTab('cumul'); setShowAll(false) }}>
            전체 누적
          </button>
        </div>
        <div className="flex gap-2 items-center pb-1">
          <span className="text-xs text-gray-500">{filtered.length}개 거래처</span>
          <input
            type="text"
            placeholder="거래처 검색..."
            value={search}
            onChange={e => setSearch(e.target.value)}
            className="border border-gray-300 rounded px-3 py-1 text-sm focus:outline-none focus:border-blue-400 w-40"
          />
        </div>
      </div>

      <div className="overflow-x-auto border border-gray-200 rounded-lg">
        <table className="text-sm w-full border-collapse bg-white">
          <thead>
            <tr>
              <th className={`${thBase} text-right`}>순위</th>
              <th className={`${thBase} text-left`}>거래처명</th>
              <th className={`${thBase} text-right`}>OJC 총구매(EA)</th>
              {[1, 2, 3].map(rank => (
                <Fragment key={rank}>
                  <th className={`${thBase} text-left`}>TOP{rank} 품목명</th>
                  <th className={`${thBase} text-center`}>OJC 분류</th>
                  <th className={`${thBase} text-right`}>수량(EA)</th>
                  <th className={`${thBase} text-right`}>수량점유율</th>
                  <th className={`${thBase} text-right`}>공급가액</th>
                  <th className={`${thBase} text-right`}>가액점유율</th>
                </Fragment>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {displayed.map(([cust, top3], i) => {
              const qtyTotal   = top3.reduce((s, x) => s + x.qty, 0)
              const priceTotal = top3.reduce((s, x) => s + x.price, 0)
              return (
                <tr key={cust} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                  <td className="px-3 py-2 text-right text-gray-400 text-xs">{i + 1}</td>
                  <td className="px-3 py-2 font-semibold text-gray-800">{cust}</td>
                  <td className="px-3 py-2 text-right font-mono font-bold text-[#1F3864]">
                    {qtyTotal.toLocaleString()}
                  </td>
                  {[0, 1, 2].map(rank => {
                    const item     = top3[rank]
                    const ojcCat   = item ? classifyOjc(item.name) : null
                    const qtyPct   = item && qtyTotal   > 0 ? Math.round(item.qty   / qtyTotal   * 100) : null
                    const pricePct = item && priceTotal > 0 ? Math.round(item.price / priceTotal * 100) : null
                    const pctCls   = (v: number | null) => v === null ? 'text-gray-300' : `font-semibold ${v >= 60 ? 'text-orange-600' : v >= 30 ? 'text-blue-600' : 'text-gray-500'}`
                    return (
                      <Fragment key={rank}>
                        <td className="px-3 py-2 max-w-[200px] text-sm text-gray-700" title={item?.name ?? ''}>
                          <div className="text-[10px] text-gray-400 font-mono">{item?.code || '—'}</div>
                          <div className="truncate">{item?.name ?? <span className="text-gray-300">—</span>}</div>
                        </td>
                        <td className="px-3 py-2 text-center">
                          {ojcCat
                            ? <span className="text-[10px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-medium whitespace-nowrap">{ojcCat}</span>
                            : <span className="text-gray-300">—</span>}
                        </td>
                        <td className="px-3 py-2 text-right font-mono text-sm">
                          {item ? item.qty.toLocaleString() : <span className="text-gray-300">—</span>}
                        </td>
                        <td className="px-3 py-2 text-right text-xs">
                          <span className={pctCls(qtyPct)}>{qtyPct !== null ? `${qtyPct}%` : '—'}</span>
                        </td>
                        <td className="px-3 py-2 text-right text-xs text-gray-600">
                          {item ? fmtPrice(item.price) : <span className="text-gray-300">—</span>}
                        </td>
                        <td className="px-3 py-2 text-right text-xs">
                          <span className={pctCls(pricePct)}>{pricePct !== null ? `${pricePct}%` : '—'}</span>
                        </td>
                      </Fragment>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>

      {filtered.length > TOP3_DISPLAY_LIMIT && (
        <button
          onClick={() => setShowAll(v => !v)}
          className="w-full text-xs text-blue-600 hover:text-blue-800 py-1.5 border border-dashed border-blue-200 rounded hover:bg-blue-50 transition"
        >
          {showAll
            ? `▲ 상위 ${TOP3_DISPLAY_LIMIT}개만 보기`
            : `▼ 전체 보기 (${filtered.length - TOP3_DISPLAY_LIMIT}개 더)`}
        </button>
      )}

      <div className="text-xs text-gray-400">
        ※ {activeTab === 'cumul' ? '전체 기간 누적' : `20${activeTab}년`} OJC 수량 기준 | 거래처당 OJC 품목 최다 구매 탑3 | 점유율: 해당 품목이 거래처 OJC 총구매량의 비중
      </div>
    </div>
  )
}

// ── ⓪ 요약 대시보드 ────────────────────────────────────────────────────
function SalesSummaryView({
  ojcByCategory, etcByCategory, customerTop3, fullProducts, years, latestYr,
}: {
  ojcByCategory:  Record<string, CategoryData>
  etcByCategory:  Record<string, CategoryData>
  customerTop3:   Record<string, Array<{ name: string; code: string; qty: number; price: number }>>
  fullProducts:   Array<{ name: string; code: string; ojcCat: string | null; etcCat: string | null; annuals: Record<string, number>; priceAnnuals: Record<string, number> }>
  years:          string[]
  latestYr:       string
}) {
  const prevYr = years.length >= 2 ? years[years.length - 2] : null

  const ojcTotal     = Object.values(ojcByCategory).reduce((s, d) => s + (d.annuals[latestYr] ?? 0), 0)
  const ojcTotalPrev = prevYr ? Object.values(ojcByCategory).reduce((s, d) => s + (d.annuals[prevYr] ?? 0), 0) : 0
  const ojcYoY       = prevYr && ojcTotalPrev > 0 ? (ojcTotal - ojcTotalPrev) / ojcTotalPrev : null
  const etcTotal     = Object.values(etcByCategory).reduce((s, d) => s + (d.annuals[latestYr] ?? 0), 0)
  const custCount    = Object.keys(customerTop3).length
  const productCount = fullProducts.length

  const thB = 'px-3 py-1.5 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-50 whitespace-nowrap text-right'
  const ojcCatsSorted = Object.entries(ojcByCategory).sort((a, b) => (b[1].annuals[latestYr] ?? 0) - (a[1].annuals[latestYr] ?? 0))
  const top5Customers = Object.entries(customerTop3).slice(0, 5)

  const KpiCard = ({ label, value, sub, valueClass }: { label: string; value: string; sub?: string; valueClass?: string }) => (
    <div className="bg-white border border-gray-200 rounded-lg px-4 py-3 flex flex-col gap-0.5">
      <div className="text-xs text-gray-500">{label}</div>
      <div className={`text-xl font-bold ${valueClass ?? 'text-[#1F3864]'}`}>{value}</div>
      {sub && <div className="text-xs text-gray-400">{sub}</div>}
    </div>
  )

  const yoyColor = ojcYoY === null ? 'text-gray-400' : ojcYoY >= 0 ? 'text-green-600' : 'text-red-500'
  const yoyText  = ojcYoY === null ? '—' : `${ojcYoY >= 0 ? '+' : ''}${(ojcYoY * 100).toFixed(1)}%`

  return (
    <div className="space-y-4">
      {/* KPI 카드 */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <KpiCard label={`20${latestYr}년 OJC 총판매`} value={`${ojcTotal.toLocaleString()} EA`} sub="OJC 전체 카테고리 합산" valueClass="text-[#2E75B6]" />
        <KpiCard label="전년 대비 (OJC)" value={yoyText} sub={prevYr ? `vs 20${prevYr}년` : '비교 연도 없음'} valueClass={yoyColor} />
        <KpiCard label="거래처 수 (누적)" value={`${custCount}개`} sub="전체 기간 구매 이력" />
        <KpiCard label="전체 품목 수" value={`${productCount}종`} sub={`OJC ${Object.values(ojcByCategory).reduce((s, d) => s + Object.keys(d.products).length, 0)}종 + 기타`} />
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-5 gap-4">
        {/* OJC 카테고리 × 연도 */}
        <div className="xl:col-span-3 overflow-x-auto border border-gray-200 rounded-lg">
          <div className="bg-[#1F3864] text-white text-xs font-bold px-3 py-2">OJC 카테고리별 연간 판매 추이</div>
          <table className="text-sm w-full border-collapse bg-white">
            <thead>
              <tr>
                <th className={`${thB} text-left sticky left-0 bg-gray-50`}>카테고리</th>
                {years.map(yr => (
                  <th key={yr} className={thB}>{yr}년 (EA)</th>
                ))}
                {prevYr && <th className={thB}>전년비</th>}
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {ojcCatsSorted.map(([cat, data], i) => {
                const curr = data.annuals[latestYr] ?? 0
                const prev = prevYr ? (data.annuals[prevYr] ?? 0) : 0
                const yoy  = prevYr && prev > 0 ? (curr - prev) / prev : null
                return (
                  <tr key={cat} className={i % 2 === 0 ? 'bg-white' : 'bg-blue-50/30'}>
                    <td className="px-3 py-1.5 text-xs font-semibold text-[#2E75B6] sticky left-0 bg-inherit whitespace-nowrap">{cat}</td>
                    {years.map(yr => (
                      <td key={yr} className={`px-3 py-1.5 text-right text-xs tabular-nums ${yr === latestYr ? 'font-bold text-[#1F3864]' : 'text-gray-500'}`}>
                        {(data.annuals[yr] ?? 0) > 0 ? (data.annuals[yr]!).toLocaleString() : '—'}
                      </td>
                    ))}
                    {prevYr && (
                      <td className={`px-3 py-1.5 text-right text-xs font-bold ${yoy === null ? 'text-gray-300' : yoy >= 0 ? 'text-green-600' : 'text-red-500'}`}>
                        {yoy !== null ? `${yoy >= 0 ? '+' : ''}${(yoy * 100).toFixed(1)}%` : '—'}
                      </td>
                    )}
                  </tr>
                )
              })}
              {/* OJC 합계 */}
              <tr className="bg-[#1F3864] text-white font-bold">
                <td className="px-3 py-1.5 text-xs sticky left-0 bg-[#1F3864]">OJC 합계</td>
                {years.map(yr => {
                  const tot = Object.values(ojcByCategory).reduce((s, d) => s + (d.annuals[yr] ?? 0), 0)
                  return <td key={yr} className="px-3 py-1.5 text-right text-xs tabular-nums">{tot > 0 ? tot.toLocaleString() : '—'}</td>
                })}
                {prevYr && (
                  <td className={`px-3 py-1.5 text-right text-xs ${ojcYoY === null ? '' : ojcYoY >= 0 ? 'text-green-300' : 'text-red-300'}`}>
                    {ojcYoY !== null ? `${ojcYoY >= 0 ? '+' : ''}${(ojcYoY * 100).toFixed(1)}%` : '—'}
                  </td>
                )}
              </tr>
              {/* ETC 합계 */}
              {etcTotal > 0 && (
                <tr className="bg-[#7030A0]/10">
                  <td className="px-3 py-1.5 text-xs font-semibold text-[#7030A0] sticky left-0 bg-[#F3EEF8]">기타 품목 합계</td>
                  {years.map(yr => {
                    const tot = Object.values(etcByCategory).reduce((s, d) => s + (d.annuals[yr] ?? 0), 0)
                    return <td key={yr} className={`px-3 py-1.5 text-right text-xs tabular-nums ${yr === latestYr ? 'font-bold text-[#7030A0]' : 'text-gray-500'}`}>{tot > 0 ? tot.toLocaleString() : '—'}</td>
                  })}
                  {prevYr && <td className="px-3 py-1.5 text-right text-xs text-gray-400">—</td>}
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {/* 거래처 TOP5 */}
        <div className="xl:col-span-2 space-y-4">
          <div className="border border-gray-200 rounded-lg overflow-hidden">
            <div className="bg-[#2E75B6] text-white text-xs font-bold px-3 py-2">거래처 TOP5 (전체 누적)</div>
            <table className="text-sm w-full border-collapse bg-white">
              <thead>
                <tr>
                  <th className={`${thB} text-center`}>#</th>
                  <th className={`${thB} text-left`}>거래처명</th>
                  <th className={thB}>총구매 (EA)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {top5Customers.map(([cust, top3], i) => {
                  const total = top3.reduce((s, x) => s + x.qty, 0)
                  return (
                    <tr key={cust} className={i % 2 === 0 ? 'bg-white' : 'bg-blue-50/30'}>
                      <td className="px-3 py-1.5 text-center text-xs font-bold text-gray-400">{i + 1}</td>
                      <td className={`px-3 py-1.5 text-xs font-semibold ${i === 0 ? 'text-red-700' : i === 1 ? 'text-orange-800' : 'text-[#1F3864]'}`}>{cust}</td>
                      <td className="px-3 py-1.5 text-right text-xs font-bold text-[#1F3864] tabular-nums">{total.toLocaleString()}</td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>

          {/* 최신년 OJC 구성 비중 */}
          <div className="border border-gray-200 rounded-lg overflow-hidden">
            <div className="bg-[#7030A0] text-white text-xs font-bold px-3 py-2">20{latestYr}년 OJC 카테고리 비중</div>
            <div className="p-3 space-y-1.5">
              {ojcCatsSorted.slice(0, 6).map(([cat, data]) => {
                const v   = data.annuals[latestYr] ?? 0
                const pct = ojcTotal > 0 ? v / ojcTotal : 0
                return (
                  <div key={cat}>
                    <div className="flex justify-between text-xs mb-0.5">
                      <span className="text-[#2E75B6] font-medium truncate max-w-[60%]">{cat}</span>
                      <span className="text-gray-600 tabular-nums">{(pct * 100).toFixed(1)}%  ({v.toLocaleString()} EA)</span>
                    </div>
                    <div className="h-1.5 bg-gray-100 rounded-full overflow-hidden">
                      <div className="h-full bg-[#2E75B6] rounded-full" style={{ width: `${(pct * 100).toFixed(1)}%` }} />
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        </div>
      </div>

      {/* 판매량 TOP10 */}
      <div className="border border-gray-200 rounded-lg overflow-hidden">
        <div className="bg-[#2E75B6] text-white text-xs font-bold px-3 py-2">전체 판매량 TOP 10 (20{latestYr}년 기준)</div>
        <div className="overflow-x-auto">
          <table className="text-sm w-full border-collapse bg-white">
            <thead>
              <tr>
                <th className={`${thB} text-center`}>#</th>
                <th className={`${thB} text-left`}>품목명</th>
                <th className={`${thB} text-left`}>분류</th>
                {years.map(yr => <th key={yr} className={thB}>{yr}년 (EA)</th>)}
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {fullProducts.slice(0, 10).map((p, i) => (
                <tr key={p.name} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                  <td className="px-3 py-1.5 text-center text-xs font-bold text-gray-400">{i + 1}</td>
                  <td className="px-3 py-1.5 text-xs font-medium max-w-xs truncate" title={p.name}>{p.name}</td>
                  <td className="px-3 py-1.5">
                    <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${
                      p.ojcCat ? 'bg-blue-100 text-blue-700' : p.etcCat ? 'bg-purple-100 text-purple-700' : 'bg-gray-100 text-gray-500'
                    }`}>{p.ojcCat ?? p.etcCat ?? '기타'}</span>
                  </td>
                  {years.map(yr => (
                    <td key={yr} className={`px-3 py-1.5 text-right text-xs tabular-nums ${yr === latestYr ? 'font-bold text-[#1F3864]' : 'text-gray-500'}`}>
                      {(p.annuals[yr] ?? 0) > 0 ? (p.annuals[yr]!).toLocaleString() : '—'}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

// ── ④ 전체 판매량 ──────────────────────────────────────────────────
function downloadFullSales(
  filtered: Array<{ name: string; code: string; ojcCat: string | null; etcCat: string | null; annuals: Record<string, number>; priceAnnuals: Record<string, number> }>,
  years: string[],
) {
  const header = ['품목코드', '품목명', '분류', ...years.flatMap(y => [`${y}년 판매(EA)`, `${y}년 공급가액`]), '합계(EA)']
  const rows: unknown[][] = [header]
  for (const p of filtered) {
    const total = years.reduce((s, yr) => s + (p.annuals[yr] ?? 0), 0)
    rows.push([p.code, p.name, p.ojcCat ?? p.etcCat ?? '기타', ...years.flatMap(yr => [p.annuals[yr] ?? 0, p.priceAnnuals[yr] ?? 0]), total])
  }
  const totals = years.map(yr => filtered.reduce((s, p) => s + (p.annuals[yr] ?? 0), 0))
  rows.push(['합계', '', '', ...totals.flatMap(t => [t, null]), totals.reduce((a, b) => a + b, 0)])
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '전체판매량')
  downloadXlsx(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer, `전체판매량_${today()}.xlsx`)
}

type SortKey = 'name' | 'total' | `annual_${string}` | `price_${string}`

function FullSalesView({
  products, years, latestYr,
}: {
  products:  Array<{ name: string; code: string; ojcCat: string | null; etcCat: string | null; annuals: Record<string, number>; priceAnnuals: Record<string, number> }>
  years:     string[]
  latestYr:  string
}) {
  const [search, setSearch]         = useState('')
  const [filterOjc, setFilterOjc]   = useState<'all' | 'ojc' | 'non'>('all')
  const [sortKey, setSortKey]        = useState<SortKey>(`annual_${latestYr}`)
  const [sortDir, setSortDir]        = useState<'desc' | 'asc'>('desc')

  const ojcCount    = products.filter(p => p.ojcCat !== null).length
  const nonOjcCount = products.length - ojcCount

  const handleSort = (key: SortKey) => {
    if (sortKey === key) setSortDir(d => d === 'desc' ? 'asc' : 'desc')
    else { setSortKey(key); setSortDir('desc') }
  }
  const sortIcon = (key: SortKey) => sortKey === key ? (sortDir === 'desc' ? ' ▼' : ' ▲') : ''

  const withTotal = products.map(p => ({
    ...p,
    total: years.reduce((s, yr) => s + (p.annuals[yr] ?? 0), 0),
  }))

  const filtered = withTotal
    .filter(p => {
      const matchOjc =
        filterOjc === 'all' ||
        (filterOjc === 'ojc' && p.ojcCat !== null) ||
        (filterOjc === 'non' && p.ojcCat === null)
      const matchSearch = !search || p.name.toLowerCase().includes(search.toLowerCase())
      return matchOjc && matchSearch
    })
    .sort((a, b) => {
      let av = 0, bv = 0
      if (sortKey === 'name')  { av = a.name < b.name ? -1 : 1; bv = 0 }
      else if (sortKey === 'total') { av = a.total; bv = b.total }
      else if (sortKey.startsWith('annual_')) { const yr = sortKey.slice(7); av = a.annuals[yr] ?? 0; bv = b.annuals[yr] ?? 0 }
      else if (sortKey.startsWith('price_'))  { const yr = sortKey.slice(6); av = a.priceAnnuals[yr] ?? 0; bv = b.priceAnnuals[yr] ?? 0 }
      if (sortKey === 'name') return sortDir === 'asc' ? av : -av
      return sortDir === 'desc' ? bv - av : av - bv
    })

  // 연도별 합계 (필터 기준)
  const yearTotals = years.reduce<Record<string, number>>((acc, yr) => {
    acc[yr] = filtered.reduce((s, p) => s + (p.annuals[yr] ?? 0), 0)
    return acc
  }, {})
  const grandTotal = filtered.reduce((s, p) => s + p.total, 0)

  const thBase   = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'
  const thSort   = `${thBase} cursor-pointer hover:bg-gray-200 select-none`

  // 전년비 계산
  const growth = (p: typeof filtered[0], yr: string) => {
    const idx = years.indexOf(yr)
    if (idx <= 0) return null
    const prev = p.annuals[years[idx - 1]] ?? 0
    const curr = p.annuals[yr] ?? 0
    if (prev === 0) return null
    return (curr - prev) / prev
  }

  return (
    <div className="space-y-3">
      {/* 필터 */}
      <div className="flex gap-2 items-center flex-wrap">
        <div className="flex gap-1">
          {(['all', 'ojc', 'non'] as const).map(f => (
            <button
              key={f}
              onClick={() => setFilterOjc(f)}
              className={`px-3 py-1 text-xs font-medium rounded border transition ${
                filterOjc === f
                  ? 'bg-[#2E75B6] text-white border-[#2E75B6]'
                  : 'bg-white text-gray-600 border-gray-300 hover:border-[#2E75B6]'
              }`}
            >
              {f === 'all' ? `전체 (${products.length})` : f === 'ojc' ? `OJC (${ojcCount})` : `비OJC (${nonOjcCount})`}
            </button>
          ))}
        </div>
        <div className="ml-auto flex gap-2 items-center">
          <input
            type="text"
            placeholder="품목명 검색..."
            value={search}
            onChange={e => setSearch(e.target.value)}
            className="border border-gray-300 rounded px-3 py-1 text-sm focus:outline-none focus:border-blue-400 w-52"
          />
          <button
            onClick={() => downloadFullSales(filtered, years)}
            className="px-3 py-1.5 text-sm bg-[#2E75B6] hover:bg-[#1a5a9e] text-white font-semibold rounded transition whitespace-nowrap"
          >
            📥 Excel 저장
          </button>
        </div>
      </div>

      <div className="overflow-x-auto border border-gray-200 rounded-lg">
        <table className="text-xs w-full border-collapse bg-white">
          <thead>
            <tr>
              <th className={`${thSort} text-left`} onClick={() => handleSort('name')}>품목명{sortIcon('name')}</th>
              <th className={`${thBase} text-center`}>분류</th>
              {years.map((yr, idx) => (
                <Fragment key={yr}>
                  <th className={`${thSort} text-right`} onClick={() => handleSort(`annual_${yr}`)}>{yr}년 판매(EA){sortIcon(`annual_${yr}`)}</th>
                  {idx > 0 && <th className={`${thBase} text-right`}>전년비</th>}
                  <th className={`${thSort} text-right`} onClick={() => handleSort(`price_${yr}`)}>{yr}년 공급가액{sortIcon(`price_${yr}`)}</th>
                </Fragment>
              ))}
              <th className={`${thSort} text-right`} onClick={() => handleSort('total')}>합계(EA){sortIcon('total')}</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {filtered.map((p, i) => (
              <tr key={p.name} className={i % 2 === 0 ? 'bg-white hover:bg-gray-50' : 'bg-gray-50 hover:bg-gray-100'}>
                <td className="px-3 py-1.5 max-w-sm text-gray-800 text-sm" title={p.name}>
                  <div className="text-[10px] text-gray-400 font-mono">{p.code || '—'}</div>
                  <div className="truncate">{p.name}</div>
                </td>
                <td className="px-3 py-1.5 text-center">
                  <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${
                    p.ojcCat ? 'bg-blue-100 text-blue-700' : p.etcCat ? 'bg-purple-100 text-purple-700' : 'bg-gray-100 text-gray-500'
                  }`}>
                    {p.ojcCat ?? p.etcCat ?? '기타'}
                  </span>
                </td>
                {years.map((yr, idx) => {
                  const g = growth(p, yr)
                  return (
                    <Fragment key={yr}>
                      <td className="px-3 py-1.5 text-right font-mono text-sm">{num(p.annuals[yr] ?? 0)}</td>
                      {idx > 0 && (
                        <td className={`px-3 py-1.5 text-right text-xs font-semibold ${
                          g === null ? 'text-gray-300' : g > 0 ? 'text-green-600' : g < 0 ? 'text-red-500' : 'text-gray-400'
                        }`}>
                          {g === null ? '—' : `${g > 0 ? '+' : ''}${(g * 100).toFixed(1)}%`}
                        </td>
                      )}
                      <td className="px-3 py-1.5 text-right text-xs text-gray-500">{fmtPrice(p.priceAnnuals[yr] ?? 0)}</td>
                    </Fragment>
                  )
                })}
                <td className="px-3 py-1.5 text-right font-mono font-bold text-gray-800">{p.total > 0 ? p.total.toLocaleString() : '—'}</td>
              </tr>
            ))}

            {/* 합계 행 */}
            {filtered.length > 0 && (
              <tr className="bg-gray-100 font-bold border-t-2 border-gray-300">
                <td className="px-3 py-2 text-gray-800">합계</td>
                <td className="px-3 py-2 text-center text-gray-500 text-xs">{filtered.length}종</td>
                {years.map((yr, idx) => (
                  <Fragment key={yr}>
                    <td className="px-3 py-2 text-right font-mono">{num(yearTotals[yr] ?? 0)}</td>
                    {idx > 0 && <td className="px-3 py-2 text-right text-gray-400">—</td>}
                    <td className="px-3 py-2 text-right text-xs text-gray-400">—</td>
                  </Fragment>
                ))}
                <td className="px-3 py-2 text-right font-mono">{grandTotal > 0 ? grandTotal.toLocaleString() : '—'}</td>
              </tr>
            )}

            {filtered.length === 0 && (
              <tr><td colSpan={2 + years.length * 3} className="text-center text-gray-400 py-10">검색 결과 없음</td></tr>
            )}
          </tbody>
        </table>
      </div>
      <div className="text-xs text-gray-400 flex justify-between">
        <span>※ 헤더 클릭으로 정렬 전환 | 전년비: 전년 대비 판매수량 증감률</span>
        <span>{filtered.length.toLocaleString()}개 품목</span>
      </div>
    </div>
  )
}
