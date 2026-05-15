import { useState, useMemo, Fragment } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import { classifyOjc } from '../lib/ojcFilter'
import { parseDetailedSalesFile } from '../lib/parse/parseDetailedSales'
import type { DetailedSalesRow } from '../lib/parse/parseDetailedSales'
import { downloadXlsx, today } from '../lib/download'
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
  if (n.startsWith('SOC'))             return 'SOC 커넥터'
  if (n.startsWith('IJP'))             return 'IJP'
  if (n.startsWith('DISTRIBUTION CABLE')) return 'DISTRIBUTION CABLE'
  if (n.startsWith('DROP OPTICAL CABLE')) return 'DROP OPTICAL CABLE'
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
  if (/HOUSING|HOUSING KIT/.test(n) || name.includes('하우징'))    return '하우징 키트'
  if (/FERRULE/.test(n) || name.includes('페롤'))                  return '페롤'
  if (/BOOT/.test(n) || name.includes('부트'))                     return '부트'
  if (/UTP|모듈러/.test(n) || name.includes('UTP') || name.includes('모듈러')) return 'UTP/통신자재'
  if (/TOOL|STRIPPER|CLEAVER|CUTTER|METER|공구|측정기/.test(n))   return '공구/측정기'
  return null
}

const num      = (v: number) => v === 0 ? '—' : v.toLocaleString()
const dec1     = (v: number) => v.toFixed(1)
const fmtPrice = (v: number) => v === 0 ? '—' : v.toLocaleString() + '원'

const CUSTOMER_TOP3_EXCLUDE = new Set(['칼라 열수축 슬리브', '단심 열수축 슬리브'])

type SubView = 'ojc' | 'customer' | 'full'

// ── 타입 ─────────────────────────────────────────────────────────────
interface ProductData {
  code:          string
  annuals:       Record<string, number>
  priceAnnuals:  Record<string, number>
  monthlyLatest: Record<string, number>
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
  const [running, setRunning]       = useState(false)
  const [downloading, setDownloading] = useState(false)
  const [invLoaded, setInvLoaded] = useState(false)

  // ESZ018R 파싱 상태
  interface InvRow { code: string; name: string; total: number; byWh: Record<string, number> }
  const [eszRows,      setEszRows]      = useState<InvRow[] | null>(null)
  const [eszWarehouses, setEszWarehouses] = useState<string[]>([])
  const [eszSelectedWh, setEszSelectedWh] = useState('')
  const [eszSubTab,    setEszSubTab]    = useState<'total' | 'warehouse'>('total')
  const [subView, setSubView]   = useState<SubView>('ojc')
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

  // OJC 행만 필터링
  const ojcRows = useMemo(() => rawRows.filter(r => classifyOjc(r.name) !== null), [rawRows])

  // OJC 카테고리별 집계
  const ojcByCategory = useMemo<Record<string, CategoryData>>(() => {
    const cats: Record<string, CategoryData> = {}
    for (const row of ojcRows) {
      const cat = classifyOjc(row.name)!
      if (!cats[cat]) cats[cat] = { products: {}, annuals: {}, priceAnnuals: {}, monthlyLatest: {}, monthlyByYear: {} }
      const c = cats[cat]
      if (!c.products[row.name]) c.products[row.name] = { code: row.code, annuals: {}, priceAnnuals: {}, monthlyLatest: {} }
      const p = c.products[row.name]
      p.annuals[row.year]       = (p.annuals[row.year] ?? 0) + row.qty
      p.priceAnnuals[row.year]  = (p.priceAnnuals[row.year] ?? 0) + row.price
      c.annuals[row.year]       = (c.annuals[row.year] ?? 0) + row.qty
      c.priceAnnuals[row.year]  = (c.priceAnnuals[row.year] ?? 0) + row.price
      if (!c.monthlyByYear[row.year]) c.monthlyByYear[row.year] = {}
      c.monthlyByYear[row.year][row.month] = (c.monthlyByYear[row.year][row.month] ?? 0) + row.qty
      if (row.year === latestYr) {
        p.monthlyLatest[row.month] = (p.monthlyLatest[row.month] ?? 0) + row.qty
        c.monthlyLatest[row.month] = (c.monthlyLatest[row.month] ?? 0) + row.qty
      }
    }
    return cats
  }, [ojcRows, latestYr])

  // 기타(비OJC) 행 필터링 및 카테고리별 집계
  const etcRows = useMemo(() =>
    rawRows.filter(r => classifyOjc(r.name) === null && classifyEtc(r.name) !== null),
    [rawRows]
  )

  const etcByCategory = useMemo<Record<string, CategoryData>>(() => {
    const cats: Record<string, CategoryData> = {}
    for (const row of etcRows) {
      const cat = classifyEtc(row.name)!
      if (!cats[cat]) cats[cat] = { products: {}, annuals: {}, priceAnnuals: {}, monthlyLatest: {}, monthlyByYear: {} }
      const c = cats[cat]
      if (!c.products[row.name]) c.products[row.name] = { code: row.code, annuals: {}, priceAnnuals: {}, monthlyLatest: {} }
      const p = c.products[row.name]
      p.annuals[row.year]       = (p.annuals[row.year] ?? 0) + row.qty
      p.priceAnnuals[row.year]  = (p.priceAnnuals[row.year] ?? 0) + row.price
      c.annuals[row.year]       = (c.annuals[row.year] ?? 0) + row.qty
      c.priceAnnuals[row.year]  = (c.priceAnnuals[row.year] ?? 0) + row.price
      if (!c.monthlyByYear[row.year]) c.monthlyByYear[row.year] = {}
      c.monthlyByYear[row.year][row.month] = (c.monthlyByYear[row.year][row.month] ?? 0) + row.qty
      if (row.year === latestYr) {
        p.monthlyLatest[row.month] = (p.monthlyLatest[row.month] ?? 0) + row.qty
        c.monthlyLatest[row.month] = (c.monthlyLatest[row.month] ?? 0) + row.qty
      }
    }
    return cats
  }, [etcRows, latestYr])

  // 거래처별 탑3 집계 (열수축 슬리브 제외)
  const customerTop3 = useMemo(() => {
    const custMap: Record<string, Record<string, { code: string; qty: number; price: number }>> = {}
    for (const row of rawRows) {
      if (CUSTOMER_TOP3_EXCLUDE.has(row.name)) continue
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
  }, [rawRows])

  // 전체 품목별 집계
  const fullProducts = useMemo(() => {
    const map: Record<string, { name: string; code: string; ojcCat: string | null; etcCat: string | null; annuals: Record<string, number>; priceAnnuals: Record<string, number> }> = {}
    for (const row of rawRows) {
      if (!map[row.name]) map[row.name] = { name: row.name, code: row.code, ojcCat: classifyOjc(row.name), etcCat: classifyEtc(row.name), annuals: {}, priceAnnuals: {} }
      map[row.name].annuals[row.year]      = (map[row.name].annuals[row.year] ?? 0) + row.qty
      map[row.name].priceAnnuals[row.year] = (map[row.name].priceAnnuals[row.year] ?? 0) + row.price
    }
    return Object.values(map).sort((a, b) =>
      (b.annuals[latestYr] ?? 0) - (a.annuals[latestYr] ?? 0)
    )
  }, [rawRows, latestYr])

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
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs leading-5 max-h-24 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {/* 서브탭 + 통합 다운로드 */}
      {hasData && (
        <div className="flex gap-2 flex-wrap items-center justify-between">
          <div className="flex gap-2 flex-wrap">
            {(['ojc', 'customer', 'full'] as SubView[]).map(v => (
              <button
                key={v}
                onClick={() => setSubView(v)}
                className={`px-4 py-1.5 text-sm font-medium rounded border transition ${
                  subView === v
                    ? 'bg-yellow-400 border-yellow-500 text-black'
                    : 'bg-white border-gray-300 text-gray-600 hover:bg-gray-50'
                }`}
              >
                {v === 'ojc' ? '① OJC 판매 현황' : v === 'customer' ? '② 거래처별 탑3' : '③ 전체 판매량'}
              </button>
            ))}
          </div>
          <button
            disabled={downloading}
            onClick={() => { setDownloading(true); downloadStyledExcel(ojcByCategory, etcByCategory, customerTop3, fullProducts, years, latestYr, ojcStock).finally(() => setDownloading(false)) }}
            className="px-4 py-1.5 text-sm bg-[#1F3864] hover:bg-[#162a4d] disabled:bg-gray-400 text-white font-semibold rounded transition whitespace-nowrap"
          >
            {downloading ? '⏳ 생성 중...' : '📥 전체 통합 다운로드 (6시트)'}
          </button>
        </div>
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
      {hasData && subView === 'customer' && (
        <CustomerTop3View customerTop3={customerTop3} years={years} />
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

// ── 통합 판매현황 다운로드 (ExcelJS 스타일) ───────────────────────────
async function downloadStyledExcel(
  ojcByCategory: Record<string, CategoryData>,
  etcByCategory: Record<string, CategoryData>,
  customerTop3: Record<string, Array<{ name: string; code: string; qty: number; price: number }>>,
  fullProducts: Array<{ name: string; code: string; ojcCat: string | null; etcCat: string | null; annuals: Record<string, number>; priceAnnuals: Record<string, number> }>,
  years: string[],
  latestYr: string,
  ojcStock: Record<string, number>,
) {
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

  function numCell(cell: ExcelJS.Cell, bg: string, right = true) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
    cell.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
    cell.alignment = { horizontal: right ? 'right' : 'center', vertical: 'middle' }
    if (typeof cell.value === 'number') cell.numFmt = '#,##0'
  }

  // ── Sheet 1: 판매현황_요약 (화면과 동일) ──────────────────────────────
  const ws1 = wb.addWorksheet('① 판매현황_요약')
  ws1.views = [{ state: 'frozen', ySplit: 3 }]
  const S1 = 2 + years.length + 4
  addTitle(ws1, `AJW 판매현황 분석  ·  ${years.map(y => '20' + y + '년').join(' / ')}`,
    `피크커버·평균커버 = 현재고 ÷ ${latestYr}년 수요 (개월)  |  생성: ${today()}`, S1)
  styleHdr(ws1.addRow(['카테고리', '품목수', ...years.map(y => `${y}년\n판매(EA)`),
    `${latestYr}년\n피크월(EA)`, `${latestYr}년\n월평균(EA)`, '현재고\n(EA)', '피크커버\n(개월)', '평균커버\n(개월)']), C.midBlue, 32)

  function addSumRow(ws: ExcelJS.Worksheet, cat: string, data: CategoryData, isEtc: boolean, rn: number) {
    const la = data.annuals[latestYr] ?? 0; const mv = MONTHS.map(m => data.monthlyLatest[m] ?? 0)
    const pk = Math.max(0, ...mv); const avg = la / 12; const st = ojcStock[cat] ?? 0
    const cpk = st > 0 && pk > 0 ? parseFloat((st / pk).toFixed(1)) : null
    const cav = st > 0 && avg > 0 ? parseFloat((st / avg).toFixed(1)) : null
    const bg = rn % 2 === 0 ? C.white : (isEtc ? C.altPurple : C.altBlue)
    const row = ws.addRow([cat, Object.keys(data.products).length,
      ...years.map(yr => data.annuals[yr] || null), pk || null, la > 0 ? Math.round(avg) : null, st || null, cpk, cav])
    row.height = 18
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ci === 1) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { bold: true, color: { argb: isEtc ? C.purple : C.midBlue } } }
      else if (ci === 2) { c.alignment = { horizontal: 'center', vertical: 'middle' } }
      else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
    })
    const pkC = row.getCell(2 + years.length + 1)
    if (pk > 0) pkC.font = { bold: true, color: { argb: C.orange } }
    if (cpk !== null) row.getCell(S1 - 1).font = { bold: true, color: { argb: coverFg(cpk) } }
    if (cav !== null) row.getCell(S1).font = { bold: true, color: { argb: coverFg(cav) } }
  }

  let ri = 0
  for (const [c, d] of Object.entries(ojcByCategory)) addSumRow(ws1, c, d, false, ri++)
  addSep(ws1, '  기타 품목 판매 현황', S1); ri = 0
  for (const [c, d] of Object.entries(etcByCategory)) addSumRow(ws1, c, d, true, ri++)
  ws1.columns = [{ width: 28 }, { width: 8 }, ...years.map(() => ({ width: 14 })), { width: 13 }, { width: 13 }, { width: 12 }, { width: 11 }, { width: 11 }]

  // ── Sheet 2: 연도별_피크분석 ──────────────────────────────────────────
  const ws2 = wb.addWorksheet('② 연도별_피크분석')
  ws2.views = [{ state: 'frozen', ySplit: 3 }]
  const S2 = 2 + years.length * 3 + 3
  addTitle(ws2, '연도별 피크월 분석 (OJC + 기타)', `각 연도의 최대 판매 월·수량  |  커버리지는 ${latestYr}년 기준`, S2)
  styleHdr(ws2.addRow(['카테고리', '품목수',
    ...years.flatMap(y => [`${y}년\n판매(EA)`, `${y}년\n피크월`, `${y}년\n피크(EA)`]),
    '현재고\n(EA)', '피크커버\n(개월)', '평균커버\n(개월)']), C.darkBlue, 32)

  function addPeakRow(ws: ExcelJS.Worksheet, cat: string, data: CategoryData, isEtc: boolean, rn: number) {
    const la = data.annuals[latestYr] ?? 0; const mv = MONTHS.map(m => data.monthlyLatest[m] ?? 0)
    const pk = Math.max(0, ...mv); const avg = la / 12; const st = ojcStock[cat] ?? 0
    const cpk = st > 0 && pk > 0 ? parseFloat((st / pk).toFixed(1)) : null
    const cav = st > 0 && avg > 0 ? parseFloat((st / avg).toFixed(1)) : null
    const bg = rn % 2 === 0 ? C.white : (isEtc ? C.altPurple : C.altBlue)
    const yearCols = years.flatMap(yr => { const p = peakOfYear(data, yr); return [data.annuals[yr] || null, p.month || null, p.val || null] })
    const row = ws.addRow([cat, Object.keys(data.products).length, ...yearCols, st || null, cpk, cav])
    row.height = 18
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ci === 1) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { bold: true, color: { argb: isEtc ? C.purple : C.midBlue } } }
      else if (ci === 2) { c.alignment = { horizontal: 'center', vertical: 'middle' } }
      else {
        const rel = (ci - 3) % 3
        if (rel === 1) { c.alignment = { horizontal: 'center', vertical: 'middle' }; if (c.value) c.font = { bold: true, color: { argb: C.orange } } }
        else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
      }
    })
    if (cpk !== null) row.getCell(S2 - 1).font = { bold: true, color: { argb: coverFg(cpk) } }
    if (cav !== null) row.getCell(S2).font = { bold: true, color: { argb: coverFg(cav) } }
  }

  ri = 0
  for (const [c, d] of Object.entries(ojcByCategory)) addPeakRow(ws2, c, d, false, ri++)
  addSep(ws2, '  기타 품목', S2); ri = 0
  for (const [c, d] of Object.entries(etcByCategory)) addPeakRow(ws2, c, d, true, ri++)
  ws2.columns = [{ width: 28 }, { width: 7 }, ...years.flatMap(() => [{ width: 13 }, { width: 9 }, { width: 12 }]), { width: 11 }, { width: 11 }, { width: 11 }]

  // ── Sheet 3: 품목별_상세 (OJC + 기타) ────────────────────────────────
  const ws3 = wb.addWorksheet('③ 품목별_상세')
  ws3.views = [{ state: 'frozen', ySplit: 3 }]
  const S3 = 3 + years.length * 2 + 2
  addTitle(ws3, '품목별 판매 상세 (OJC + 기타)', `${latestYr}년 판매량 기준 내림차순  |  카테고리별 묶음`, S3)
  styleHdr(ws3.addRow(['카테고리', '품목코드', '품목명', ...years.flatMap(y => [`${y}년\n판매(EA)`, `${y}년\n공급가액`]), '피크월\n(EA)', '월평균\n(EA)']), C.midBlue, 28)

  ri = 0
  function addProdRows(ws: ExcelJS.Worksheet, cat: string, data: CategoryData, isEtc: boolean) {
    for (const [name, prod] of Object.entries(data.products)
        .sort((a, b) => (b[1].annuals[latestYr] ?? 0) - (a[1].annuals[latestYr] ?? 0))) {
      const pla = prod.annuals[latestYr] ?? 0
      const ppk = Math.max(0, ...MONTHS.map(m => prod.monthlyLatest[m] ?? 0))
      const bg = ri++ % 2 === 0 ? C.white : (isEtc ? C.altPurple : C.altBlue)
      const row = ws.addRow([cat, prod.code || null, name, ...years.flatMap(yr => [prod.annuals[yr] || null, prod.priceAnnuals[yr] || null]), ppk || null, pla > 0 ? Math.round(pla / 12) : null])
      row.height = 16
      row.eachCell({ includeEmpty: true }, (c, ci) => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
        c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
        if (ci <= 3) { c.alignment = { horizontal: 'left', vertical: 'middle' }; if (ci === 1) c.font = { color: { argb: isEtc ? C.purple : C.midBlue } } }
        else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
      })
    }
  }

  for (const [c, d] of Object.entries(ojcByCategory)) addProdRows(ws3, c, d, false)
  addSep(ws3, '  기타 품목', S3); ri = 0
  for (const [c, d] of Object.entries(etcByCategory)) addProdRows(ws3, c, d, true)
  ws3.columns = [{ width: 22 }, { width: 16 }, { width: 40 }, ...years.flatMap(() => [{ width: 13 }, { width: 16 }]), { width: 11 }, { width: 11 }]

  // ── Sheet 4: 월간현황 ─────────────────────────────────────────────────
  const ws4 = wb.addWorksheet(`④ 월간현황_${latestYr}년`)
  ws4.views = [{ state: 'frozen', ySplit: 3 }]
  const S4 = 2 + 12
  addTitle(ws4, `${latestYr}년 월간 판매 현황 (OJC + 기타)`, '카테고리 합계 → 품목별 상세  |  주황 = 피크월', S4)
  styleHdr(ws4.addRow(['카테고리', '품목명', ...MONTHS.map(m => `${parseInt(m)}월`)]), C.midBlue, 24)

  function addMonthRows(ws: ExcelJS.Worksheet, cat: string, data: CategoryData, isEtc: boolean) {
    const mv = MONTHS.map(m => data.monthlyLatest[m] ?? 0); const pk = Math.max(0, ...mv)
    const catRow = ws.addRow([cat, '(합계)', ...MONTHS.map(m => data.monthlyLatest[m] || null)])
    catRow.height = 20
    catRow.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEtc ? C.lightPurple : C.lightBlue } }
      c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      c.font = { bold: true, color: { argb: isEtc ? C.purple : C.midBlue } }
      c.alignment = { horizontal: ci > 2 ? 'right' : 'left', vertical: 'middle' }
      if (ci > 2 && typeof c.value === 'number') c.numFmt = '#,##0'
    })
    MONTHS.forEach((m, mi) => { if ((data.monthlyLatest[m] ?? 0) === pk && pk > 0) { const pc = catRow.getCell(3 + mi); pc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } }; pc.font = { bold: true, color: { argb: C.orange } } } })
    for (const [name, prod] of Object.entries(data.products).sort((a, b) => (b[1].annuals[latestYr] ?? 0) - (a[1].annuals[latestYr] ?? 0))) {
      const row = ws.addRow([cat, name, ...MONTHS.map(m => prod.monthlyLatest[m] || null)]); row.height = 16
      row.eachCell({ includeEmpty: true }, (c, ci) => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.white } }
        c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
        c.alignment = { horizontal: ci > 2 ? 'right' : 'left', vertical: 'middle' }
        if (ci === 1) c.font = { color: { argb: isEtc ? C.purple : C.midBlue } }
        if (ci > 2 && typeof c.value === 'number') c.numFmt = '#,##0'
      })
    }
  }

  for (const [c, d] of Object.entries(ojcByCategory)) addMonthRows(ws4, c, d, false)
  addSep(ws4, '  기타 품목 판매 현황', S4); for (const [c, d] of Object.entries(etcByCategory)) addMonthRows(ws4, c, d, true)
  ws4.columns = [{ width: 22 }, { width: 38 }, ...MONTHS.map(() => ({ width: 8 }))]

  // ── Sheet 5: 거래처별탑3 ──────────────────────────────────────────────
  const ws5 = wb.addWorksheet('⑤ 거래처별탑3')
  ws5.views = [{ state: 'frozen', ySplit: 3 }]
  addTitle(ws5, '거래처별 TOP3 구매 품목', `전체 기간(${years.map(y => '20' + y + '년').join('~')}) 누적  |  총구매량 내림차순`, 15)
  styleHdr(ws5.addRow(['순위', '거래처명', '총구매(EA)', 'TOP1 품목코드', 'TOP1 품목명', 'TOP1\n수량(EA)', 'TOP1\n공급가액', 'TOP2 품목코드', 'TOP2 품목명', 'TOP2\n수량(EA)', 'TOP2\n공급가액', 'TOP3 품목코드', 'TOP3 품목명', 'TOP3\n수량(EA)', 'TOP3\n공급가액']), C.midBlue, 26)
  Object.entries(customerTop3).forEach(([cust, top3], i) => {
    const total = top3.reduce((s, x) => s + x.qty, 0)
    const bg = i % 2 === 0 ? C.white : C.altBlue
    const row = ws5.addRow([
      i + 1, cust, total,
      top3[0]?.code ?? null, top3[0]?.name ?? null, top3[0]?.qty ?? null, top3[0]?.price || null,
      top3[1]?.code ?? null, top3[1]?.name ?? null, top3[1]?.qty ?? null, top3[1]?.price || null,
      top3[2]?.code ?? null, top3[2]?.name ?? null, top3[2]?.qty ?? null, top3[2]?.price || null,
    ])
    row.height = 17
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }; c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ci === 1) { c.alignment = { horizontal: 'center', vertical: 'middle' }; c.font = { color: { argb: C.gray3 } } }
      else if (ci === 2) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { bold: true } }
      else if (ci === 3) { c.alignment = { horizontal: 'right', vertical: 'middle' }; c.numFmt = '#,##0'; c.font = { bold: true } }
      else {
        const rel = (ci - 4) % 4
        if (rel === 0) { c.alignment = { horizontal: 'left', vertical: 'middle' }; c.font = { size: 9, color: { argb: C.gray3 } } }
        else if (rel === 1) { c.alignment = { horizontal: 'left', vertical: 'middle' } }
        else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
      }
    })
    if (i === 0) row.getCell(2).font = { bold: true, color: { argb: 'FFCC0000' } }
    else if (i === 1) row.getCell(2).font = { bold: true, color: { argb: 'FF833C00' } }
  })
  ws5.columns = [{ width: 6 }, { width: 24 }, { width: 13 }, { width: 16 }, { width: 36 }, { width: 11 }, { width: 14 }, { width: 16 }, { width: 36 }, { width: 11 }, { width: 14 }, { width: 16 }, { width: 36 }, { width: 11 }, { width: 14 }]

  // ── Sheet 6: 전체판매량 ───────────────────────────────────────────────
  const ws6 = wb.addWorksheet('⑥ 전체판매량')
  ws6.views = [{ state: 'frozen', ySplit: 3 }]
  const S6 = 4 + years.length * 2
  addTitle(ws6, '전체 품목 판매량', `${latestYr}년 판매량 기준 내림차순`, S6)
  styleHdr(ws6.addRow(['품목코드', '품목명', '분류', ...years.flatMap(y => [`${y}년\n판매(EA)`, `${y}년\n공급가액`]), '합계\n(EA)']), C.darkBlue, 28)
  fullProducts.forEach((p, i) => {
    const total = years.reduce((s, yr) => s + (p.annuals[yr] ?? 0), 0)
    const bg = i % 2 === 0 ? C.white : C.gray1
    const row = ws6.addRow([p.code || null, p.name, p.ojcCat ?? p.etcCat ?? '기타', ...years.flatMap(yr => [p.annuals[yr] || null, p.priceAnnuals[yr] || null]), total || null])
    row.height = 16
    row.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }; c.border = { bottom: { style: 'thin', color: { argb: C.gray2 } } }
      if (ci <= 3) { c.alignment = { horizontal: 'left', vertical: 'middle' }; if (ci === 3 && p.ojcCat) c.font = { color: { argb: C.midBlue } } }
      else { c.alignment = { horizontal: 'right', vertical: 'middle' }; if (typeof c.value === 'number') c.numFmt = '#,##0' }
    })
  })
  if (fullProducts.length > 0) {
    const tots = years.map(yr => fullProducts.reduce((s, p) => s + (p.annuals[yr] ?? 0), 0))
    const sr = ws6.addRow(['합계', `${fullProducts.length}종`, '', ...tots.flatMap(t => [t, null]), tots.reduce((a, b) => a + b, 0)]); sr.height = 22
    sr.eachCell({ includeEmpty: true }, (c, ci) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.darkBlue } }; c.font = { bold: true, color: { argb: C.white } }
      c.border = { top: { style: 'medium', color: { argb: C.midBlue } } }
      if (ci > 3) { c.alignment = { horizontal: 'right', vertical: 'middle' }; c.numFmt = '#,##0' } else c.alignment = { horizontal: 'left', vertical: 'middle' }
    })
  }
  ws6.columns = [{ width: 16 }, { width: 40 }, { width: 18 }, ...years.flatMap(() => [{ width: 13 }, { width: 16 }]), { width: 13 }]

  const buffer = await wb.xlsx.writeBuffer()
  downloadXlsx(buffer as ArrayBuffer, `판매현황분석_${today()}.xlsx`)
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
  const [expanded, setExpanded] = useState<string | null>(null)

  const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'

  // 재고는 있지만 판매 데이터가 전혀 없는 카테고리 (OJC/ETC 모두 아닌 것)
  const ojcCats      = new Set(Object.keys(ojcByCategory))
  const etcCats      = new Set(Object.keys(etcByCategory))
  const stockOnlyCats = Object.keys(ojcStock).filter(k => !ojcCats.has(k) && !etcCats.has(k) && ojcStock[k] > 0)

  return (
    <div className="space-y-3">
      <div className="overflow-x-auto border border-gray-200 rounded-lg">
        <table className="text-sm w-full border-collapse bg-white">
          <thead>
            <tr>
              <th className={`${thBase} text-left`}>카테고리</th>
              <th className={`${thBase} text-right`}>품목 수</th>
              {years.map(yr => (
                <th key={yr} className={`${thBase} text-right`}>{yr}년 판매(EA)</th>
              ))}
              {years.map(yr => (
                <th key={`p${yr}`} className={`${thBase} text-right`}>{yr}년 공급가액</th>
              ))}
              <th className={`${thBase} text-right`}>{latestYr}년<br/>최다 판매월</th>
              <th className={`${thBase} text-right`}>월평균(EA)</th>
              <th className={`${thBase} text-right`}>현재고(EA)</th>
              <th className={`${thBase} text-right`}>최다월<br/>재고커버</th>
              <th className={`${thBase} text-right`}>평균<br/>재고커버</th>
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

              return (
                <Fragment key={cat}>
                  <tr className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="px-3 py-2 font-semibold text-gray-800">
                      <button
                        onClick={() => setExpanded(isExpanded ? null : cat)}
                        className="text-left hover:text-blue-600 transition flex items-center gap-1"
                      >
                        <span className="text-xs">{isExpanded ? '▼' : '▶'}</span> {cat}
                      </button>
                    </td>
                    <td className="px-3 py-2 text-right text-gray-500">{prodEntries.length}</td>
                    {years.map(yr => (
                      <td key={yr} className="px-3 py-2 text-right font-mono">{num(data.annuals[yr] ?? 0)}</td>
                    ))}
                    {years.map(yr => (
                      <td key={`p${yr}`} className="px-3 py-2 text-right text-xs text-gray-600">{fmtPrice(data.priceAnnuals[yr] ?? 0)}</td>
                    ))}
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
                          <div className="text-xs font-semibold text-blue-700 mb-1.5">{latestYr}년 월간 판매 현황</div>
                          <div className="flex gap-1 flex-wrap">
                            {MONTHS.map(m => {
                              const v = data.monthlyLatest[m] ?? 0
                              const isPeak = v === peakMonthly && v > 0
                              return (
                                <div key={m} className={`text-center rounded px-2 py-1 min-w-[52px] ${isPeak ? 'bg-orange-200 font-bold text-orange-800' : 'bg-white border border-gray-200 text-gray-700'}`}>
                                  <div className="text-[10px] text-gray-400">{parseInt(m)}월</div>
                                  <div className="text-xs">{v > 0 ? v.toLocaleString() : '—'}</div>
                                </div>
                              )
                            })}
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
                              <td className="pl-7 pr-3 py-1.5 text-xs text-gray-600 max-w-xs" title={name}>
                                <div className="text-[10px] text-blue-400 font-mono">{prod.code || '—'}</div>
                                <div className="truncate">↳ {name}</div>
                              </td>
                              <td className="px-3 py-1.5 text-right text-xs text-gray-400">—</td>
                              {years.map(yr => (
                                <td key={yr} className="px-3 py-1.5 text-right text-xs font-mono text-gray-600">
                                  {num(prod.annuals[yr] ?? 0)}
                                </td>
                              ))}
                              {years.map(yr => (
                                <td key={`p${yr}`} className="px-3 py-1.5 text-right text-xs text-gray-500">
                                  {fmtPrice(prod.priceAnnuals[yr] ?? 0)}
                                </td>
                              ))}
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
                  return (
                    <Fragment key={`etc-${cat}`}>
                      <tr className={i % 2 === 0 ? 'bg-white' : 'bg-purple-50/20'}>
                        <td className="px-3 py-2 font-semibold text-purple-800">
                          <button
                            onClick={() => setExpanded(isExpanded ? null : `etc-${cat}`)}
                            className="text-left hover:text-purple-600 transition flex items-center gap-1"
                          >
                            <span className="text-xs">{isExpanded ? '▼' : '▶'}</span> {cat}
                          </button>
                        </td>
                        <td className="px-3 py-2 text-right text-gray-500">{prodEntries.length}</td>
                        {years.map(yr => (
                          <td key={yr} className="px-3 py-2 text-right font-mono">{num(data.annuals[yr] ?? 0)}</td>
                        ))}
                        {years.map(yr => (
                          <td key={`p${yr}`} className="px-3 py-2 text-right text-xs text-gray-600">{fmtPrice(data.priceAnnuals[yr] ?? 0)}</td>
                        ))}
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
                              <div className="text-xs font-semibold text-purple-700 mb-1.5">{latestYr}년 월간 판매 현황</div>
                              <div className="flex gap-1 flex-wrap">
                                {MONTHS.map(m => {
                                  const v = data.monthlyLatest[m] ?? 0
                                  const isPeak = v === peakMonthly && v > 0
                                  return (
                                    <div key={m} className={`text-center rounded px-2 py-1 min-w-[52px] ${isPeak ? 'bg-orange-200 font-bold text-orange-800' : 'bg-white border border-gray-200 text-gray-700'}`}>
                                      <div className="text-[10px] text-gray-400">{parseInt(m)}월</div>
                                      <div className="text-xs">{v > 0 ? v.toLocaleString() : '—'}</div>
                                    </div>
                                  )
                                })}
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
                                  <td className="pl-7 pr-3 py-1.5 text-xs text-purple-700 max-w-xs" title={name}>
                                    <div className="text-[10px] text-purple-400 font-mono">{prod.code || '—'}</div>
                                    <div className="truncate">↳ {name}</div>
                                  </td>
                                  <td className="px-3 py-1.5 text-right text-xs text-gray-400">—</td>
                                  {years.map(yr => (
                                    <td key={yr} className="px-3 py-1.5 text-right text-xs font-mono text-gray-600">
                                      {num(prod.annuals[yr] ?? 0)}
                                    </td>
                                  ))}
                                  {years.map(yr => (
                                    <td key={`p${yr}`} className="px-3 py-1.5 text-right text-xs text-gray-500">
                                      {fmtPrice(prod.priceAnnuals[yr] ?? 0)}
                                    </td>
                                  ))}
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
                    <tr key={cat} className="bg-gray-50/40 hover:bg-gray-100">
                      <td className="px-3 py-2 font-semibold text-gray-500 text-sm">{cat}</td>
                      <td className="px-3 py-2 text-right text-gray-300 text-xs">—</td>
                      {years.map(yr => (
                        <td key={yr} className="px-3 py-2 text-right text-gray-300 text-xs">—</td>
                      ))}
                      {years.map(yr => (
                        <td key={`p${yr}`} className="px-3 py-2 text-right text-gray-300 text-xs">—</td>
                      ))}
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
        <div>※ 피크월 / 월평균은 <b>{latestYr}년</b> 기준 | 피크커버·평균커버 = 현재고 ÷ 해당 월 수요 (개월)</div>
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

function CustomerTop3View({
  customerTop3, years,
}: {
  customerTop3: Record<string, Array<{ name: string; code: string; qty: number; price: number }>>
  years:        string[]
}) {
  const [search, setSearch] = useState('')
  const customers = Object.entries(customerTop3)
  const filtered  = customers.filter(([cust]) =>
    !search || cust.toLowerCase().includes(search.toLowerCase())
  )

  const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'

  if (!customers.length) {
    return (
      <div className="text-center text-gray-400 py-10 text-sm">
        거래처 데이터가 없습니다. 파일에 <code className="bg-gray-100 px-1 rounded">거래처명</code> 컬럼이 있어야 합니다.
      </div>
    )
  }

  return (
    <div className="space-y-3">
      <div className="flex justify-between items-center gap-2 flex-wrap">
        <div className="text-sm text-gray-600 font-medium">
          전체 {customers.length}개 거래처 | 전체 기간({years.map(y => '20' + y + '년').join('~')}) 기준
        </div>
        <div className="flex gap-2 items-center">
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
              <th className={`${thBase} text-right`}>총구매(EA)</th>
              {[1, 2, 3].map(rank => (
                <Fragment key={rank}>
                  <th className={`${thBase} text-left`}>TOP{rank} 품목명</th>
                  <th className={`${thBase} text-right`}>수량(EA)</th>
                  <th className={`${thBase} text-right`}>공급가액</th>
                </Fragment>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {filtered.map(([cust, top3], i) => {
              const total = top3.reduce((s, x) => s + x.qty, 0)
              return (
                <tr key={cust} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                  <td className="px-3 py-2 text-right text-gray-400 text-xs">{i + 1}</td>
                  <td className="px-3 py-2 font-semibold text-gray-800">{cust}</td>
                  <td className="px-3 py-2 text-right font-mono font-bold text-gray-900">
                    {total.toLocaleString()}
                  </td>
                  {[0, 1, 2].map(rank => (
                    <Fragment key={rank}>
                      <td className="px-3 py-2 max-w-[200px] text-sm text-gray-700" title={top3[rank]?.name ?? ''}>
                        <div className="text-[10px] text-gray-400 font-mono">{top3[rank]?.code || '—'}</div>
                        <div className="truncate">{top3[rank]?.name ?? <span className="text-gray-300">—</span>}</div>
                      </td>
                      <td className="px-3 py-2 text-right font-mono text-sm">
                        {top3[rank] ? top3[rank].qty.toLocaleString() : <span className="text-gray-300">—</span>}
                      </td>
                      <td className="px-3 py-2 text-right text-xs text-gray-600">
                        {top3[rank] ? fmtPrice(top3[rank].price) : <span className="text-gray-300">—</span>}
                      </td>
                    </Fragment>
                  ))}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
      <div className="text-xs text-gray-400">
        ※ 전체 기간 누적 수량 기준 | 거래처당 최다 구매 품목 탑3
      </div>
    </div>
  )
}

// ── ③ 전체 판매량 ──────────────────────────────────────────────────
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

function FullSalesView({
  products, years, latestYr,
}: {
  products:  Array<{ name: string; code: string; ojcCat: string | null; etcCat: string | null; annuals: Record<string, number>; priceAnnuals: Record<string, number> }>
  years:     string[]
  latestYr:  string
}) {
  const [search, setSearch]       = useState('')
  const [filterOjc, setFilterOjc] = useState<'all' | 'ojc' | 'non'>('all')

  const ojcCount    = products.filter(p => p.ojcCat !== null).length
  const nonOjcCount = products.length - ojcCount

  const filtered = products.filter(p => {
    const matchOjc =
      filterOjc === 'all' ||
      (filterOjc === 'ojc' && p.ojcCat !== null) ||
      (filterOjc === 'non' && p.ojcCat === null)
    const matchSearch = !search || p.name.toLowerCase().includes(search.toLowerCase())
    return matchOjc && matchSearch
  })

  // 연도별 합계
  const yearTotals = years.reduce<Record<string, number>>((acc, yr) => {
    acc[yr] = filtered.reduce((s, p) => s + (p.annuals[yr] ?? 0), 0)
    return acc
  }, {})

  const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'

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
              <th className={`${thBase} text-left`}>품목명</th>
              <th className={`${thBase} text-center`}>분류</th>
              {years.map(yr => (
                <Fragment key={yr}>
                  <th className={`${thBase} text-right`}>{yr}년 판매(EA)</th>
                  <th className={`${thBase} text-right`}>{yr}년 공급가액</th>
                </Fragment>
              ))}
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
                {years.map(yr => (
                  <Fragment key={yr}>
                    <td className="px-3 py-1.5 text-right font-mono text-sm">{num(p.annuals[yr] ?? 0)}</td>
                    <td className="px-3 py-1.5 text-right text-xs text-gray-500">{fmtPrice(p.priceAnnuals[yr] ?? 0)}</td>
                  </Fragment>
                ))}
              </tr>
            ))}

            {/* 합계 행 */}
            {filtered.length > 0 && (
              <tr className="bg-gray-100 font-bold border-t-2 border-gray-300">
                <td className="px-3 py-2 text-gray-800">합계</td>
                <td className="px-3 py-2 text-center text-gray-500 text-xs">{filtered.length}종</td>
                {years.map(yr => (
                  <Fragment key={yr}>
                    <td className="px-3 py-2 text-right font-mono">{num(yearTotals[yr] ?? 0)}</td>
                    <td className="px-3 py-2 text-right text-xs text-gray-400">—</td>
                  </Fragment>
                ))}
              </tr>
            )}

            {filtered.length === 0 && (
              <tr><td colSpan={2 + years.length * 2} className="text-center text-gray-400 py-10">검색 결과 없음</td></tr>
            )}
          </tbody>
        </table>
      </div>
      <div className="text-xs text-gray-400 flex justify-between">
        <span>※ {latestYr}년 판매량 기준 내림차순 정렬</span>
        <span>{filtered.length.toLocaleString()}개 품목</span>
      </div>
    </div>
  )
}
