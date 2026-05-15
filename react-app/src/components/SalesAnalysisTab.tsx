import { useState, useMemo, Fragment } from 'react'
import { classifyOjc } from '../lib/ojcFilter'
import { parseDetailedSalesFile } from '../lib/parse/parseDetailedSales'
import type { DetailedSalesRow } from '../lib/parse/parseDetailedSales'
import FileUploader from './FileUploader'

// ── 상수 ────────────────────────────────────────────────────────────
const MONTHS = ['01','02','03','04','05','06','07','08','09','10','11','12']

const num   = (v: number) => v === 0 ? '—' : v.toLocaleString()
const dec1  = (v: number) => v.toFixed(1)

type SubView = 'ojc' | 'customer' | 'full'

// ── 타입 ─────────────────────────────────────────────────────────────
interface ProductData {
  annuals:       Record<string, number>
  monthlyLatest: Record<string, number>
}
interface CategoryData {
  products:      Record<string, ProductData>
  annuals:       Record<string, number>
  monthlyLatest: Record<string, number>
}

// ── 메인 컴포넌트 ─────────────────────────────────────────────────────
export default function SalesAnalysisTab() {
  const [rawRows, setRawRows]   = useState<DetailedSalesRow[]>([])
  const [logs, setLogs]         = useState<string[]>([])
  const [file, setFile]         = useState<File | null>(null)
  const [running, setRunning]   = useState(false)
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
      if (!cats[cat]) cats[cat] = { products: {}, annuals: {}, monthlyLatest: {} }
      const c = cats[cat]
      if (!c.products[row.name]) c.products[row.name] = { annuals: {}, monthlyLatest: {} }
      const p = c.products[row.name]
      p.annuals[row.year]       = (p.annuals[row.year] ?? 0) + row.qty
      c.annuals[row.year]       = (c.annuals[row.year] ?? 0) + row.qty
      if (row.year === latestYr) {
        p.monthlyLatest[row.month] = (p.monthlyLatest[row.month] ?? 0) + row.qty
        c.monthlyLatest[row.month] = (c.monthlyLatest[row.month] ?? 0) + row.qty
      }
    }
    return cats
  }, [ojcRows, latestYr])

  // 거래처별 탑3 집계
  const customerTop3 = useMemo(() => {
    const custMap: Record<string, Record<string, number>> = {}
    for (const row of rawRows) {
      const cust = row.customer || '(미상)'
      if (!custMap[cust]) custMap[cust] = {}
      custMap[cust][row.name] = (custMap[cust][row.name] ?? 0) + row.qty
    }
    const result: Record<string, Array<{ name: string; qty: number }>> = {}
    for (const [cust, products] of Object.entries(custMap)) {
      result[cust] = Object.entries(products)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3)
        .map(([name, qty]) => ({ name, qty }))
    }
    return Object.fromEntries(
      Object.entries(result).sort((a, b) =>
        b[1].reduce((s, x) => s + x.qty, 0) - a[1].reduce((s, x) => s + x.qty, 0)
      )
    )
  }, [rawRows])

  // 전체 품목별 집계
  const fullProducts = useMemo(() => {
    const map: Record<string, { name: string; ojcCat: string | null; annuals: Record<string, number> }> = {}
    for (const row of rawRows) {
      if (!map[row.name]) map[row.name] = { name: row.name, ojcCat: classifyOjc(row.name), annuals: {} }
      map[row.name].annuals[row.year] = (map[row.name].annuals[row.year] ?? 0) + row.qty
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
      <div className="flex gap-3 items-end">
        <div className="flex-1">
          <FileUploader
            label="대외비_판매현황.xlsx — 컬럼: 거래처명, 품목코드, 품목명, 년, 월, 수량"
            fileName={file?.name ?? ''}
            onFile={setFile}
          />
        </div>
        <button
          onClick={run}
          disabled={!file || running}
          className="bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold px-5 py-2 rounded transition text-sm h-[42px] whitespace-nowrap"
        >
          {running ? '⏳ 분석 중...' : '▶ 분석 실행'}
        </button>
      </div>

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs leading-5 max-h-24 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {/* 서브탭 */}
      {hasData && (
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
      )}

      {hasData && subView === 'ojc' && (
        <OjcSalesView
          ojcByCategory={ojcByCategory}
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

// ── ① OJC 판매 현황 ───────────────────────────────────────────────────
function OjcSalesView({
  ojcByCategory, years, latestYr, ojcStock, onStockChange,
}: {
  ojcByCategory:  Record<string, CategoryData>
  years:          string[]
  latestYr:       string
  ojcStock:       Record<string, number>
  onStockChange:  (cat: string, val: string) => void
}) {
  const [expanded, setExpanded] = useState<string | null>(null)

  const thBase = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-300 bg-gray-100 whitespace-nowrap'

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
              <th className={`${thBase} text-right`}>피크월(EA)</th>
              <th className={`${thBase} text-right`}>월평균(EA)</th>
              <th className={`${thBase} text-right`}>현재고(EA)</th>
              <th className={`${thBase} text-right`}>피크커버</th>
              <th className={`${thBase} text-right`}>평균커버</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {Object.entries(ojcByCategory).map(([cat, data], i) => {
              const latestAnnual  = data.annuals[latestYr] ?? 0
              const monthlyVals   = MONTHS.map(m => data.monthlyLatest[m] ?? 0)
              const peakMonthly   = Math.max(0, ...monthlyVals)
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
                    <td className="px-3 py-2 text-right font-mono font-semibold text-orange-700">{num(peakMonthly)}</td>
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
                        <td colSpan={3 + years.length + 5} className="px-4 py-2">
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
                              <td className="pl-7 pr-3 py-1.5 text-xs text-gray-600 max-w-xs truncate" title={name}>
                                ↳ {name}
                              </td>
                              <td className="px-3 py-1.5 text-right text-xs text-gray-400">—</td>
                              {years.map(yr => (
                                <td key={yr} className="px-3 py-1.5 text-right text-xs font-mono text-gray-600">
                                  {num(prod.annuals[yr] ?? 0)}
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
                <td colSpan={5 + years.length + 3} className="text-center text-gray-400 py-10">
                  OJC 품목이 감지되지 않았습니다. (품목명 앞부분: OJC-, SOJC-, DOJC-, MOJC-, DROP-CABLE, PIGTAIL- 등)
                </td>
              </tr>
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
function CustomerTop3View({
  customerTop3, years,
}: {
  customerTop3: Record<string, Array<{ name: string; qty: number }>>
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
      <div className="flex justify-between items-center">
        <div className="text-sm text-gray-600 font-medium">
          전체 {customers.length}개 거래처 | 전체 기간({years.map(y => '20' + y + '년').join('~')}) 기준
        </div>
        <input
          type="text"
          placeholder="거래처 검색..."
          value={search}
          onChange={e => setSearch(e.target.value)}
          className="border border-gray-300 rounded px-3 py-1 text-sm focus:outline-none focus:border-blue-400 w-48"
        />
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
                      <td
                        className="px-3 py-2 max-w-[200px] truncate text-sm text-gray-700"
                        title={top3[rank]?.name ?? ''}
                      >
                        {top3[rank]?.name ?? <span className="text-gray-300">—</span>}
                      </td>
                      <td className="px-3 py-2 text-right font-mono text-sm">
                        {top3[rank] ? top3[rank].qty.toLocaleString() : <span className="text-gray-300">—</span>}
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
function FullSalesView({
  products, years, latestYr,
}: {
  products:  Array<{ name: string; ojcCat: string | null; annuals: Record<string, number> }>
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
        <input
          type="text"
          placeholder="품목명 검색..."
          value={search}
          onChange={e => setSearch(e.target.value)}
          className="ml-auto border border-gray-300 rounded px-3 py-1 text-sm focus:outline-none focus:border-blue-400 w-52"
        />
      </div>

      <div className="overflow-x-auto border border-gray-200 rounded-lg">
        <table className="text-xs w-full border-collapse bg-white">
          <thead>
            <tr>
              <th className={`${thBase} text-left`}>품목명</th>
              <th className={`${thBase} text-center`}>분류</th>
              {years.map(yr => (
                <th key={yr} className={`${thBase} text-right`}>{yr}년 판매(EA)</th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {filtered.map((p, i) => (
              <tr key={p.name} className={i % 2 === 0 ? 'bg-white hover:bg-gray-50' : 'bg-gray-50 hover:bg-gray-100'}>
                <td className="px-3 py-1.5 max-w-sm truncate text-gray-800 text-sm" title={p.name}>{p.name}</td>
                <td className="px-3 py-1.5 text-center">
                  <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${
                    p.ojcCat ? 'bg-blue-100 text-blue-700' : 'bg-gray-100 text-gray-500'
                  }`}>
                    {p.ojcCat ?? '비OJC'}
                  </span>
                </td>
                {years.map(yr => (
                  <td key={yr} className="px-3 py-1.5 text-right font-mono text-sm">{num(p.annuals[yr] ?? 0)}</td>
                ))}
              </tr>
            ))}

            {/* 합계 행 */}
            {filtered.length > 0 && (
              <tr className="bg-gray-100 font-bold border-t-2 border-gray-300">
                <td className="px-3 py-2 text-gray-800">합계</td>
                <td className="px-3 py-2 text-center text-gray-500 text-xs">{filtered.length}종</td>
                {years.map(yr => (
                  <td key={yr} className="px-3 py-2 text-right font-mono">{num(yearTotals[yr] ?? 0)}</td>
                ))}
              </tr>
            )}

            {filtered.length === 0 && (
              <tr><td colSpan={2 + years.length} className="text-center text-gray-400 py-10">검색 결과 없음</td></tr>
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
