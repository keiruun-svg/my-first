import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import { saveMetadata, saveInventory, saveMaterialStock, CAN_WRITE } from '../lib/supabase'
import { parseMaterialStock } from '../lib/parse/parseMaterialStock'
import { today } from '../lib/download'
import type { Metadata, Inventory, CableMeta, HousingComp, FerruleMeta, InventoryHousingItem } from '../lib/types'
interface Props {
  metadata:     Metadata
  setMetadata:  (m: Metadata) => void
  inventory:    Inventory
  setInventory: (i: Inventory) => void
}

interface PreviewRow {
  kind: 'cable' | 'housing' | 'ferrule'
  key: string; pai: string; type: string
  compIdx: number  // 0 for cable/ferrule; component index for housing
  품번: string; 품명: string; 구매처: string; 리드타임: string
  현재고: number; 기발주: number
  isNew: boolean
}

interface ImportRow {
  code: string; name: string; total: number; byWh: Record<string, number>
  matchType: 'cable' | 'housing' | 'ferrule' | null; metaKey: string | null
  compIdx: number
}

const DEFAULT_COMP: HousingComp          = { 품번: '', 품명: '', 구매처: '', 리드타임: null }
const DEFAULT_INV:  InventoryHousingItem = { 현재고: 0, 기발주: 0 }

interface PrevSnapshot {
  date:    string
  cable:   Record<string, { 현재고: number; 기발주: number }>
  housing: Record<string, Array<{ 현재고: number; 기발주: number }>>
  ferrule: Record<string, { 현재고: number; 기발주: number }>
}

function DeltaBadge({ curr, prev }: { curr: number; prev: number | null }) {
  if (prev === null) return null
  const d = curr - prev
  if (d === 0) return <span className="text-gray-300 text-[10px] ml-1">±0</span>
  return d > 0
    ? <span className="text-green-600 text-[10px] font-bold ml-1">▲{d.toLocaleString()}</span>
    : <span className="text-red-500 text-[10px] font-bold ml-1">▼{Math.abs(d).toLocaleString()}</span>
}

const thCls = 'px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-200 bg-gray-50 whitespace-nowrap text-left'
const thR   = `${thCls} text-right`

export default function MaterialManager({ metadata: metadataRaw, setMetadata, inventory: inventoryRaw, setInventory }: Props) {
  const metadata  = { ...metadataRaw,  ferrule: metadataRaw.ferrule  ?? {} }
  const inventory = { ...inventoryRaw, ferrule: inventoryRaw.ferrule ?? {} }
  const [tab, setTab] = useState<'cable' | 'housing' | 'ferrule'>('cable')
  const [saved, setSaved] = useState<'' | 'both' | 'meta' | 'inv'>('')

  // 품번 Excel 업로드 상태
  const [pnPanel,    setPnPanel]    = useState(false)
  const [preview,    setPreview]    = useState<{ rows: PreviewRow[] } | null>(null)
  const [previewErr, setPreviewErr] = useState('')
  const pnFileRef = useRef<HTMLInputElement>(null)

  // ERP 재고 업로드 상태
  const [erpPanel,       setErpPanel]       = useState(false)
  const [importRows,     setImportRows]     = useState<ImportRow[] | null>(null)
  const [warehouses,     setWarehouses]     = useState<string[]>([])
  const [selectedWh,     setSelectedWh]     = useState('')
  const [importSubTab,   setImportSubTab]   = useState<'total' | 'warehouse'>('total')
  const [importApplied,  setImportApplied]  = useState(false)
  const [importFileName, setImportFileName] = useState('')
  const erpFileRef = useRef<HTMLInputElement>(null)

  // 의왕 재고현황 파일 업로드 상태
  interface StockMatchRow {
    품번: string; 품명: string; 총재고: number
    matchType: 'cable' | 'housing' | 'ferrule' | null
    metaKey: string | null; compIdx: number
  }
  const [stockPanel,    setStockPanel]    = useState(false)
  const [stockRows,     setStockRows]     = useState<StockMatchRow[] | null>(null)
  const [stockApplied,  setStockApplied]  = useState(false)
  const [stockFileName, setStockFileName] = useState('')
  const stockFileRef = useRef<HTMLInputElement>(null)

  // 드래그 상태
  const [pnDragging,    setPnDragging]    = useState(false)
  const [erpDragging,   setErpDragging]   = useState(false)
  const [stockDragging, setStockDragging] = useState(false)
  const pnDragCnt    = useRef(0)
  const erpDragCnt   = useRef(0)
  const stockDragCnt = useRef(0)

  function makeDrop(
    setDragging: (v: boolean) => void,
    counter: React.MutableRefObject<number>,
    onFile: (f: File) => void,
  ) {
    return {
      onDragOver:  (e: React.DragEvent) => e.preventDefault(),
      onDragEnter: (e: React.DragEvent) => { e.preventDefault(); counter.current++; setDragging(true) },
      onDragLeave: (e: React.DragEvent) => { e.preventDefault(); if (--counter.current === 0) setDragging(false) },
      onDrop: (e: React.DragEvent) => {
        e.preventDefault(); counter.current = 0; setDragging(false)
        const f = e.dataTransfer.files?.[0]; if (f) onFile(f)
      },
    }
  }

  // ── 주간 재고 비교 스냅샷 (ERP 업로드 적용 시 자동 저장) ──────
  const [prevSnapshot, setPrevSnapshot] = useState<PrevSnapshot | null>(() => {
    try { const s = localStorage.getItem('ajw_inv_snapshot_prev'); return s ? JSON.parse(s) : null } catch { return null }
  })

  // ── 구형 자재 숨김 키워드 + 검색 필터 ────────────────────────
  const [hideKeyword, setHideKeyword] = useState('')
  const [searchQuery, setSearchQuery]  = useState('')

  // ── 하우징 다중 부품 헬퍼 ─────────────────────────────────────
  function getHousingComps(key: string): HousingComp[] {
    const raw = metadata.housing[key]
    if (!raw) return [{ ...DEFAULT_COMP }]
    return Array.isArray(raw) ? raw : [raw]
  }
  function getHousingInvs(key: string): InventoryHousingItem[] {
    const raw = inventory.housing[key]
    if (!raw) return [{ ...DEFAULT_INV }]
    return Array.isArray(raw) ? raw : [raw]
  }
  function packComps(comps: HousingComp[]): HousingComp | HousingComp[] {
    return comps.length === 1 ? comps[0] : comps
  }
  function packInvs(invs: InventoryHousingItem[]): InventoryHousingItem | InventoryHousingItem[] {
    return invs.length === 1 ? invs[0] : invs
  }

  // ── 메타 수정 ──────────────────────────────────────────────────
  const updateCableMeta = (key: string, field: keyof CableMeta, value: string) => {
    setMetadata({ ...metadata, cable: {
      ...metadata.cable,
      [key]: { ...(metadata.cable[key] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }), [field]: value },
    }})
  }
  const updateHousingCompMeta = (key: string, ci: number, field: keyof HousingComp, value: string) => {
    const comps = [...getHousingComps(key)]
    comps[ci] = { ...comps[ci], [field]: value }
    setMetadata({ ...metadata, housing: { ...metadata.housing, [key]: packComps(comps) } })
  }
  const updateFerruleMeta = (key: string, field: keyof FerruleMeta, value: string) => {
    setMetadata({ ...metadata, ferrule: {
      ...metadata.ferrule,
      [key]: { ...(metadata.ferrule[key] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }), [field]: value },
    }})
  }

  // ── 재고 수정 ──────────────────────────────────────────────────
  const updateCableInv = (key: string, field: '현재고' | '기발주', value: string) => {
    const prev = inventory.cable[key] ?? { 현재고: 0, 기발주: 0 }
    setInventory({ ...inventory, cable: { ...inventory.cable, [key]: { ...prev, [field]: parseInt(value) || 0 } } })
  }
  const updateHousingCompInv = (key: string, ci: number, field: '현재고' | '기발주', value: string) => {
    const invs = [...getHousingInvs(key)]
    while (invs.length <= ci) invs.push({ ...DEFAULT_INV })
    invs[ci] = { ...invs[ci], [field]: parseInt(value) || 0 }
    setInventory({ ...inventory, housing: { ...inventory.housing, [key]: packInvs(invs) } })
  }
  const updateFerruleInv = (key: string, field: '현재고' | '기발주', value: string) => {
    const prev = inventory.ferrule[key] ?? { 현재고: 0, 기발주: 0 }
    setInventory({ ...inventory, ferrule: { ...inventory.ferrule, [key]: { ...prev, [field]: parseInt(value) || 0 } } })
  }

  // ── 하우징 다중 부품 추가 / 제거 ─────────────────────────────
  const addHousingComp = (key: string) => {
    const comps = [...getHousingComps(key), { ...DEFAULT_COMP }]
    const invs  = [...getHousingInvs(key),  { ...DEFAULT_INV }]
    setMetadata({ ...metadata, housing: { ...metadata.housing, [key]: packComps(comps) } })
    setInventory({ ...inventory, housing: { ...inventory.housing, [key]: packInvs(invs) } })
  }
  const removeHousingComp = (key: string, ci: number) => {
    const comps = [...getHousingComps(key)]
    const invs  = [...getHousingInvs(key)]
    if (comps.length <= 1) return
    comps.splice(ci, 1)
    invs.splice(ci, 1)
    setMetadata({ ...metadata, housing: { ...metadata.housing, [key]: packComps(comps) } })
    setInventory({ ...inventory, housing: { ...inventory.housing, [key]: packInvs(invs) } })
  }

  // ── 저장 ───────────────────────────────────────────────────────
  const flash = (type: typeof saved) => { setSaved(type); setTimeout(() => setSaved(''), 2000) }
  const saveAll = () => { saveMetadata(metadata); saveInventory(inventory); flash('both') }

  // ── 양식 다운로드 (케이블 + 하우징 + 페룰 3시트) ──────────────
  function downloadForm() {
    const wb = XLSX.utils.book_new()

    // 케이블 시트
    const cableKeys   = Object.keys(metadata.cable).sort()
    const cableHeader = ['파이', '케이블종류', '품번', '품명', '구매처', '리드타임(일)', '현재고(m)', '기발주(m)']
    const cableRows   = cableKeys.length
      ? cableKeys.map(k => {
          const [pai, type] = k.split('|')
          const m   = metadata.cable[k] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }
          const inv = inventory.cable[k] ?? { 현재고: 0, 기발주: 0 }
          return [pai, type, m.품번 ?? '', m.품명 ?? '', m.구매처 ?? '', m.리드타임 ?? '', inv.현재고 ?? 0, inv.기발주 ?? 0]
        })
      : [['A1.6', '예시종류', '', '', '', '60', 0, 0]]
    const cableWs = XLSX.utils.aoa_to_sheet([cableHeader, ...cableRows])
    cableWs['!cols'] = [8, 20, 16, 30, 16, 12, 12, 12].map(w => ({ wch: w }))
    XLSX.utils.book_append_sheet(wb, cableWs, '케이블')

    // 하우징 시트 — 다중 부품 전개 (같은 파이+타입이 여러 행으로)
    const housingKeys   = Object.keys(metadata.housing).sort()
    const housingHeader = ['파이', '하우징타입', '품번', '품명', '구매처', '리드타임(일)', '현재고(EA)', '기발주(EA)']
    const housingRows   = housingKeys.length
      ? housingKeys.flatMap(k => {
          const [pai, type] = k.split('|')
          const comps = getHousingComps(k)
          const invs  = getHousingInvs(k)
          return comps.map((m, ci) => {
            const inv = invs[ci] ?? DEFAULT_INV
            return [pai, type, m.품번 ?? '', m.품명 ?? '', m.구매처 ?? '', m.리드타임 ?? '', inv.현재고 ?? 0, inv.기발주 ?? 0]
          })
        })
      : [['A1.6', '예시타입', '', '', '', '60', 0, 0]]
    const housingWs = XLSX.utils.aoa_to_sheet([housingHeader, ...housingRows])
    housingWs['!cols'] = [8, 20, 16, 30, 16, 12, 12, 12].map(w => ({ wch: w }))
    XLSX.utils.book_append_sheet(wb, housingWs, '하우징')

    // 페룰 시트
    const ferruleKeys   = Object.keys(metadata.ferrule).sort()
    const ferruleHeader = ['페룰타입', '품번', '품명', '구매처', '리드타임(일)', '현재고(EA)', '기발주(EA)']
    const ferruleRows   = ferruleKeys.length
      ? ferruleKeys.map(k => {
          const m   = metadata.ferrule[k] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }
          const inv = inventory.ferrule[k] ?? { 현재고: 0, 기발주: 0 }
          return [k, m.품번 ?? '', m.품명 ?? '', m.구매처 ?? '', m.리드타임 ?? '', inv.현재고 ?? 0, inv.기발주 ?? 0]
        })
      : [['LC/PC 청색', '', '', '', '30', 0, 0]]
    const ferruleWs = XLSX.utils.aoa_to_sheet([ferruleHeader, ...ferruleRows])
    ferruleWs['!cols'] = [20, 16, 30, 16, 12, 12, 12].map(w => ({ wch: w }))
    XLSX.utils.book_append_sheet(wb, ferruleWs, '페룰')

    XLSX.writeFile(wb, `자재관리_양식_${today()}.xlsx`)
  }

  // ── 양식 Excel 업로드 파싱 ────────────────────────────────────
  function handlePnUpload(file: File) {
    setPreviewErr('')
    const reader = new FileReader()
    reader.onload = ev => {
      try {
        const wb   = XLSX.read(ev.target?.result, { type: 'array' })
        const parsed: PreviewRow[] = []

        for (const kind of ['cable', 'housing', 'ferrule'] as const) {
          const sheetName = kind === 'cable' ? '케이블' : kind === 'housing' ? '하우징' : '페룰'
          const ws = wb.Sheets[sheetName]
          if (!ws) continue
          const raw = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' }) as string[][]

          if (kind === 'ferrule') {
            for (let i = 1; i < raw.length; i++) {
              const r = raw[i]
              const type = String(r[0] ?? '').trim()
              if (!type) continue
              parsed.push({
                kind, key: type, pai: '', type, compIdx: 0,
                품번: String(r[1] ?? '').trim(), 품명: String(r[2] ?? '').trim(),
                구매처: String(r[3] ?? '').trim(), 리드타임: String(r[4] ?? '').trim(),
                현재고: Number(r[5]) || 0, 기발주: Number(r[6]) || 0,
                isNew: !metadata.ferrule[type],
              })
            }
          } else if (kind === 'cable') {
            for (let i = 1; i < raw.length; i++) {
              const r = raw[i]
              const pai  = String(r[0] ?? '').trim()
              const type = String(r[1] ?? '').trim()
              if (!pai || !type) continue
              const key = `${pai}|${type}`
              parsed.push({
                kind, key, pai, type, compIdx: 0,
                품번: String(r[2] ?? '').trim(), 품명: String(r[3] ?? '').trim(),
                구매처: String(r[4] ?? '').trim(), 리드타임: String(r[5] ?? '').trim(),
                현재고: Number(r[6]) || 0, 기발주: Number(r[7]) || 0,
                isNew: !metadata.cable[key],
              })
            }
          } else {
            // housing — 같은 파이+타입이 연속으로 나오면 같은 키의 다중 부품
            const compCountPerKey: Record<string, number> = {}
            for (let i = 1; i < raw.length; i++) {
              const r = raw[i]
              const pai  = String(r[0] ?? '').trim()
              const type = String(r[1] ?? '').trim()
              if (!pai || !type) continue
              const key = `${pai}|${type}`
              const ci = compCountPerKey[key] ?? 0
              compCountPerKey[key] = ci + 1
              parsed.push({
                kind, key, pai, type, compIdx: ci,
                품번: String(r[2] ?? '').trim(), 품명: String(r[3] ?? '').trim(),
                구매처: String(r[4] ?? '').trim(), 리드타임: String(r[5] ?? '').trim(),
                현재고: Number(r[6]) || 0, 기발주: Number(r[7]) || 0,
                isNew: !metadata.housing[key],
              })
            }
          }
        }

        if (!parsed.length) { setPreviewErr('유효한 데이터가 없습니다. 케이블/하우징 시트를 확인하세요.'); return }
        setPreview({ rows: parsed })
      } catch { setPreviewErr('파일을 읽는 중 오류가 발생했습니다.') }
    }
    reader.readAsArrayBuffer(file)
  }

  function applyPreview() {
    if (!preview) return
    const nextMeta = { ...metadata }
    const nextInv  = { ...inventory }
    const cable      = { ...nextMeta.cable }
    const housing    = { ...nextMeta.housing }
    const ferrule    = { ...nextMeta.ferrule }
    const cableInv   = { ...nextInv.cable }
    const housingInv = { ...nextInv.housing }
    const ferruleInv = { ...nextInv.ferrule }

    // housing: 키별로 컴포넌트 배열 누적
    const housingCompsByKey: Record<string, Array<{ meta: HousingComp; inv: InventoryHousingItem }>> = {}
    for (const r of preview.rows) {
      if (r.kind !== 'housing') continue
      if (!housingCompsByKey[r.key]) housingCompsByKey[r.key] = []
      housingCompsByKey[r.key][r.compIdx] = {
        meta: { 품번: r.품번, 품명: r.품명, 구매처: r.구매처, 리드타임: r.리드타임 || null },
        inv:  { 현재고: r.현재고, 기발주: r.기발주 },
      }
    }
    for (const [key, entries] of Object.entries(housingCompsByKey)) {
      const comps = entries.map(e => e.meta)
      const invs  = entries.map(e => e.inv)
      housing[key]    = packComps(comps)
      housingInv[key] = packInvs(invs)
    }

    for (const r of preview.rows) {
      if (r.kind === 'cable') {
        cable[r.key]    = { 품번: r.품번, 품명: r.품명, 구매처: r.구매처, 리드타임: r.리드타임 || null }
        const prev      = cableInv[r.key] ?? { 현재고: 0, 기발주: 0 }
        cableInv[r.key] = { ...prev, 현재고: r.현재고, 기발주: r.기발주 }
      } else if (r.kind === 'ferrule') {
        ferrule[r.key]    = { 품번: r.품번, 품명: r.품명, 구매처: r.구매처, 리드타임: r.리드타임 || null }
        const prev        = ferruleInv[r.key] ?? { 현재고: 0, 기발주: 0 }
        ferruleInv[r.key] = { ...prev, 현재고: r.현재고, 기발주: r.기발주 }
      }
    }

    nextMeta.cable = cable; nextMeta.housing = housing; nextMeta.ferrule = ferrule
    nextInv.cable  = cableInv; nextInv.housing = housingInv; nextInv.ferrule = ferruleInv
    setMetadata(nextMeta); saveMetadata(nextMeta)
    setInventory(nextInv); saveInventory(nextInv)
    setPreview(null); flash('both')
  }

  // ── ERP 재고 파일 파싱 ────────────────────────────────────────
  function handleErpUpload(file: File) {
    setImportFileName(file.name)
    const reader = new FileReader()
    reader.onload = ev => {
      const wb = XLSX.read(ev.target?.result, { type: 'array' })
      const ws = wb.Sheets['재고현황']
      if (!ws) { alert('재고현황 시트를 찾을 수 없습니다.'); return }
      const range = XLSX.utils.decode_range(ws['!ref']!)
      const headerRow: string[] = []
      for (let c = 0; c <= range.e.c; c++)
        headerRow.push(ws[XLSX.utils.encode_cell({ r: 1, c })]?.v?.toString().trim() ?? '')
      const whCols = headerRow.slice(4).filter(h => h)
      setWarehouses(whCols); setSelectedWh(whCols[0] ?? '')

      // 품번 → 자재 키 맵 (케이블/페룰)
      const cableMap:   Record<string, string> = {}
      const ferruleMap: Record<string, string> = {}
      Object.entries(metadata.cable).forEach(([k, m]) => { if (m.품번) cableMap[m.품번] = k })
      Object.entries(metadata.ferrule).forEach(([k, m]) => { if (m.품번) ferruleMap[m.품번] = k })

      // 하우징: 품번 → { key, compIdx }
      const housingCompMap: Record<string, { key: string; ci: number }> = {}
      Object.entries(metadata.housing).forEach(([k, mRaw]) => {
        const comps = Array.isArray(mRaw) ? mRaw : [mRaw]
        comps.forEach((m, ci) => { if (m?.품번) housingCompMap[m.품번] = { key: k, ci } })
      })

      const rows: ImportRow[] = []
      for (let r = 2; r <= range.e.r; r++) {
        const code  = ws[XLSX.utils.encode_cell({ r, c: 0 })]?.v?.toString().trim() ?? ''
        const name  = ws[XLSX.utils.encode_cell({ r, c: 1 })]?.v?.toString().trim() ?? ''
        const total = Number(ws[XLSX.utils.encode_cell({ r, c: 3 })]?.v ?? 0)
        if (!code) continue
        const byWh: Record<string, number> = {}
        whCols.forEach((wh, wi) => { byWh[wh] = Number(ws[XLSX.utils.encode_cell({ r, c: 4 + wi })]?.v ?? 0) })
        let matchType: 'cable' | 'housing' | 'ferrule' | null = null
        let metaKey: string | null = null
        let compIdx = 0
        if (cableMap[code])              { matchType = 'cable';   metaKey = cableMap[code] }
        else if (housingCompMap[code])   { matchType = 'housing'; metaKey = housingCompMap[code].key; compIdx = housingCompMap[code].ci }
        else if (ferruleMap[code])       { matchType = 'ferrule'; metaKey = ferruleMap[code] }
        rows.push({ code, name, total, byWh, matchType, metaKey, compIdx })
      }
      setImportRows(rows); setImportApplied(false)
    }
    reader.readAsArrayBuffer(file)
  }

  function applyErpImport(useTotal: boolean) {
    if (!importRows) return
    // 현재 재고를 이전 주 스냅샷으로 보관 (ERP 업로드 적용 직전)
    const snap: PrevSnapshot = {
      date: new Date().toLocaleDateString('ko-KR'),
      cable:   Object.fromEntries(Object.entries(inventory.cable).map(([k, v]) => [k, { 현재고: v?.현재고 ?? 0, 기발주: v?.기발주 ?? 0 }])),
      housing: Object.fromEntries(Object.entries(inventory.housing).map(([k, vRaw]) => {
        const arr = Array.isArray(vRaw) ? vRaw : [vRaw as InventoryHousingItem]
        return [k, arr.map(v => ({ 현재고: v?.현재고 ?? 0, 기발주: v?.기발주 ?? 0 }))]
      })),
      ferrule: Object.fromEntries(Object.entries(inventory.ferrule).map(([k, v]) => [k, { 현재고: v?.현재고 ?? 0, 기발주: v?.기발주 ?? 0 }])),
    }
    localStorage.setItem('ajw_inv_snapshot_prev', JSON.stringify(snap))
    setPrevSnapshot(snap)
    const newCable    = { ...inventory.cable }
    const newHousing  = { ...inventory.housing }
    const newFerrule  = { ...inventory.ferrule }
    importRows.forEach(row => {
      if (!row.matchType || !row.metaKey) return
      const qty = useTotal ? row.total : (row.byWh[selectedWh] ?? 0)
      if (row.matchType === 'cable') {
        const prev = newCable[row.metaKey] ?? { 현재고: 0, 기발주: 0 }
        newCable[row.metaKey] = { ...prev, 현재고: qty }
      } else if (row.matchType === 'housing') {
        const invs = Array.isArray(newHousing[row.metaKey])
          ? [...(newHousing[row.metaKey] as InventoryHousingItem[])]
          : [newHousing[row.metaKey] as InventoryHousingItem ?? { ...DEFAULT_INV }]
        while (invs.length <= row.compIdx) invs.push({ ...DEFAULT_INV })
        invs[row.compIdx] = { ...(invs[row.compIdx] ?? DEFAULT_INV), 현재고: qty }
        newHousing[row.metaKey] = packInvs(invs)
      } else {
        const prev = newFerrule[row.metaKey] ?? { 현재고: 0, 기발주: 0 }
        newFerrule[row.metaKey] = { ...prev, 현재고: qty }
      }
    })
    setInventory({ cable: newCable, housing: newHousing, ferrule: newFerrule })
    setImportApplied(true)
  }

  // ── 의왕 재고현황 파일 업로드 + 매칭 ─────────────────────────
  function handleStockUpload(file: File) {
    setStockFileName(file.name)
    const reader = new FileReader()
    reader.onload = ev => {
      try {
        const parsed = parseMaterialStock(ev.target?.result as ArrayBuffer)
        saveMaterialStock(parsed)

        // 품번 역방향 맵 생성
        const cableMap:   Record<string, string> = {}
        const ferruleMap: Record<string, string> = {}
        Object.entries(metadata.cable).forEach(([k, m]) => { if (m.품번) cableMap[m.품번] = k })
        Object.entries(metadata.ferrule).forEach(([k, m]) => { if (m.품번) ferruleMap[m.품번] = k })
        const housingCompMap: Record<string, { key: string; ci: number }> = {}
        Object.entries(metadata.housing).forEach(([k, mRaw]) => {
          const comps = Array.isArray(mRaw) ? mRaw : [mRaw]
          comps.forEach((m, ci) => { if (m?.품번) housingCompMap[m.품번] = { key: k, ci } })
        })

        const rows: StockMatchRow[] = Object.entries(parsed).map(([품번, item]) => {
          let matchType: 'cable' | 'housing' | 'ferrule' | null = null
          let metaKey: string | null = null
          let compIdx = 0
          if (cableMap[품번])             { matchType = 'cable';   metaKey = cableMap[품번] }
          else if (housingCompMap[품번])  { matchType = 'housing'; metaKey = housingCompMap[품번].key; compIdx = housingCompMap[품번].ci }
          else if (ferruleMap[품번])      { matchType = 'ferrule'; metaKey = ferruleMap[품번] }
          return { 품번, 품명: item.품명, 총재고: item.총재고, matchType, metaKey, compIdx }
        })
        setStockRows(rows); setStockApplied(false)
      } catch (err) { alert(`파일 파싱 오류: ${err}`) }
    }
    reader.readAsArrayBuffer(file)
  }

  function applyStockImport() {
    if (!stockRows) return
    const snap: PrevSnapshot = {
      date: new Date().toLocaleDateString('ko-KR'),
      cable:   Object.fromEntries(Object.entries(inventory.cable).map(([k, v]) => [k, { 현재고: v?.현재고 ?? 0, 기발주: v?.기발주 ?? 0 }])),
      housing: Object.fromEntries(Object.entries(inventory.housing).map(([k, vRaw]) => {
        const arr = Array.isArray(vRaw) ? vRaw : [vRaw as InventoryHousingItem]
        return [k, arr.map(v => ({ 현재고: v?.현재고 ?? 0, 기발주: v?.기발주 ?? 0 }))]
      })),
      ferrule: Object.fromEntries(Object.entries(inventory.ferrule).map(([k, v]) => [k, { 현재고: v?.현재고 ?? 0, 기발주: v?.기발주 ?? 0 }])),
    }
    localStorage.setItem('ajw_inv_snapshot_prev', JSON.stringify(snap))
    setPrevSnapshot(snap)

    const newCable   = { ...inventory.cable }
    const newHousing = { ...inventory.housing }
    const newFerrule = { ...inventory.ferrule }
    stockRows.forEach(row => {
      if (!row.matchType || !row.metaKey) return
      if (row.matchType === 'cable') {
        newCable[row.metaKey] = { ...(newCable[row.metaKey] ?? { 현재고: 0, 기발주: 0 }), 현재고: row.총재고 }
      } else if (row.matchType === 'housing') {
        const invs = Array.isArray(newHousing[row.metaKey])
          ? [...(newHousing[row.metaKey] as InventoryHousingItem[])]
          : [newHousing[row.metaKey] as InventoryHousingItem ?? { ...DEFAULT_INV }]
        while (invs.length <= row.compIdx) invs.push({ ...DEFAULT_INV })
        invs[row.compIdx] = { ...(invs[row.compIdx] ?? DEFAULT_INV), 현재고: row.총재고 }
        newHousing[row.metaKey] = packInvs(invs)
      } else {
        newFerrule[row.metaKey] = { ...(newFerrule[row.metaKey] ?? { 현재고: 0, 기발주: 0 }), 현재고: row.총재고 }
      }
    })
    const next = { cable: newCable, housing: newHousing, ferrule: newFerrule }
    setInventory(next); saveInventory(next)
    setStockApplied(true)
  }

  // ── 행 삭제 ───────────────────────────────────────────────────
  const deleteCableKey = (key: string) => {
    if (!window.confirm(`케이블 항목 "${key}"를 삭제하시겠습니까?\n(저장 버튼을 눌러야 DB에 반영됩니다)`)) return
    const cable = { ...metadata.cable }; delete cable[key]
    const cableInv = { ...inventory.cable }; delete cableInv[key]
    setMetadata({ ...metadata, cable })
    setInventory({ ...inventory, cable: cableInv })
  }
  const deleteFerruleKey = (key: string) => {
    if (!window.confirm(`페룰 항목 "${key}"를 삭제하시겠습니까?\n(저장 버튼을 눌러야 DB에 반영됩니다)`)) return
    const ferrule = { ...metadata.ferrule }; delete ferrule[key]
    const ferruleInv = { ...inventory.ferrule }; delete ferruleInv[key]
    setMetadata({ ...metadata, ferrule })
    setInventory({ ...inventory, ferrule: ferruleInv })
  }

  // ── 공통 헬퍼 ─────────────────────────────────────────────────
  const getCm  = (k: string): CableMeta   => metadata.cable[k] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }
  const getFm  = (k: string): FerruleMeta => metadata.ferrule[k] ?? { 품번:'', 품명:'', 구매처:'', 리드타임:null }
  const getCiv = (k: string) => inventory.cable[k] ?? { 현재고: 0, 기발주: 0 }
  const getFiv = (k: string) => inventory.ferrule[k] ?? { 현재고: 0, 기발주: 0 }

  const cableKeys    = Object.keys(metadata.cable).sort()
  const housingKeys  = Object.keys(metadata.housing).sort()
  const ferruleKeys  = Object.keys(metadata.ferrule).sort()
  const cableMissing   = cableKeys.filter(k => !getCm(k).품번).length
  // 하우징: 어떤 부품이라도 품번 미입력이면 ⚠
  const housingMissing = housingKeys.filter(k => getHousingComps(k).some(c => !c.품번)).length
  const ferruleMissing = ferruleKeys.filter(k => !getFm(k).품번).length
  const matched   = importRows?.filter(r => r.matchType) ?? []
  const unmatched = importRows?.filter(r => !r.matchType) ?? []

  // ── 구형 숨김 + 검색 필터 ─────────────────────────────────────
  const hideWords = hideKeyword.split(/[,\s]+/).map(w => w.trim().toLowerCase()).filter(Boolean)
  const sq = searchQuery.trim().toLowerCase()

  function isHidden(pai: string, type: string) {
    return hideWords.length > 0 && hideWords.some(w => pai.toLowerCase().includes(w) || type.toLowerCase().includes(w))
  }
  function matchesSearch(key: string, kind: 'cable' | 'housing' | 'ferrule') {
    if (!sq) return true
    const [pai, type] = kind === 'ferrule' ? ['', key] : key.split('|')
    if (pai.toLowerCase().includes(sq) || type.toLowerCase().includes(sq)) return true
    if (kind === 'cable')   return (getCm(key).품번 ?? '').toLowerCase().includes(sq) || (getCm(key).품명 ?? '').toLowerCase().includes(sq)
    if (kind === 'ferrule') return (getFm(key).품번 ?? '').toLowerCase().includes(sq) || (getFm(key).품명 ?? '').toLowerCase().includes(sq)
    return getHousingComps(key).some(c => (c.품번 ?? '').toLowerCase().includes(sq) || (c.품명 ?? '').toLowerCase().includes(sq))
  }
  const filteredCableKeys   = cableKeys.filter(k => { const [p, t] = k.split('|'); return !isHidden(p, t) && matchesSearch(k, 'cable') })
  const filteredHousingKeys = housingKeys.filter(k => { const [p, t] = k.split('|'); return !isHidden(p, t) && matchesSearch(k, 'housing') })
  const filteredFerruleKeys = ferruleKeys.filter(k => !isHidden('', k) && matchesSearch(k, 'ferrule'))

  // 하우징 표시용 전개 행
  interface HousingDisplayRow {
    key: string; ci: number; pai: string; type: string
    comp: HousingComp; inv: InventoryHousingItem
    missing: boolean; isFirst: boolean; isLast: boolean; total: number
  }
  const housingDisplayRows: HousingDisplayRow[] = filteredHousingKeys.flatMap(key => {
    const [pai, type] = key.split('|')
    const comps = getHousingComps(key)
    const invs  = getHousingInvs(key)
    return comps.map((comp, ci) => ({
      key, ci, pai, type,
      comp: comp ?? { ...DEFAULT_COMP },
      inv:  invs[ci] ?? { ...DEFAULT_INV },
      missing: !comp?.품번,
      isFirst: ci === 0,
      isLast:  ci === comps.length - 1,
      total:   comps.length,
    }))
  })

  const inputCls    = 'w-full border border-gray-200 rounded px-2 py-0.5 text-xs focus:outline-none focus:border-blue-400'
  const inputRoCls  = `${inputCls} bg-gray-50 text-gray-400 cursor-default`
  const inputInvCls = 'w-24 border border-gray-200 rounded px-2 py-0.5 text-xs text-right focus:outline-none focus:border-blue-400'

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>자재 관리</b> — 품번·구매처·리드타임과 현재고·기발주를 한 화면에서 관리합니다.
        {!CAN_WRITE && <span className="ml-2 text-gray-500">(품번 정보는 읽기 전용)</span>}
      </div>

      {/* ── 툴바 ─────────────────────────────────────────────── */}
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="flex gap-2">
          {([
            ['cable',   '케이블', cableKeys.length,   cableMissing],
            ['housing', '하우징', housingKeys.length,  housingMissing],
            ['ferrule', '페롤',   ferruleKeys.length,  ferruleMissing],
          ] as [typeof tab, string, number, number][]).map(([id, lbl, cnt, miss]) => (
            <button key={id} onClick={() => { setTab(id); setPreview(null) }}
              className={`px-4 py-2 rounded font-semibold text-sm flex items-center gap-2 transition ${
                tab === id ? 'bg-[#2E75B6] text-white' : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
              }`}>
              {lbl}
              <span className={`text-xs px-1.5 py-0.5 rounded-full font-bold ${tab === id ? 'bg-white/20' : 'bg-gray-300'}`}>{cnt}</span>
              {miss > 0 && <span className="text-xs bg-red-500 text-white px-1.5 py-0.5 rounded-full">⚠ {miss}</span>}
            </button>
          ))}
        </div>

        <div className="flex items-center gap-2">
          {CAN_WRITE && (
            <>
              <button onClick={downloadForm}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border border-gray-300 rounded bg-white hover:bg-gray-50 transition">
                📤 양식 다운로드
              </button>
              <button onClick={() => { setPnPanel(v => !v); setErpPanel(false) }}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border rounded transition ${
                  pnPanel ? 'border-[#2E75B6] bg-blue-50 text-[#2E75B6]' : 'border-[#2E75B6] text-[#2E75B6] bg-white hover:bg-blue-50'
                }`}>
                📥 양식 업로드
              </button>
            </>
          )}
          <button onClick={() => { setErpPanel(v => !v); setPnPanel(false); setStockPanel(false) }}
            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border rounded transition ${
              erpPanel ? 'border-yellow-500 bg-yellow-50 text-yellow-700' : 'border-yellow-500 text-yellow-700 bg-white hover:bg-yellow-50'
            }`}>
            🏭 ERP 재고 업로드
          </button>
          <button onClick={() => { setStockPanel(v => !v); setErpPanel(false); setPnPanel(false) }}
            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border rounded transition ${
              stockPanel ? 'border-green-600 bg-green-50 text-green-700' : 'border-green-600 text-green-700 bg-white hover:bg-green-50'
            }`}>
            📋 의왕 재고현황 업로드
          </button>
        </div>
      </div>

      {/* ── 양식 Excel 업로드 패널 ──────────────────────────── */}
      {pnPanel && CAN_WRITE && (
        <div className={`border rounded-lg p-4 space-y-3 transition-colors ${pnDragging ? 'border-[#2E75B6] bg-[#dbeeff]' : 'border-blue-200 bg-blue-50'}`}
          {...makeDrop(setPnDragging, pnDragCnt, handlePnUpload)}>
          {pnDragging && <div className="text-center text-sm font-semibold text-[#2E75B6] py-2">📂 여기에 놓으세요</div>}
          <div className="flex items-center justify-between">
            <span className="text-sm font-semibold text-blue-800">
              양식 업로드 — 케이블 + 하우징 동시 적용
            </span>
            <button onClick={() => { setPnPanel(false); setPreview(null); setPreviewErr('') }}
              className="text-xs text-gray-400 hover:text-gray-600">✕</button>
          </div>
          <p className="text-xs text-blue-600">
            양식 다운로드로 받은 <b>자재관리_양식.xlsx</b> 파일을 수정 후 업로드하세요.
            하우징은 동일한 파이+타입이 여러 행이면 다중 부품으로 처리됩니다.
          </p>
          {previewErr && <div className="text-sm text-red-600 bg-red-50 border border-red-200 px-3 py-2 rounded">{previewErr}</div>}
          {!preview ? (
            <button onClick={() => pnFileRef.current?.click()}
              className="px-4 py-2 bg-[#2E75B6] text-white text-sm font-semibold rounded hover:bg-[#1f5a9a] transition">
              📂 Excel 파일 선택
            </button>
          ) : (
            <div className="space-y-2">
              <div className="text-xs text-blue-700 flex gap-4">
                <span>케이블 <b>{preview.rows.filter(r => r.kind==='cable').length}</b>행</span>
                <span>하우징 <b>{preview.rows.filter(r => r.kind==='housing').length}</b>행</span>
                <span>신규 <b className="text-green-600">{preview.rows.filter(r => r.isNew).length}</b>건 / 업데이트 <b className="text-blue-600">{preview.rows.filter(r => !r.isNew).length}</b>건</span>
              </div>
              <div className="overflow-x-auto max-h-60 overflow-y-auto rounded border border-blue-100">
                <table className="text-xs w-full bg-white">
                  <thead className="sticky top-0 bg-blue-100 text-blue-800">
                    <tr>{['구분','파이','종류','#','품번','품명','구매처','LT','현재고','기발주','상태'].map(h=>(
                      <th key={h} className="px-2 py-1.5 text-left font-semibold whitespace-nowrap">{h}</th>
                    ))}</tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {preview.rows.map((r, ri) => (
                      <tr key={`${r.kind}-${r.key}-${ri}`} className={r.isNew ? 'bg-green-50' : r.compIdx > 0 ? 'bg-purple-50' : ''}>
                        <td className="px-2 py-1">
                          <span className={`text-xs px-1.5 py-0.5 rounded-full font-medium ${r.kind==='cable'?'bg-blue-100 text-blue-700':r.kind==='housing'?'bg-purple-100 text-purple-700':'bg-orange-100 text-orange-700'}`}>
                            {r.kind==='cable'?'케이블':r.kind==='housing'?'하우징':'페룰'}
                          </span>
                        </td>
                        <td className="px-2 py-1 font-mono text-gray-500">{r.pai}</td>
                        <td className="px-2 py-1 font-semibold">{r.compIdx > 0 ? <span className="text-purple-400">└ 동일</span> : r.type}</td>
                        <td className="px-2 py-1 text-center text-gray-400">{r.kind==='housing' ? r.compIdx+1 : ''}</td>
                        <td className="px-2 py-1">{r.품번 || <span className="text-red-400">미입력</span>}</td>
                        <td className="px-2 py-1 max-w-[120px] truncate" title={r.품명}>{r.품명 || '—'}</td>
                        <td className="px-2 py-1">{r.구매처 || '—'}</td>
                        <td className="px-2 py-1">{r.리드타임 || '—'}</td>
                        <td className="px-2 py-1 text-right font-mono">{r.현재고.toLocaleString()}</td>
                        <td className="px-2 py-1 text-right font-mono">{r.기발주.toLocaleString()}</td>
                        <td className="px-2 py-1">{r.isNew ? <span className="text-green-600 font-semibold">신규</span> : <span className="text-blue-600">업데이트</span>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex gap-2">
                <button onClick={applyPreview}
                  className="bg-[#2E75B6] hover:bg-[#1F5A9E] text-white font-bold px-5 py-2 rounded-lg text-sm transition">
                  ✅ 적용 및 저장
                </button>
                <button onClick={() => setPreview(null)} className="px-4 py-2 text-sm text-gray-500 hover:text-gray-700">취소</button>
              </div>
            </div>
          )}
          <input ref={pnFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e => { const f = e.target.files?.[0]; if (f) { handlePnUpload(f); e.target.value = '' } }} />
        </div>
      )}

      {/* ── 의왕 재고현황 업로드 패널 ───────────────────────── */}
      {stockPanel && (
        <div className={`border rounded-lg p-4 space-y-3 transition-colors ${stockDragging ? 'border-green-600 bg-[#d1fae5]' : 'border-green-200 bg-green-50'}`}
          {...makeDrop(setStockDragging, stockDragCnt, handleStockUpload)}>
          {stockDragging && <div className="text-center text-sm font-semibold text-green-700 py-2">📂 여기에 놓으세요</div>}
          <div className="flex items-center justify-between">
            <span className="text-sm font-semibold text-green-800">의왕 생산자재 재고현황 업로드</span>
            <button onClick={() => { setStockPanel(false); setStockRows(null); setStockFileName('') }}
              className="text-xs text-gray-400 hover:text-gray-600">✕</button>
          </div>
          <p className="text-xs text-green-700">
            <b>의왕_생산자재_재고현황</b> Excel 파일을 업로드하면 D열(품번) + O열(총재고)를 읽어<br />
            자재 관리의 <b>품번</b> 필드와 매칭해 현재고를 자동 갱신합니다.
            품번 미등록 항목은 건너뜁니다.
          </p>
          <div className="flex items-center gap-3">
            <button onClick={() => stockFileRef.current?.click()}
              className="px-4 py-2 bg-green-600 text-white text-sm font-semibold rounded hover:bg-green-700 transition">
              📂 파일 선택
            </button>
            {stockFileName && <span className="text-sm text-gray-600">{stockFileName}</span>}
          </div>
          {stockRows && (() => {
            const matched   = stockRows.filter(r => r.matchType)
            const unmatched = stockRows.filter(r => !r.matchType)
            return (
              <div className="space-y-2">
                <div className="flex gap-4 text-xs text-green-800">
                  <span>전체 <b>{stockRows.length}</b>개</span>
                  <span>매칭 <b className="text-green-700">{matched.length}</b>개</span>
                  <span>미매칭 <b className="text-orange-600">{unmatched.length}</b>개
                    {unmatched.length > 0 && <span className="text-gray-400 ml-1">(품번 미등록)</span>}
                  </span>
                </div>
                {matched.length > 0 && (
                  <div className="overflow-x-auto max-h-60 overflow-y-auto rounded border border-green-100">
                    <table className="text-xs w-full bg-white">
                      <thead className="sticky top-0 bg-green-100 text-green-800">
                        <tr>{['구분','품번','품명','총재고','매핑키'].map(h => (
                          <th key={h} className="px-2 py-1.5 text-left font-semibold whitespace-nowrap">{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {matched.map((r, i) => (
                          <tr key={i}>
                            <td className="px-2 py-1">
                              <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-medium ${
                                r.matchType === 'cable' ? 'bg-blue-100 text-blue-700'
                                : r.matchType === 'housing' ? 'bg-purple-100 text-purple-700'
                                : 'bg-orange-100 text-orange-700'
                              }`}>
                                {r.matchType === 'cable' ? '케이블' : r.matchType === 'housing' ? '하우징' : '페룰'}
                              </span>
                            </td>
                            <td className="px-2 py-1 font-mono">{r.품번}</td>
                            <td className="px-2 py-1 max-w-[160px] truncate" title={r.품명}>{r.품명}</td>
                            <td className="px-2 py-1 text-right font-mono font-semibold">{r.총재고.toLocaleString()}</td>
                            <td className="px-2 py-1 text-gray-500 font-mono text-[10px]">{r.metaKey}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
                {!stockApplied ? (
                  <div className="flex gap-2">
                    <button onClick={applyStockImport}
                      disabled={matched.length === 0}
                      className="bg-green-600 hover:bg-green-700 disabled:bg-gray-300 text-white font-bold px-5 py-2 rounded-lg text-sm transition">
                      ✅ {matched.length}개 현재고 반영
                    </button>
                    <button onClick={() => setStockRows(null)} className="px-4 py-2 text-sm text-gray-500 hover:text-gray-700">취소</button>
                  </div>
                ) : (
                  <div className="flex items-center gap-2 bg-green-100 border border-green-300 rounded px-4 py-2">
                    <span className="text-green-800 font-semibold text-sm">✅ {matched.length}개 현재고 반영 완료 — 저장 버튼을 눌러 확정하세요</span>
                  </div>
                )}
              </div>
            )
          })()}
          <input ref={stockFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e => { const f = e.target.files?.[0]; if (f) { handleStockUpload(f); e.target.value = '' } }} />
        </div>
      )}

      {/* ── ERP 재고 업로드 패널 ────────────────────────────── */}
      {erpPanel && (
        <div className={`border rounded-lg p-4 space-y-3 transition-colors ${erpDragging ? 'border-yellow-500 bg-[#fef3c7]' : 'border-yellow-200 bg-yellow-50'}`}
          {...makeDrop(setErpDragging, erpDragCnt, handleErpUpload)}>
          {erpDragging && <div className="text-center text-sm font-semibold text-yellow-700 py-2">📂 여기에 놓으세요</div>}
          <div className="flex items-center justify-between">
            <span className="text-sm font-semibold text-yellow-800">ERP 재고 업로드 (ESZ018R)</span>
            <button onClick={() => { setErpPanel(false); setImportRows(null); setImportFileName('') }}
              className="text-xs text-gray-400 hover:text-gray-600">✕</button>
          </div>
          <div className="flex items-center gap-4">
            <button onClick={() => erpFileRef.current?.click()}
              className="px-4 py-2 bg-[#2E75B6] text-white text-sm font-semibold rounded hover:bg-[#1f5a9a] transition">
              📂 파일 선택
            </button>
            <span className="text-sm text-gray-500">{importFileName || 'ESZ018R xlsx (재고현황 시트)'}</span>
            {importRows && (
              <span className="ml-auto text-xs text-gray-500">
                매칭 <span className="text-green-600 font-bold">{matched.length}</span>건 / 미매칭 {unmatched.length}건
              </span>
            )}
          </div>
          {importRows && (
            <>
              <div className="flex border-b border-yellow-200">
                {([['total','총수량(D열)'],['warehouse','창고별']] as ['total'|'warehouse',string][]).map(([id,lbl])=>(
                  <button key={id} onClick={() => { setImportSubTab(id); setImportApplied(false) }}
                    className={`px-4 py-2 text-sm font-medium border-b-2 -mb-px transition ${
                      importSubTab===id ? 'border-[#2E75B6] text-[#2E75B6]' : 'border-transparent text-gray-500 hover:text-gray-700'
                    }`}>{lbl}</button>
                ))}
              </div>
              {importSubTab === 'warehouse' && (
                <div className="flex items-center gap-3">
                  <span className="text-sm font-medium text-yellow-800">창고 선택</span>
                  <select value={selectedWh} onChange={e => { setSelectedWh(e.target.value); setImportApplied(false) }}
                    className="border border-yellow-300 rounded px-2 py-1 text-sm bg-white focus:outline-none">
                    {warehouses.map(wh => <option key={wh} value={wh}>{wh}</option>)}
                  </select>
                </div>
              )}
              <div className="overflow-x-auto max-h-52 overflow-y-auto rounded border border-yellow-100">
                <table className="text-xs w-full bg-white">
                  <thead className="sticky top-0 bg-yellow-100">
                    <tr>
                      <th className={thCls}>품목코드</th>
                      <th className={thCls}>품목명</th>
                      <th className={thCls}>파이</th>
                      <th className={thCls}>종류</th>
                      <th className={thCls}>구분</th>
                      <th className={thR}>{importSubTab==='total'?'총수량':selectedWh}</th>
                      <th className={thCls}>매칭</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {matched.map((r, i) => {
                      const qty = importSubTab==='total' ? r.total : (r.byWh[selectedWh]??0)
                      const [pai, type] = (r.metaKey ?? '').split('|')
                      return (
                        <tr key={i} className={i%2===0?'bg-white':'bg-gray-50'}>
                          <td className="px-3 py-1 font-mono text-gray-500">{r.code}</td>
                          <td className="px-3 py-1 max-w-[140px] truncate" title={r.name}>{r.name}</td>
                          <td className="px-3 py-1 font-mono text-gray-500">{pai ?? '—'}</td>
                          <td className="px-3 py-1 font-semibold">{type ?? '—'}{r.matchType==='housing' && r.compIdx > 0 ? <span className="text-purple-400 ml-1">[{r.compIdx+1}]</span> : ''}</td>
                          <td className="px-3 py-1">
                            <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${r.matchType==='cable'?'bg-blue-50 text-blue-700':r.matchType==='housing'?'bg-purple-50 text-purple-700':'bg-orange-50 text-orange-700'}`}>
                              {r.matchType==='cable'?'케이블':r.matchType==='housing'?'하우징':'페룰'}
                            </span>
                          </td>
                          <td className="px-3 py-1 text-right font-mono font-semibold">{qty.toLocaleString()}</td>
                          <td className="px-3 py-1 text-xs text-green-600">✓</td>
                        </tr>
                      )
                    })}
                    {unmatched.map((r, i) => {
                      const qty = importSubTab==='total' ? r.total : (r.byWh[selectedWh]??0)
                      return (
                        <tr key={`u-${i}`} className="opacity-40">
                          <td className="px-3 py-1 font-mono text-gray-500">{r.code}</td>
                          <td className="px-3 py-1 max-w-[140px] truncate" title={r.name}>{r.name}</td>
                          <td className="px-3 py-1 text-gray-400">—</td>
                          <td className="px-3 py-1 text-gray-400">—</td>
                          <td className="px-3 py-1 text-xs text-gray-400">—</td>
                          <td className="px-3 py-1 text-right font-mono text-gray-400">{qty.toLocaleString()}</td>
                          <td className="px-3 py-1 text-xs text-gray-400">미매칭</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
              <div className="flex items-center gap-3">
                <button onClick={() => applyErpImport(importSubTab==='total')}
                  className="px-5 py-2 bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold text-sm rounded-lg transition">
                  {importSubTab==='total' ? `총수량으로 현재고 적용 (${matched.length}건)` : `${selectedWh} 수량으로 현재고 적용 (${matched.length}건)`}
                </button>
                {importApplied && <span className="text-green-600 text-sm font-semibold">✅ 적용됐습니다. 저장 버튼으로 확정하세요.</span>}
              </div>
            </>
          )}
          <input ref={erpFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e => { const f = e.target.files?.[0]; if (f) { handleErpUpload(f); e.target.value = '' } }} />
        </div>
      )}

      {/* ── 필터 바 ─────────────────────────────────────────── */}
      <div className="flex items-center gap-4 flex-wrap text-xs bg-gray-50 border border-gray-200 rounded-lg px-4 py-2.5">
        <div className="flex items-center gap-1.5">
          <span className="text-gray-500 whitespace-nowrap">🔍 검색</span>
          <input
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            placeholder="파이·타입·품번·품명"
            className="border border-gray-200 rounded px-2 py-1 text-xs w-36 bg-white focus:outline-none focus:border-blue-400"
          />
        </div>
        <div className="flex items-center gap-1.5">
          <span className="text-gray-500 whitespace-nowrap">🙈 구형 숨김</span>
          <input
            value={hideKeyword}
            onChange={e => setHideKeyword(e.target.value)}
            placeholder="키워드 (예: 구형, 단종)"
            className="border border-gray-200 rounded px-2 py-1 text-xs w-44 bg-white focus:outline-none focus:border-orange-400"
          />
          {hideKeyword && (() => {
            const hiddenCount = tab === 'cable'
              ? cableKeys.length - filteredCableKeys.length
              : tab === 'housing'
              ? housingKeys.length - filteredHousingKeys.length
              : ferruleKeys.length - filteredFerruleKeys.length
            return hiddenCount > 0
              ? <span className="text-orange-500 font-semibold">{hiddenCount}건 숨김</span>
              : null
          })()}
        </div>
        {prevSnapshot && (
          <span className="text-gray-400 flex items-center gap-1">
            📅 전주 기준: <b className="text-gray-600">{prevSnapshot.date}</b>
            <span className="text-gray-300 ml-1">— ERP 업로드 적용 시 자동 갱신</span>
          </span>
        )}
        {!prevSnapshot && (
          <span className="text-gray-400 italic">ERP 재고 업로드 적용 시 전주 비교 컬럼이 표시됩니다</span>
        )}
        {(hideKeyword || searchQuery) && (
          <button onClick={() => { setHideKeyword(''); setSearchQuery('') }}
            className="text-gray-400 hover:text-gray-700 underline ml-auto">초기화</button>
        )}
      </div>

      {/* ── 통합 테이블 ─────────────────────────────────────── */}
      <div className="overflow-x-auto border border-gray-200 rounded-lg">
        <table className="text-xs w-full border-collapse bg-white">
          <thead>
            <tr>
              {tab !== 'ferrule' && <th className={thCls}>파이</th>}
              <th className={thCls}>{tab==='cable'?'케이블종류':tab==='housing'?'하우징타입':'페룰타입'}</th>
              {tab === 'housing' && <th className={thCls} style={{ width: 24 }}>#</th>}
              <th className={thCls}>품번</th>
              <th className={thCls}>품명</th>
              <th className={thCls}>구매처</th>
              <th className={thCls}>LT(일)</th>
              <th className={thR}>현재고{tab==='cable'?' (m)':' (EA)'}</th>
              <th className={thR}>기발주{tab==='cable'?' (m)':' (EA)'}</th>
              <th className={thCls}>상태</th>
              {CAN_WRITE && <th className={thCls}></th>}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {tab === 'housing' ? (
              housingDisplayRows.length > 0 ? housingDisplayRows.map(({ key, ci, pai, type, comp, inv, missing, isFirst, isLast, total }, rowIdx) => (
                <tr key={`${key}-${ci}`} className={missing ? 'bg-red-50' : !isFirst ? 'bg-purple-50' : rowIdx % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                  {/* 파이 — 첫 번째 부품 행만 표시 */}
                  <td className="px-3 py-1 font-mono text-gray-500 whitespace-nowrap">
                    {isFirst ? pai : ''}
                  </td>
                  {/* 하우징 타입 — 첫 번째 부품 행만 표시 */}
                  <td className="px-3 py-1.5 font-semibold whitespace-nowrap">
                    {isFirst ? type : <span className="text-purple-400 font-normal">└ 동일</span>}
                  </td>
                  {/* 부품 번호 */}
                  <td className="px-2 py-1 text-center text-gray-400 text-xs">{total > 1 ? ci + 1 : ''}</td>
                  {/* 품번, 품명, 구매처, LT */}
                  {(['품번','품명','구매처','리드타임'] as const).map(f => (
                    <td key={f} className="px-1 py-1">
                      <input
                        value={String(comp[f] ?? '')}
                        readOnly={!CAN_WRITE}
                        onChange={e => updateHousingCompMeta(key, ci, f, e.target.value)}
                        className={CAN_WRITE
                          ? `${inputCls} ${missing && f==='품번' ? 'border-red-400 bg-red-50' : ''}`
                          : inputRoCls}
                        placeholder={f==='리드타임' ? '숫자(일)' : f}
                      />
                    </td>
                  ))}
                  {/* 현재고 */}
                  <td className="px-1 py-1">
                    <input type="number" min="0" value={inv.현재고}
                      onChange={e => updateHousingCompInv(key, ci, '현재고', e.target.value)}
                      className={inputInvCls} />
                  </td>
                  {/* 기발주 */}
                  <td className="px-1 py-1">
                    <input type="number" min="0" value={inv.기발주 ?? 0}
                      onChange={e => updateHousingCompInv(key, ci, '기발주', e.target.value)}
                      className={inputInvCls} />
                  </td>
                  {/* 상태 */}
                  <td className="px-2 py-1 text-center">
                    {missing ? <span className="text-red-500 font-semibold">⚠</span> : <span className="text-green-600">✓</span>}
                  </td>
                  {/* 추가/제거 버튼 */}
                  {CAN_WRITE && (
                    <td className="px-1 py-1 whitespace-nowrap">
                      {isLast && (
                        <button
                          onClick={() => addHousingComp(key)}
                          title="부품 추가"
                          className="text-xs px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 hover:bg-purple-200 transition mr-1">
                          +
                        </button>
                      )}
                      {ci > 0 && (
                        <button
                          onClick={() => removeHousingComp(key, ci)}
                          title="부품 제거"
                          className="text-xs px-1.5 py-0.5 rounded bg-red-100 text-red-500 hover:bg-red-200 transition">
                          ×
                        </button>
                      )}
                    </td>
                  )}
                </tr>
              )) : (
                <tr>
                  <td colSpan={11} className="text-center text-gray-400 py-10">
                    STEP 1 실행 후 품번을 입력하세요.
                  </td>
                </tr>
              )
            ) : (tab==='cable' ? cableKeys : ferruleKeys).map((key, i) => {
              const [pai, type] = tab === 'ferrule' ? ['', key] : key.split('|')
              const m   = tab==='cable' ? getCm(key) : getFm(key)
              const inv = tab==='cable' ? getCiv(key) : getFiv(key)
              const missing = !m.품번
              return (
                <tr key={key} className={missing ? 'bg-red-50' : i%2===0 ? 'bg-white' : 'bg-gray-50'}>
                  {tab !== 'ferrule' && <td className="px-3 py-1 font-mono text-gray-500 whitespace-nowrap">{pai}</td>}
                  <td className="px-3 py-1.5 font-semibold whitespace-nowrap">{type}</td>
                  {(['품번','품명','구매처','리드타임'] as const).map(f => (
                    <td key={f} className="px-1 py-1">
                      <input
                        value={String(m[f] ?? '')}
                        readOnly={!CAN_WRITE}
                        onChange={e => tab==='cable' ? updateCableMeta(key, f, e.target.value) : updateFerruleMeta(key, f, e.target.value)}
                        className={CAN_WRITE
                          ? `${inputCls} ${missing && f==='품번' ? 'border-red-400 bg-red-50' : ''}`
                          : inputRoCls}
                        placeholder={f==='리드타임' ? '숫자(일)' : f}
                      />
                    </td>
                  ))}
                  <td className="px-1 py-1">
                    <div className="flex items-center gap-1">
                      <input type="number" min="0" value={inv.현재고}
                        onChange={e => tab==='cable' ? updateCableInv(key,'현재고',e.target.value) : updateFerruleInv(key,'현재고',e.target.value)}
                        className={inputInvCls} />
                      {prevSnapshot && (
                        <DeltaBadge
                          curr={inv.현재고}
                          prev={tab==='cable' ? (prevSnapshot.cable[key]?.현재고 ?? null) : (prevSnapshot.ferrule[key]?.현재고 ?? null)}
                        />
                      )}
                    </div>
                  </td>
                  <td className="px-1 py-1">
                    <input type="number" min="0" value={inv.기발주 ?? 0}
                      onChange={e => tab==='cable' ? updateCableInv(key,'기발주',e.target.value) : updateFerruleInv(key,'기발주',e.target.value)}
                      className={inputInvCls} />
                  </td>
                  <td className="px-2 py-1 text-center">
                    {missing ? <span className="text-red-500 font-semibold">⚠</span> : <span className="text-green-600">✓</span>}
                  </td>
                  {CAN_WRITE && (
                    <td className="px-1 py-1 text-center">
                      <button
                        onClick={() => tab === 'cable' ? deleteCableKey(key) : deleteFerruleKey(key)}
                        title="항목 삭제"
                        className="text-xs px-1.5 py-0.5 rounded bg-red-100 text-red-500 hover:bg-red-200 transition">
                        ×
                      </button>
                    </td>
                  )}
                </tr>
              )
            })}
            {tab !== 'housing' && (tab==='cable' ? cableKeys : ferruleKeys).length === 0 && (
              <tr>
                <td colSpan={tab === 'ferrule' ? (CAN_WRITE ? 9 : 8) : (CAN_WRITE ? 10 : 9)} className="text-center text-gray-400 py-10">
                  STEP 1 실행 후 품번을 입력하세요.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* ── 하우징 공용 부품 분석 ───────────────────────────── */}
      {tab === 'housing' && housingKeys.length > 0 && (() => {
        // 품번 → { 품명, 등록된 하우징 타입 목록 }
        const shared = new Map<string, { 품명: string; types: string[] }>()
        for (const key of housingKeys) {
          const [pai, type] = key.split('|')
          const comps = getHousingComps(key)
          for (const comp of comps) {
            const bn = (comp.품번 || '').trim()
            if (!bn) continue
            if (!shared.has(bn)) shared.set(bn, { 품명: comp.품명 ?? '', types: [] })
            shared.get(bn)!.types.push(`${pai} ${type}`)
          }
        }
        const multiRows = [...shared.entries()]
          .filter(([, v]) => v.types.length > 1)
          .sort((a, b) => b[1].types.length - a[1].types.length)
        if (!multiRows.length) return null
        return (
          <div className="border border-purple-200 rounded-lg bg-purple-50 p-4 space-y-2">
            <div className="text-sm font-semibold text-purple-800">
              🔗 공용 부품 현황 — 여러 하우징 타입에 동일 품번이 등록된 부품
              <span className="ml-2 text-xs font-normal text-purple-600">(발주 집계 시 해당 타입 수요 합산됨)</span>
            </div>
            <div className="overflow-x-auto rounded border border-purple-100">
              <table className="text-xs w-full bg-white">
                <thead>
                  <tr className="bg-purple-100">
                    <th className={thCls}>품번</th>
                    <th className={thCls}>품명</th>
                    <th className="px-3 py-2 text-xs font-bold text-gray-700 border-b-2 border-gray-200 bg-purple-100 text-center">타입 수</th>
                    <th className={thCls}>등록된 하우징 타입</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {multiRows.map(([bn, v]) => (
                    <tr key={bn} className="hover:bg-purple-50">
                      <td className="px-3 py-1.5 font-mono text-gray-600">{bn}</td>
                      <td className="px-3 py-1.5 text-gray-800">{v.품명 || '—'}</td>
                      <td className="px-3 py-1.5 text-center">
                        <span className="bg-purple-200 text-purple-800 font-bold px-2 py-0.5 rounded-full">{v.types.length}</span>
                      </td>
                      <td className="px-3 py-1.5 text-gray-600">{v.types.join(', ')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )
      })()}

      {/* ── 저장 ────────────────────────────────────────────── */}
      <div className="flex items-center gap-3">
        <button onClick={saveAll}
          className="bg-[#FF4B4B] hover:bg-[#e03030] text-white font-bold px-6 py-2 rounded-lg transition">
          💾 저장
        </button>
        {saved && (
          <span className="text-green-600 text-sm font-semibold">✅ 저장됐습니다.</span>
        )}
      </div>
    </div>
  )
}
