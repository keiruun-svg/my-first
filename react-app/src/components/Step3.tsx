import { useState } from 'react'
import ExcelJS from 'exceljs'
import { aggregateStats, buildOrderPlan } from '../lib/step2Core'
import type { Metadata, Inventory, SalesAnalysis, AppSettings } from '../lib/types'
import FileUploader from './FileUploader'

interface Props {
  metadata: Metadata
  inventory: Inventory
  sales: SalesAnalysis
  settings: AppSettings
}

const MONTHS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']

export default function Step3({ metadata, inventory, sales, settings }: Props) {
  const [logs, setLogs] = useState<string[]>([])
  const [running, setRunning] = useState(false)
  const [done, setDone] = useState(false)
  const [gaongFile, setGaongFile] = useState<File | null>(null)

  const nCableInv = Object.values(inventory.cable).filter(v => (v?.현재고 ?? 0) > 0).length
  const nHousingInv = Object.values(inventory.housing).filter(v => {
    const items = Array.isArray(v) ? v : [v]
    return items.some(i => (i?.현재고 ?? 0) > 0 || (i?.기발주 ?? 0) > 0)
  }).length
  const nSales = Object.values(sales).filter(v =>
    Object.keys(v).filter(k => /^\d{2}$/.test(k)).some(yr => ((v[yr] as { sales: number } | undefined)?.sales ?? 0) > 0)
  ).length

  const run = async () => {
    if (!gaongFile) return alert('가공파일을 선택해주세요.')
    setRunning(true); setDone(false); setLogs([])
    try {
      const buf = await gaongFile.arrayBuffer()
      const stats = aggregateStats(buf)
      setLogs(stats.logs)

      const rows = buildOrderPlan(stats, metadata, inventory, sales, settings.lead_time_default)
      const wb = new ExcelJS.Workbook()
      wb.creator = 'AJW 발주계획 시스템'

      const mainColor = settings.colors.main_header
      const headerFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF' + mainColor } }
      const border: Partial<ExcelJS.Border> = { style: 'thin' as const, color: { argb: 'FF000000' } }
      const allBorders = { top: border, bottom: border, left: border, right: border }

      // ── 케이블 사용내역 시트 ──
      const wsc = wb.addWorksheet('케이블 사용내역')
      wsc.mergeCells('A1:W1')
      const title = wsc.getCell('A1')
      title.value = '2026 연간 발주 계획 — 케이블 사용내역'
      title.fill = headerFill
      title.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } }
      title.alignment = { horizontal: 'center', vertical: 'middle' }
      wsc.getRow(1).height = 26

      wsc.getRow(2).height = 18
      const subHeaders = [
        ['A2:G2','기본 정보','374151'],
        ['H2:I2','2023년',settings.colors.year_23],
        ['J2:K2','2024년',settings.colors.year_24],
        ['L2:M2','2025년',settings.colors.year_25],
        ['N2:Q2','📊 트렌드 분석','375623'],
        ['R2:R2','⚠ 안전재고','C00000'],
        ['S2:T2','재고 현황','7030A0'],
        ['U2:V2','✏ 2026 발주 계획','C55A11'],
        ['W2:W2','비고','595959'],
      ]
      for (const [rng, lbl, color] of subHeaders) {
        if (rng.includes(':')) wsc.mergeCells(rng)
        const cell = wsc.getCell(rng.split(':')[0])
        cell.value = lbl
        cell.fill = { type:'pattern', pattern:'solid', fgColor: { argb:'FF'+color } }
        cell.font = { name:'Arial', bold:true, size:9, color:{ argb:'FFFFFFFF' } }
        cell.alignment = { horizontal:'center', vertical:'middle' }
        cell.border = allBorders
      }

      wsc.getRow(3).height = 42
      const hdrs = ['NO','파이','케이블종류','품번','품명','구매처','리드타임\n(일)',
        '연간(m)','피크(m)','연간(m)','피크(m)','연간(m)','피크(m)',
        '3개년\n평균연간','3개년\n피크평균','23→24\n증감률','24→25\n증감률',
        '안전재고\n(m)','현재고\n(m)','기발주\n(참고)','2026목표\n(m)','필요발주\n(m)','비고']
      const colWidths = [5,8,20,16,38,14,9,12,12,12,12,12,12,13,13,11,11,11,11,11,13,13,20]
      hdrs.forEach((h, i) => {
        const cell = wsc.getCell(3, i + 1)
        cell.value = h
        cell.fill = { type:'pattern', pattern:'solid', fgColor: { argb:'FFBDD7EE' } }
        cell.font = { name:'Arial', bold:true, size:9, color:{ argb:'FFFFFFFF' } }
        cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true }
        cell.border = allBorders
        wsc.getColumn(i + 1).width = colWidths[i]
      })

      let ri = 4; let no = 1
      const cableRows = rows.filter(r => r.type === 'cable')
      for (const row of cableRows) {
        const rf = ri % 2 === 0 ? { type:'pattern' as const, pattern:'solid' as const, fgColor:{ argb:'FFF5F5F5' } } : undefined
        const vals = [no, row.pai, row.ctype, row.품번, row.품명, row.구매처, row.리드타임,
          row.yearStats['23']?.annual||null, row.yearStats['23']?.peak||null,
          row.yearStats['24']?.annual||null, row.yearStats['24']?.peak||null,
          row.yearStats['25']?.annual||null, row.yearStats['25']?.peak||null,
          { formula: `=ROUND(AVERAGE(H${ri},J${ri},L${ri}),0)` },
          { formula: `=ROUND(AVERAGE(I${ri},K${ri},M${ri}),0)` },
          { formula: `=IFERROR((J${ri}-H${ri})/H${ri},"")` },
          { formula: `=IFERROR((L${ri}-J${ri})/J${ri},"")` },
          { formula: `=ROUND(O${ri}*G${ri}/30,0)` },
          row.현재고 || null, row.기발주 || null, null,
          { formula: `=IFERROR(U${ri}-IFERROR(S${ri},0)-IFERROR(T${ri},0),"")` }, ''
        ]
        vals.forEach((v, ci) => {
          const cell = wsc.getCell(ri, ci + 1)
          if (typeof v === 'object' && v !== null && 'formula' in v) cell.value = v as ExcelJS.CellFormulaValue
          else cell.value = v as ExcelJS.CellValue
          cell.border = allBorders
          if (rf) cell.fill = rf
          cell.font = { name:'Arial', size:9 }
          if (ci >= 7) { cell.numFmt = '#,##0'; cell.alignment = { horizontal:'right', vertical:'middle' } }
          else if (ci === 15 || ci === 16) { cell.numFmt = '0.0%;[Red]-0.0%'; cell.alignment = { horizontal:'center', vertical:'middle' } }
          else { cell.alignment = { horizontal: ci === 0 ? 'center' : 'left', vertical:'middle' } }
        })
        wsc.getCell(ri, 18).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF2CC' } }
        wsc.getCell(ri, 18).font = { name:'Arial', bold:true, size:9 }
        const targetCell = wsc.getCell(ri, 21)
        targetCell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFC0' } }
        targetCell.font = { name:'Arial', bold:true, size:9, color:{ argb:'FF0000FF' } }
        targetCell.border = { top:{style:'medium',color:{argb:'FFC55A11'}}, bottom:{style:'medium',color:{argb:'FFC55A11'}}, left:{style:'medium',color:{argb:'FFC55A11'}}, right:{style:'medium',color:{argb:'FFC55A11'}} }
        wsc.getCell(ri, 22).font = { name:'Arial', bold:true, size:9, color:{ argb:'FFC00000' } }
        wsc.getRow(ri).height = 17
        ri++; no++
      }
      wsc.views = [{ state: 'frozen', xSplit: 3, ySplit: 3 }]

      // ── 하우징 사용내역 시트 ──
      const wsh = wb.addWorksheet('하우징 사용내역')
      wsh.mergeCells('A1:W1')
      const titleH = wsh.getCell('A1')
      titleH.value = '2026 연간 발주 계획 — 하우징 사용내역'
      titleH.fill = headerFill
      titleH.font = { name:'Arial', bold:true, size:13, color:{ argb:'FFFFFFFF' } }
      titleH.alignment = { horizontal:'center', vertical:'middle' }
      wsh.getRow(1).height = 26

      ri = 4; no = 1
      const housingRows = rows.filter(r => r.type === 'housing')
      for (const [i, col] of hdrs.entries()) {
        const cell = wsh.getCell(3, i + 1)
        cell.value = col
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFBDD7EE' } }
        cell.font = { name:'Arial', bold:true, size:9, color:{ argb:'FFFFFFFF' } }
        cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true }
        cell.border = allBorders
        wsh.getColumn(i + 1).width = colWidths[i]
      }
      for (const row of housingRows) {
        const rf = ri % 2 === 0 ? { type:'pattern' as const, pattern:'solid' as const, fgColor:{ argb:'FFF5F5F5' } } : undefined
        const vals = [no, row.pai, row.ctype, row.품번, row.품명, row.구매처, row.리드타임,
          row.yearStats['23']?.annual||null, row.yearStats['23']?.peak||null,
          row.yearStats['24']?.annual||null, row.yearStats['24']?.peak||null,
          row.yearStats['25']?.annual||null, row.yearStats['25']?.peak||null,
          { formula: `=ROUND(AVERAGE(H${ri},J${ri},L${ri}),0)` },
          { formula: `=ROUND(AVERAGE(I${ri},K${ri},M${ri}),0)` },
          { formula: `=IFERROR((J${ri}-H${ri})/H${ri},"")` },
          { formula: `=IFERROR((L${ri}-J${ri})/J${ri},"")` },
          { formula: `=ROUND(O${ri}*G${ri}/30,0)` },
          row.현재고||null, row.기발주||null, null,
          { formula: `=IFERROR(U${ri}-IFERROR(S${ri},0)-IFERROR(T${ri},0),"")` }, ''
        ]
        vals.forEach((v, ci) => {
          const cell = wsh.getCell(ri, ci + 1)
          if (typeof v === 'object' && v !== null && 'formula' in v) cell.value = v as ExcelJS.CellFormulaValue
          else cell.value = v as ExcelJS.CellValue
          cell.border = allBorders
          if (rf) cell.fill = rf
          cell.font = { name:'Arial', size:9 }
          if (ci >= 7) { cell.numFmt = '#,##0'; cell.alignment = { horizontal:'right', vertical:'middle' } }
          else if (ci === 15 || ci === 16) { cell.numFmt = '0.0%;[Red]-0.0%'; cell.alignment = { horizontal:'center', vertical:'middle' } }
          else { cell.alignment = { horizontal: ci === 0 ? 'center' : 'left', vertical:'middle' } }
        })
        wsh.getCell(ri, 18).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF2CC' } }
        wsh.getCell(ri, 18).font = { name:'Arial', bold:true, size:9 }
        const tc = wsh.getCell(ri, 21)
        tc.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFC0' } }
        tc.font = { name:'Arial', bold:true, size:9, color:{ argb:'FF0000FF' } }
        tc.border = { top:{style:'medium',color:{argb:'FFC55A11'}}, bottom:{style:'medium',color:{argb:'FFC55A11'}}, left:{style:'medium',color:{argb:'FFC55A11'}}, right:{style:'medium',color:{argb:'FFC55A11'}} }
        wsh.getCell(ri, 22).font = { name:'Arial', bold:true, size:9, color:{ argb:'FFC00000' } }
        wsh.getRow(ri).height = 17
        ri++; no++
      }
      wsh.views = [{ state: 'frozen', xSplit: 3, ySplit: 3 }]

      // ── 월별 발주계획 시트 ──
      const wsm = wb.addWorksheet('2026 월별 발주계획')
      wsm.mergeCells(`A1:T1`)
      const titleM = wsm.getCell('A1')
      titleM.value = '2026 월별 발주 계획 (과거 계절 패턴 기반 자동 분배)'
      titleM.fill = headerFill
      titleM.font = { name:'Arial', bold:true, size:13, color:{ argb:'FFFFFFFF' } }
      titleM.alignment = { horizontal:'center', vertical:'middle' }
      wsm.getRow(1).height = 26

      const mHdrs = ['NO','분류','파이','종류','품번','단위','연간목표',...MONTHS,'합계검증']
      const mWidths = [5,10,8,20,16,6,14,...new Array(12).fill(9),10]
      mHdrs.forEach((h, i) => {
        const cell = wsm.getCell(3, i + 1)
        cell.value = h
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFBDD7EE' } }
        cell.font = { name:'Arial', bold:true, size:9, color:{ argb:'FFFFFFFF' } }
        cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true }
        cell.border = allBorders
        wsm.getColumn(i + 1).width = mWidths[i]
      })
      wsm.getRow(3).height = 40

      let mri = 4; let midx = 1
      for (const row of rows) {
        const rf = mri % 2 === 0 ? { type:'pattern' as const, pattern:'solid' as const, fgColor:{ argb:'FFF5F5F5' } } : undefined
        const combined = new Array(12).fill(0)
        let total = 0
        for (const yr of stats.years) {
          const m = row.yearStats[yr]?.monthly ?? new Array(12).fill(0)
          m.forEach((v, i) => { combined[i] += v; total += v })
        }
        const ratios = combined.map(v => total > 0 ? Math.round(v / total * 1000000) / 1000000 : Math.round(1/12 * 1000000) / 1000000)
        const unit = row.type === 'cable' ? 'm' : 'EA'
        const baseVals = [midx, row.type === 'cable' ? '케이블' : '하우징', row.pai, row.ctype, row.품번, unit]
        baseVals.forEach((v, ci) => {
          const cell = wsm.getCell(mri, ci + 1)
          cell.value = v as ExcelJS.CellValue
          cell.font = { name:'Arial', size:9 }; cell.border = allBorders
          cell.alignment = { horizontal: ci <= 2 ? 'center' : 'left', vertical:'middle' }
          if (rf) cell.fill = rf
        })
        const tgtCell = wsm.getCell(mri, 7)
        tgtCell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFC0' } }
        tgtCell.font = { name:'Arial', bold:true, size:9, color:{ argb:'FF0000FF' } }
        tgtCell.border = { top:{style:'medium',color:{argb:'FFC55A11'}}, bottom:{style:'medium',color:{argb:'FFC55A11'}}, left:{style:'medium',color:{argb:'FFC55A11'}}, right:{style:'medium',color:{argb:'FFC55A11'}} }
        tgtCell.numFmt = '#,##0'; tgtCell.alignment = { horizontal:'right', vertical:'middle' }
        ratios.forEach((ratio, mi) => {
          const cell = wsm.getCell(mri, 8 + mi)
          cell.value = { formula: `=IFERROR(ROUND($G${mri}*${ratio},0),"")` } as ExcelJS.CellFormulaValue
          cell.font = { name:'Arial', size:9 }; cell.border = allBorders
          cell.numFmt = '#,##0'; cell.alignment = { horizontal:'right', vertical:'middle' }
          if (rf) cell.fill = rf
        })
        const sumCell = wsm.getCell(mri, 20)
        sumCell.value = { formula: `=IFERROR(SUM(H${mri}:S${mri}),"")` } as ExcelJS.CellFormulaValue
        sumCell.font = { name:'Arial', size:9, bold:true, color:{ argb:'FF375623' } }
        sumCell.border = allBorders; sumCell.numFmt = '#,##0'; sumCell.alignment = { horizontal:'right', vertical:'middle' }
        if (rf) sumCell.fill = rf
        wsm.getRow(mri).height = 17; mri++; midx++
      }
      wsm.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]

      // ── 이상항목 검토 시트 ──
      const wsa = wb.addWorksheet('⚠ 이상항목 검토')
      wsa.mergeCells('A1:D1')
      const titleA = wsa.getCell('A1')
      titleA.value = '데이터 이상 항목 검토 (자동 분석)'
      titleA.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFC00000' } }
      titleA.font = { name:'Arial', bold:true, size:13, color:{ argb:'FFFFFFFF' } }
      titleA.alignment = { horizontal:'center', vertical:'middle' }
      wsa.getRow(1).height = 26
      wsa.getRow(3).height = 30
      const aHdrs = ['구분','항목','품번','내용 및 조치 권고']
      const aWidths = [10, 22, 16, 65]
      aHdrs.forEach((h, i) => {
        const cell = wsa.getCell(3, i + 1)
        cell.value = h
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFBDD7EE' } }
        cell.font = { name:'Arial', bold:true, size:9, color:{ argb:'FFFFFFFF' } }
        cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true }
        cell.border = allBorders
        wsa.getColumn(i + 1).width = aWidths[i]
      })
      const anomalies: [string, string, string, string, string][] = [
        ['주의','2.0mm 자켓 피그테일','(확인필요)','참고파일에 품번 미등재. 사용 케이블 품번 확인 필요.','orange'],
        ['정보','OM4 피그테일 케이블','P14-RM-417K','3개년 사용 없음. 재고 보유. 단종 검토 필요.','red'],
        ['정보','OM3 피그테일 케이블','P14-RM-417H','23년 648m → 24~25년 0m. 미사용 추세.','orange'],
        ['정보','피그테일 전체','(전 색상)','23년 대비 25년 약 81% 급감. 2026 목표량 보수적 설정 권고.','blue'],
      ]
      const anomalyColors: Record<string, string> = { red:'FFD7D7', orange:'FFE6C8', blue:'D7E8FF' }
      anomalies.forEach(([type, item, bunho, desc, color], idx) => {
        const ari = 4 + idx
        const rf2 = { type:'pattern' as const, pattern:'solid' as const, fgColor:{ argb:'FF' + (anomalyColors[color] ?? 'FFFFFF') } }
        ;[type, item, bunho, desc].forEach((v, ci) => {
          const cell = wsa.getCell(ari, ci + 1)
          cell.value = v
          cell.fill = rf2
          cell.font = { name:'Arial', size:9, bold: ci === 0 }
          cell.border = allBorders
          cell.alignment = ci === 0 ? { horizontal:'center', vertical:'middle' }
            : ci === 3 ? { horizontal:'left', vertical:'middle', wrapText:true }
            : { horizontal:'left', vertical:'middle' }
        })
        wsa.getRow(ari).height = 36
      })
      wsa.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]

      const arrayBuf = await wb.xlsx.writeBuffer()
      const blob = new Blob([arrayBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a'); a.href = url
      a.download = `연간발주계획_${new Date().toISOString().slice(0,10)}.xlsx`
      a.click(); URL.revokeObjectURL(url)
      setLogs(l => [...l, `✅ STEP 3 완료 — 케이블 ${cableRows.length}행 / 하우징 ${housingRows.length}행`])
      setDone(true)
    } catch (e) {
      setLogs(l => [...l, `❌ 오류: ${e}`])
    } finally {
      setRunning(false)
    }
  }

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>STEP 3 — 발주계획 생성</b>: STEP 1에서 생성한 <b>가공파일</b>을 업로드하면
        연간발주계획.xlsx를 생성합니다.
        현재고·기발주는 <b>📦 재고 현황</b> 탭, 수요 기반 분석은 <b>📈 STEP 2</b> 탭에서 사전 실행하세요.
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-start">
        <div className="md:col-span-2">
          <FileUploader
            label="가공파일.xlsx (STEP 1 결과)"
            fileName={gaongFile?.name ?? ''}
            onFile={setGaongFile}
          />
        </div>

        <div className="bg-[#e8f4fd] rounded-lg p-4 text-sm text-gray-700 space-y-2">
          <div className="font-semibold mb-1">재고 현황 탭 입력값</div>
          <div>케이블 <b>{nCableInv}</b>항목 / 하우징 <b>{nHousingInv}</b>항목</div>
          <hr className="border-gray-300" />
          {nSales > 0 ? (
            <div className="text-[#1a6a2a] font-semibold">✅ 판매 분석 <b>{nSales}</b>개 품목 포함</div>
          ) : (
            <div className="text-[#856404]">⚠ 판매 분석 없음 — STEP 2 먼저 실행하세요.</div>
          )}
        </div>
      </div>

      <hr className="border-gray-200" />

      <div className="flex gap-2">
        <button
          onClick={run}
          disabled={running || !gaongFile}
          className="flex-1 bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold py-2 px-4 rounded transition text-sm"
        >
          {running ? '⏳ 생성 중...' : '▶ STEP 3 실행 — 발주계획 생성'}
        </button>
        {(done || logs.length > 0) && (
          <button
            onClick={() => { setDone(false); setLogs([]); setGaongFile(null) }}
            className="px-4 py-2 text-sm border border-gray-300 rounded text-gray-600 hover:bg-gray-100 transition"
          >
            🗑 초기화
          </button>
        )}
      </div>

      {done && (
        <div className="bg-[#D6F0D8] px-4 py-3 rounded-md text-sm font-semibold text-[#1a6a2a]">
          ✅ STEP 3 완료 — 발주계획 생성! 💡 노란색 셀(2026 목표 발주량)에 목표량을 입력하면 필요 발주량이 자동 계산됩니다.
        </div>
      )}

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs space-y-0.5 max-h-48 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}
    </div>
  )
}
