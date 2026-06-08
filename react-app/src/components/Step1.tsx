import { useState, useEffect, useRef } from 'react'
import { parseERP, detectYears } from '../lib/parse/parseERP'
import { buildPivot } from '../lib/aggregate/pivot'
import { writeGaong } from '../lib/output/writeGaong'
import { runStep1 } from '../lib/step1Core'
import { saveMetadata } from '../lib/supabase'
import { downloadXlsx, today } from '../lib/download'
import type { Metadata, AppSettings } from '../lib/types'
import FileUploader from './FileUploader'

interface Props {
  metadata:    Metadata
  setMetadata: (m: Metadata) => void
  settings?:   AppSettings
}

export default function Step1({ metadata, setMetadata, settings }: Props) {
  const [logs, setLogs]               = useState<string[]>([])
  const [running, setRunning]         = useState(false)
  const [done, setDone]               = useState(false)
  const [erpFile, setErpFile]         = useState<File | null>(null)
  const [availYears, setAvailYears]   = useState<string[]>([])
  const [selectedYears, setSelectedYears] = useState<string[]>([])
  const bufRef = useRef<ArrayBuffer | null>(null)   // 한 번 읽은 버퍼 재사용

  const nCable   = Object.keys(metadata.cable).length
  const nHousing = Object.keys(metadata.housing).length
  const ltDefault = settings?.lead_time_default ?? 60

  // 파일 선택 시 연도 자동 감지
  useEffect(() => {
    if (!erpFile) { setAvailYears([]); setSelectedYears([]); bufRef.current = null; return }
    erpFile.arrayBuffer().then(buf => {
      bufRef.current = buf
      try {
        const years = detectYears(buf)
        setAvailYears(years)
        setSelectedYears(years)  // 기본값: 전체 선택
      } catch {
        setAvailYears([])
      }
    })
  }, [erpFile])

  const toggleYear = (yr: string) =>
    setSelectedYears(prev =>
      prev.includes(yr) ? prev.filter(y => y !== yr) : [...prev, yr].sort()
    )

  const run = async () => {
    if (!erpFile || !bufRef.current) return alert('ERP 파일을 선택해주세요.')
    if (!selectedYears.length) return alert('포함할 연도를 하나 이상 선택해주세요.')
    setRunning(true); setDone(false); setLogs([])
    try {
      const buf = bufRef.current
      const parseLogs: string[] = []

      // A: 파싱 → 피벗 → 가공파일 생성
      const rows     = parseERP(buf, parseLogs)
      const pivot    = buildPivot(rows, selectedYears)
      const gaongBuf = writeGaong(pivot)

      // B: 가공파일 다운로드
      downloadXlsx(gaongBuf, `가공파일_${today()}.xlsx`)

      // C: 신규 품번 타입 감지 → 품번 관리 업데이트
      const result  = runStep1(gaongBuf, metadata)
      const allLogs = [...parseLogs, ...result.logs]
      setLogs(allLogs)

      if (result.newCableKeys.length || result.newHousingKeys.length) {
        const newMeta: Metadata = {
          cable:   { ...metadata.cable },
          housing: { ...metadata.housing },
          ferrule: { ...metadata.ferrule },
        }
        for (const k of result.newCableKeys)
          newMeta.cable[k] = { 품번: '', 품명: '', 구매처: '', 리드타임: null }
        for (const k of result.newHousingKeys)
          newMeta.housing[k] = { 품번: '', 품명: '', 구매처: '', 리드타임: null }
        saveMetadata(newMeta)
        setMetadata(newMeta)
        setLogs(l => [...l, `ℹ️ 신규 타입 ${result.newCableKeys.length + result.newHousingKeys.length}건 품번 관리에 추가됨 — 품번을 입력해주세요.`])
      }
      setDone(true)
    } catch (e) {
      const stack = e instanceof Error ? e.stack ?? '' : ''
      setLogs([`❌ 오류: ${e}`, ...(stack ? [stack] : [])])
    } finally {
      setRunning(false)
    }
  }

  const reset = () => { setDone(false); setLogs([]); setErpFile(null) }

  return (
    <div className="space-y-4">
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>STEP 1 — ERP 파일 가공</b>: 맥산 ERP에서 추출한 <b>구매조회</b> 또는 <b>구매현황</b> 파일을
        업로드하면 <b>가공파일.xlsx</b>를 자동 다운로드하고, 품번 관리를 업데이트합니다.
        두 형식 모두 지원하며 자동으로 감지합니다.
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-start">
        <div className="md:col-span-2 space-y-3">
          <FileUploader
            label="구매조회 / 구매현황 파일 (ERP 원본)"
            fileName={erpFile?.name ?? ''}
            onFile={setErpFile}
          />

          {/* 연도 선택 — 파일 선택 후 표시 */}
          {availYears.length > 0 && (
            <div className="bg-white border border-gray-200 rounded-lg px-4 py-3">
              <div className="text-sm font-semibold text-gray-700 mb-2">포함할 연도 선택</div>
              <div className="flex flex-wrap gap-2">
                {availYears.map(yr => {
                  const checked = selectedYears.includes(yr)
                  return (
                    <button
                      key={yr}
                      onClick={() => toggleYear(yr)}
                      className={`px-3 py-1 rounded-full text-sm font-medium border transition
                        ${checked
                          ? 'bg-[#2E75B6] text-white border-[#2E75B6]'
                          : 'bg-white text-gray-500 border-gray-300 hover:border-[#2E75B6]'
                        }`}
                    >
                      {yr}년
                    </button>
                  )
                })}
              </div>
              {selectedYears.length === 0 && (
                <p className="text-xs text-red-500 mt-1">연도를 하나 이상 선택해주세요.</p>
              )}
            </div>
          )}
        </div>

        <div className="bg-[#e8f4fd] rounded-lg p-4 text-sm text-gray-700">
          <div className="font-semibold mb-2">저장된 품번</div>
          <div>케이블 <b>{nCable}</b> 타입 / 하우징 <b>{nHousing}</b> 타입</div>
          <hr className="border-gray-300 my-2" />
          <div>리드타임 기본값: <b>{ltDefault}일</b></div>
        </div>
      </div>

      <hr className="border-gray-200" />

      <div className="flex gap-2">
        <button
          onClick={run}
          disabled={running || !erpFile || selectedYears.length === 0}
          className="flex-1 bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold py-2 px-4 rounded transition text-sm"
        >
          {running
            ? '⏳ 처리 중...'
            : `▶ STEP 1 실행 — ERP 파일 가공 & 사용내역 생성${selectedYears.length ? ` (${selectedYears.map(y => y + '년').join(', ')})` : ''}`
          }
        </button>
        {(done || logs.length > 0) && (
          <button onClick={reset}
            className="px-4 py-2 text-sm border border-gray-300 rounded text-gray-600 hover:bg-gray-100 transition">
            🗑 초기화
          </button>
        )}
      </div>

      {done && (
        <div className="bg-green-50 border border-green-200 rounded-lg px-4 py-3 text-sm font-semibold text-green-700">
          ✅ 처리 완료! 품번 관리 탭에서 신규 타입의 품번·품명·구매처·리드타임을 입력해주세요.
        </div>
      )}

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs leading-5 max-h-48 overflow-y-auto whitespace-pre-wrap">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      <hr className="border-gray-200" />

      <div className="bg-gray-50 border border-gray-200 rounded-lg p-4">
        <div className="text-sm font-semibold text-gray-700 mb-3">📋 전체 워크플로우</div>
        <ol className="text-sm text-gray-600 space-y-1.5 list-none">
          {[
            ['STEP 1', '맥산 ERP 구매조회/구매현황 파일 업로드 → 가공파일.xlsx 자동 다운로드'],
            ['품번 관리', '신규 케이블·하우징 타입에 품번·품명·구매처·리드타임 입력 (로컬 전용)'],
            ['재고 현황', '현재고·기발주 수량 입력 후 저장'],
            ['STEP 2', '(선택) 전체 판매량 + 맥산 생산량 업로드 → 판매 분석 Excel 저장'],
            ['STEP 3', '가공파일 + (선택) 판매분석 파일 업로드 → 발주계획 Excel 다운로드'],
            ['품번 생성기', 'OJC 품목명·규격 입력 시 품번 자동 생성 (A14-xxx-xxx-xxx-xxxX 형식)'],
          ].map(([step, desc], i) => (
            <li key={i} className="flex gap-2">
              <span className="text-xs font-bold text-[#2E75B6] bg-[#e8f4fd] rounded px-1.5 py-0.5 whitespace-nowrap flex-shrink-0">{step}</span>
              <span>{desc}</span>
            </li>
          ))}
        </ol>
      </div>
    </div>
  )
}
