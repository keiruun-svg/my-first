import { useState, useRef } from 'react'
import { runStep1 } from '../lib/step1Core'
import { saveMetadata } from '../lib/supabase'
import type { Metadata, AppSettings } from '../lib/types'

interface Props {
  metadata: Metadata
  setMetadata: (m: Metadata) => void
  settings?: AppSettings
}

export default function Step1({ metadata, setMetadata, settings }: Props) {
  const [logs, setLogs] = useState<string[]>([])
  const [running, setRunning] = useState(false)
  const [done, setDone] = useState(false)
  const [fileName, setFileName] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)

  const nCable = Object.keys(metadata.cable).length
  const nHousing = Object.keys(metadata.housing).length
  const ltDefault = settings?.lead_time_default ?? 60

  const run = async () => {
    const file = fileRef.current?.files?.[0]
    if (!file) return alert('ERP 파일을 선택해주세요.')
    setRunning(true); setDone(false); setLogs([])
    try {
      const buf = await file.arrayBuffer()
      const result = runStep1(buf, metadata)
      setLogs(result.logs)
      if (result.newCableKeys.length || result.newHousingKeys.length) {
        const newMeta: Metadata = {
          cable: { ...metadata.cable },
          housing: { ...metadata.housing },
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
      setLogs([`❌ 오류: ${e}`])
    } finally {
      setRunning(false)
    }
  }

  return (
    <div className="space-y-4">
      {/* Step-box */}
      <div className="bg-[#E0F2FE] border-l-4 border-[#0EA5E9] px-4 py-3 rounded text-sm">
        <b>STEP 1 — ERP 파일 가공</b>: 맥산 ERP에서 추출한 <b>구매조회</b> 또는 <b>구매현황</b> 파일을
        업로드하면 생산자재_사용내역.xlsx를 자동 생성합니다.
        두 형식 모두 지원하며 자동으로 감지합니다. 품번·품명·구매처·리드타임은 <b>📋 품번 관리</b> 탭 정보로 자동 채워집니다.
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-start">
        {/* 파일 업로더 — Streamlit 스타일 */}
        <div className="md:col-span-2 space-y-1.5">
          <label className="block text-sm font-semibold text-gray-700">
            구매조회 / 구매현황 파일 (ERP 원본)
          </label>
          <div className="bg-white border border-gray-300 rounded-lg px-4 py-3 flex items-center gap-3">
            <label className="flex items-center gap-1.5 bg-white border border-gray-400 rounded px-3 py-1.5 cursor-pointer hover:bg-gray-50 transition text-sm font-medium text-gray-700 shrink-0">
              <span className="text-base">↑</span> Upload
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={e => setFileName(e.target.files?.[0]?.name ?? '')}
              />
            </label>
            <span className="text-sm text-gray-500 truncate">
              {fileName ? <span className="text-green-600 font-medium">✓ {fileName}</span> : '200MB per file • XLSX'}
            </span>
          </div>
        </div>

        {/* 상태 정보 */}
        <div className="bg-sky-50 border border-sky-200 rounded-lg p-4 text-sm space-y-1">
          <div className="font-semibold text-sky-800 mb-2">저장된 품번</div>
          <div>케이블 <b>{nCable}</b> 타입 / 하우징 <b>{nHousing}</b> 타입</div>
          <div className="border-t border-sky-200 pt-2 mt-2 text-sky-700">
            리드타임 기본값: <b>{ltDefault}일</b>
          </div>
        </div>
      </div>

      <hr className="border-gray-200" />

      <div className="flex gap-2">
        <button
          onClick={run}
          disabled={running || !fileName}
          className="bg-[#0EA5E9] hover:bg-[#0284C7] disabled:bg-gray-300 text-white font-bold px-5 py-2 rounded transition text-sm"
        >
          {running ? '⏳ 처리 중...' : '▶ STEP 1 실행 — ERP 파일 가공 & 사용내역 생성'}
        </button>
        {(done || logs.length > 0) && (
          <button
            onClick={() => { setDone(false); setLogs([]); setFileName(''); if (fileRef.current) fileRef.current.value = '' }}
            className="px-3 py-2 text-sm border border-gray-300 rounded text-gray-600 hover:bg-gray-100 transition"
          >
            🗑 초기화
          </button>
        )}
      </div>

      {done && (
        <div className="bg-[#e8f5e9] border-l-4 border-[#1a7a3c] px-4 py-2 rounded text-sm font-semibold text-[#1a7a3c]">
          ✅ 처리 완료! 품번 관리 탭에서 신규 타입의 품번·품명·구매처·리드타임을 입력해주세요.
        </div>
      )}

      {logs.length > 0 && (
        <div className="bg-[#1e1e1e] text-[#d4d4d4] rounded p-3 font-mono text-xs leading-5 max-h-48 overflow-y-auto whitespace-pre-wrap">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}
    </div>
  )
}
