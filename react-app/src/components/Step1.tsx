import { useState, useRef } from 'react'
import { runStep1 } from '../lib/step1Core'
import { saveMetadata } from '../lib/supabase'
import type { Metadata, AppSettings } from '../lib/types'
import FileUploader from './FileUploader'

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

  const reset = () => {
    setDone(false); setLogs([]); setFileName('')
    if (fileRef.current) fileRef.current.value = ''
  }

  return (
    <div className="space-y-4">
      {/* step-box — matches web_app.py .step-box CSS exactly */}
      <div className="bg-[#f0f4fa] border-l-4 border-[#2E75B6] px-5 py-4 rounded-md text-sm">
        <b>STEP 1 — ERP 파일 가공</b>: 맥산 ERP에서 추출한 <b>구매조회</b> 또는 <b>구매현황</b> 파일을
        업로드하면 생산자재_사용내역.xlsx를 자동 생성합니다.
        두 형식 모두 지원하며 자동으로 감지합니다. 품번·품명·구매처·리드타임은 <b>📋 품번 관리</b> 탭 정보로 자동 채워집니다.
      </div>

      {/* columns [2, 1] */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-start">
        <div className="md:col-span-2">
          <FileUploader
            label="구매조회 / 구매현황 파일 (ERP 원본)"
            fileRef={fileRef}
            fileName={fileName}
            onChange={setFileName}
          />
        </div>

        {/* st.info */}
        <div className="bg-[#e8f4fd] rounded-lg p-4 text-sm text-gray-700">
          <div className="font-semibold mb-2">저장된 품번</div>
          <div>케이블 <b>{nCable}</b> 타입 / 하우징 <b>{nHousing}</b> 타입</div>
          <hr className="border-gray-300 my-2" />
          <div>리드타임 기본값: <b>{ltDefault}일</b></div>
        </div>
      </div>

      <hr className="border-gray-200" />

      {/* buttons [5, 1] */}
      <div className="flex gap-2">
        <button
          onClick={run}
          disabled={running || !fileName}
          className="flex-1 bg-[#FF4B4B] hover:bg-[#e03030] disabled:bg-gray-300 text-white font-semibold py-2 px-4 rounded transition text-sm"
        >
          {running ? '⏳ 처리 중...' : '▶ STEP 1 실행 — ERP 파일 가공 & 사용내역 생성'}
        </button>
        {(done || logs.length > 0) && (
          <button onClick={reset}
            className="px-4 py-2 text-sm border border-gray-300 rounded text-gray-600 hover:bg-gray-100 transition">
            🗑 초기화
          </button>
        )}
      </div>

      {done && (
        <div className="bg-[#D6F0D8] px-4 py-3 rounded-md text-sm font-semibold text-[#1a6a2a]">
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
