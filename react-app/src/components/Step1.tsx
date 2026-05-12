import { useState, useRef } from 'react'
import { runStep1 } from '../lib/step1Core'
import { saveMetadata } from '../lib/supabase'
import type { Metadata, AppSettings } from '../lib/types'

interface Props {
  metadata: Metadata
  setMetadata: (m: Metadata) => void
  settings?: AppSettings
}

export default function Step1({ metadata, setMetadata }: Props) {
  const [logs, setLogs] = useState<string[]>([])
  const [running, setRunning] = useState(false)
  const [done, setDone] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)

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
        for (const k of result.newCableKeys) {
          newMeta.cable[k] = { 품번: '', 품명: '', 구매처: '', 리드타임: null }
        }
        for (const k of result.newHousingKeys) {
          newMeta.housing[k] = { 품번: '', 품명: '', 구매처: '', 리드타임: null }
        }
        saveMetadata(newMeta)
        setMetadata(newMeta)
        setLogs(l => [...l, `✅ 신규 타입 ${result.newCableKeys.length + result.newHousingKeys.length}건 품번관리에 추가됨`])
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
      <div className="bg-blue-50 border border-blue-200 rounded p-4">
        <h3 className="font-bold text-blue-800 mb-2">STEP 1 — ERP 파일 파싱 & 품번 등록</h3>
        <p className="text-sm text-blue-700">구매조회/구매현황 Excel을 업로드하면 케이블·하우징 타입을 자동 감지하여 품번관리 탭에 등록합니다.</p>
      </div>

      <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center">
        <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" id="step1-file" />
        <label htmlFor="step1-file" className="cursor-pointer">
          <div className="text-4xl mb-2">📂</div>
          <div className="text-gray-600">ERP 파일 선택 (.xlsx)</div>
          <div className="text-sm text-gray-400 mt-1">구매조회 또는 구매현황 파일</div>
        </label>
        {fileRef.current?.files?.[0] && (
          <div className="mt-2 text-sm text-green-600">✓ {fileRef.current.files[0].name}</div>
        )}
      </div>

      <button
        onClick={run}
        disabled={running}
        className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white font-bold py-3 rounded-lg transition"
      >
        {running ? '⏳ 처리 중...' : '▶ STEP 1 실행'}
      </button>

      {logs.length > 0 && (
        <div className="bg-gray-900 text-green-300 rounded p-4 font-mono text-sm space-y-1 max-h-60 overflow-y-auto">
          {logs.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {done && (
        <div className="bg-green-50 border border-green-300 rounded p-4 text-green-800">
          ✅ 완료! 품번관리 탭에서 신규 타입의 품번·품명·구매처·리드타임을 입력해주세요.
        </div>
      )}
    </div>
  )
}
