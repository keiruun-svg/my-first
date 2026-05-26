import { useState } from 'react'
import { assignLocation, LOCATION_RULES } from '../lib/assignLocation'

export default function EmpLocationGenerator() {
  const [name, setName]     = useState('')
  const [copied, setCopied] = useState(false)

  const location = name.trim() ? assignLocation(name.trim()) : null

  const handleCopy = () => {
    if (!location) return
    navigator.clipboard.writeText(location)
    setCopied(true)
    setTimeout(() => setCopied(false), 1500)
  }

  return (
    <div className="max-w-2xl space-y-5">

      {/* 입력 */}
      <div className="bg-white border border-gray-200 rounded-lg p-5 space-y-4">
        <p className="text-sm text-gray-600">
          품명을 입력하면 EMP 창고 로케이션이 자동 배정됩니다.
        </p>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">품명</label>
          <input
            type="text"
            value={name}
            onChange={e => setName(e.target.value)}
            placeholder="예: DOJC-SM-LC/PC, Optical Cable 0.9, Pigtail..."
            className="w-full border border-gray-300 rounded px-3 py-2 text-sm focus:outline-none focus:border-sky-500"
            autoFocus
          />
        </div>

        {/* 결과 */}
        <div className={`rounded-lg border-2 p-4 transition-colors ${
          !name.trim()   ? 'border-dashed border-gray-300 bg-gray-50' :
          location       ? 'border-sky-500 bg-sky-50' :
                           'border-orange-300 bg-orange-50'
        }`}>
          <div className="text-xs text-gray-500 mb-1">배정 로케이션</div>
          {!name.trim() ? (
            <div className="text-gray-400 text-sm">품명을 입력하세요</div>
          ) : location ? (
            <div className="flex items-center gap-3">
              <span className="font-mono text-2xl font-bold text-sky-800">{location}</span>
              <button
                onClick={handleCopy}
                className="px-3 py-1 text-sm bg-sky-600 text-white rounded hover:bg-sky-700 transition"
              >
                {copied ? '✅ 복사됨' : '복사'}
              </button>
            </div>
          ) : (
            <div className="text-orange-700 text-sm font-medium">
              ⚠ 규칙 미해당 — 직접 배정 필요
            </div>
          )}
        </div>
      </div>

      {/* 규칙 참조표 */}
      <div className="border border-gray-200 rounded-lg overflow-hidden">
        <div className="px-4 py-2.5 bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600">
          로케이션 배정 규칙 ({LOCATION_RULES.length}개)
        </div>
        <table className="w-full text-xs">
          <thead>
            <tr className="bg-[#1F3864] text-white">
              <th className="px-4 py-2 text-left font-medium w-24">로케이션</th>
              <th className="px-4 py-2 text-left font-medium w-32">품목 분류</th>
              <th className="px-4 py-2 text-left font-medium">키워드 (품명 포함 시 매칭)</th>
            </tr>
          </thead>
          <tbody>
            {LOCATION_RULES.map((rule, i) => (
              <tr
                key={i}
                className={`border-b border-gray-100 last:border-0 transition-colors ${
                  location === rule.location && name.trim() ? 'bg-sky-50' : i % 2 === 0 ? 'bg-white' : 'bg-gray-50'
                }`}
              >
                <td className="px-4 py-2 font-mono font-bold text-sky-700">{rule.location}</td>
                <td className="px-4 py-2 text-gray-700">{rule.description}</td>
                <td className="px-4 py-2 text-gray-500">
                  {rule.keyword.map((kw, j) => (
                    <code key={j} className="mr-1.5 px-1 py-0.5 bg-white border border-gray-200 rounded text-xs">{kw}</code>
                  ))}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="px-4 py-2 bg-gray-50 border-t border-gray-100 text-xs text-gray-400">
          규칙은 위에서 아래 순서로 적용되며, 첫 번째 매칭 규칙이 사용됩니다.
        </div>
      </div>
    </div>
  )
}
