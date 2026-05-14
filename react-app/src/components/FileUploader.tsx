import { useRef, useState } from 'react'

interface Props {
  label: string
  fileName: string
  onFile: (file: File) => void
  optional?: boolean
}

export default function FileUploader({ label, fileName, onFile, optional = false }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [dragging, setDragging] = useState(false)
  const dragCounter = useRef(0)

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    dragCounter.current = 0
    setDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (file) onFile(file)
  }

  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault()
    dragCounter.current++
    setDragging(true)
  }

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault()
    if (--dragCounter.current === 0) setDragging(false)
  }

  return (
    <div className="space-y-1">
      <label className="block text-sm font-medium text-gray-700">
        {label}
        {optional && <span className="text-gray-400 font-normal ml-1">(선택)</span>}
      </label>
      <div
        className={`border rounded px-4 py-6 text-center bg-[#fafafa] transition-colors ${
          dragging
            ? 'border-[#2E75B6] bg-[#e8f4fd] border-dashed'
            : fileName
              ? 'border-[#2E75B6] bg-[#f0f4fa]'
              : 'border-gray-300 hover:border-gray-400'
        }`}
        onDragOver={e => e.preventDefault()}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        {fileName ? (
          <div className="text-sm text-[#2E75B6] font-medium mb-2">📄 {fileName}</div>
        ) : (
          <>
            <div className="text-sm text-gray-500 mb-1">
              {dragging ? '여기에 놓으세요' : 'Drag and drop file here'}
            </div>
            <div className="text-xs text-gray-400 mb-3">Limit 200MB per file • XLSX</div>
          </>
        )}
        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          className="inline-flex items-center px-4 py-1.5 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded cursor-pointer hover:bg-gray-50 transition"
        >
          Browse files
        </button>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) onFile(f) }}
        />
      </div>
    </div>
  )
}
