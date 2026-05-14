import { useState } from 'react'

interface Props {
  label: string
  fileRef: React.RefObject<HTMLInputElement | null>
  fileName: string
  onChange: (name: string) => void
  optional?: boolean
}

export default function FileUploader({ label, fileRef, fileName, onChange, optional = false }: Props) {
  const [dragging, setDragging] = useState(false)

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    setDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (!file) return
    if (fileRef.current) {
      const dt = new DataTransfer()
      dt.items.add(file)
      fileRef.current.files = dt.files
    }
    onChange(file.name)
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
        onDragOver={e => { e.preventDefault(); setDragging(true) }}
        onDragEnter={e => { e.preventDefault(); setDragging(true) }}
        onDragLeave={e => { e.preventDefault(); setDragging(false) }}
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
        <label className="inline-flex items-center px-4 py-1.5 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded cursor-pointer hover:bg-gray-50 transition">
          Browse files
          <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden"
            onChange={e => onChange(e.target.files?.[0]?.name ?? '')} />
        </label>
      </div>
    </div>
  )
}
