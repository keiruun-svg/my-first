const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

export function downloadXlsx(buffer: ArrayBuffer, filename: string) {
  const blob = new Blob([buffer], { type: XLSX_MIME })
  const url  = URL.createObjectURL(blob)
  const a    = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

export async function saveAsXlsx(buffer: ArrayBuffer, filename: string): Promise<void> {
  if ('showSaveFilePicker' in window) {
    try {
      const handle = await (window as unknown as { showSaveFilePicker: (o: unknown) => Promise<FileSystemFileHandle> })
        .showSaveFilePicker({
          suggestedName: filename,
          types: [{ description: 'Excel 통합문서', accept: { [XLSX_MIME]: ['.xlsx'] } }],
        })
      const writable = await handle.createWritable()
      await writable.write(new Blob([buffer], { type: XLSX_MIME }))
      await writable.close()
      return
    } catch (e: unknown) {
      if (e instanceof Error && e.name === 'AbortError') return  // 사용자가 취소
      // 지원 안 되거나 오류 → fallback
    }
  }
  downloadXlsx(buffer, filename)
}

export function today(): string {
  return new Date().toISOString().slice(0, 10)
}
