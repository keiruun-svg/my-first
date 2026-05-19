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

// user gesture 컨텍스트가 살아있는 동안 호출해야 함 (버튼 onClick 직후)
// 반환값: 저장 핸들(선택 완료) | null(미지원·오류) | 'cancelled'(사용자 취소)
export async function pickSaveFile(filename: string): Promise<unknown | null | 'cancelled'> {
  if (!('showSaveFilePicker' in window)) return null
  try {
    return await (window as unknown as {
      showSaveFilePicker(o: unknown): Promise<unknown>
    }).showSaveFilePicker({
      suggestedName: filename,
      types: [{ description: 'Excel 통합문서', accept: { [XLSX_MIME]: ['.xlsx'] } }],
    })
  } catch (e: unknown) {
    if (e instanceof Error && e.name === 'AbortError') return 'cancelled'
    return null  // 오류 → blob fallback
  }
}

export async function writeToFileHandle(handle: unknown, buffer: ArrayBuffer): Promise<void> {
  const h = handle as {
    createWritable(): Promise<{ write(data: Blob): Promise<void>; close(): Promise<void> }>
  }
  const writable = await h.createWritable()
  await writable.write(new Blob([buffer], { type: XLSX_MIME }))
  await writable.close()
}

export function today(): string {
  const d   = new Date()
  const pad = (n: number) => String(n).padStart(2, '0')
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`
}
