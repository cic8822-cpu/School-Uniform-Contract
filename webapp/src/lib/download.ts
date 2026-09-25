/** Blob을 브라우저에서 파일로 저장한다. 서버로 전송하지 않는다(ADR-002). */
export function downloadBlob(blob: Blob, filename: string): void {
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = filename
  link.click()
  URL.revokeObjectURL(link.href)
}
