const A4_WIDTH_MM = 210
const A4_HEIGHT_MM = 297

/**
 * 문서 미리보기 DOM(A4 용지 영역)을 캡처해 PDF로 저장한다. 브라우저에서만
 * 실행되며 서버로 전송하지 않는다(ADR-002). 내용이 한 페이지보다 길면
 * 이어지는 페이지에 나눠 담는다.
 *
 * jspdf·html2canvas는 PDF 출력 버튼을 누를 때만 필요한 무거운 라이브러리라
 * 동적 import로 분리해 초기 번들 크기를 줄인다(web/performance.md 번들 예산).
 */
export async function exportElementToPdf(element: HTMLElement, filename: string): Promise<void> {
  const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
    import('html2canvas'),
    import('jspdf'),
  ])

  const canvas = await html2canvas(element, { scale: 2, useCORS: true, backgroundColor: '#ffffff' })
  const imageData = canvas.toDataURL('image/png')

  const pdf = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' })
  const imageWidthMm = A4_WIDTH_MM
  const imageHeightMm = (canvas.height * imageWidthMm) / canvas.width

  let remainingHeightMm = imageHeightMm
  let positionMm = 0
  pdf.addImage(imageData, 'PNG', 0, positionMm, imageWidthMm, imageHeightMm)
  remainingHeightMm -= A4_HEIGHT_MM

  while (remainingHeightMm > 0) {
    positionMm = remainingHeightMm - imageHeightMm
    pdf.addPage()
    pdf.addImage(imageData, 'PNG', 0, positionMm, imageWidthMm, imageHeightMm)
    remainingHeightMm -= A4_HEIGHT_MM
  }

  pdf.save(filename)
}
