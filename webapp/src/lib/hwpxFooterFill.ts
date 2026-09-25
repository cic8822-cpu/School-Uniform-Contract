// scripts/hwpx_school_footer_fill.ps1의 Set-NameNearLabel 로직을 그대로 포팅함.
// 담당자·교장 라벨만 채우고, 학교 주소·전화(우편번호 표 기반 name="주소"/"전화")는
// 웹 마스터 데이터에 학교정보 조회 원본이 아직 없어 이번 구현에서는 보류함
// (웹앱_task.md·웹앱_로그.md에 사유 기록, 원본 스크립트는 수정하지 않음).
const HP_NS = 'http://www.hancom.co.kr/hwpml/2011/paragraph'

export interface FooterFillValues {
  담당자?: string
  교장?: string
}

function getCellText(cell: Element): string {
  const texts = Array.from(cell.getElementsByTagNameNS(HP_NS, 't'))
  return texts.map((t) => t.textContent ?? '').join('').trim()
}

function getCellAddr(cell: Element): { row: number; col: number } | null {
  const addr = cell.getElementsByTagNameNS(HP_NS, 'cellAddr')[0]
  if (!addr) return null
  const row = Number(addr.getAttribute('rowAddr'))
  const col = Number(addr.getAttribute('colAddr'))
  if (Number.isNaN(row) || Number.isNaN(col)) return null
  return { row, col }
}

function setCellText(doc: Document, cell: Element, value: string): boolean {
  const texts = Array.from(cell.getElementsByTagNameNS(HP_NS, 't'))
  if (texts.length === 0) {
    const run = cell.getElementsByTagNameNS(HP_NS, 'run')[0]
    if (!run) return false
    const newText = doc.createElementNS(HP_NS, 'hp:t')
    newText.textContent = value
    run.appendChild(newText)
    return true
  }
  texts[0].textContent = value
  for (let index = 1; index < texts.length; index += 1) texts[index].textContent = ''
  const paragraphs = Array.from(cell.getElementsByTagNameNS(HP_NS, 'p'))
  for (const paragraph of paragraphs) {
    const lineSegments = Array.from(paragraph.getElementsByTagNameNS(HP_NS, 'linesegarray'))
    for (const segment of lineSegments) segment.parentNode?.removeChild(segment)
  }
  return true
}

function setNameNearLabel(doc: Document, table: Element, label: string, value: string): number {
  if (!value) return 0
  const cells = Array.from(table.getElementsByTagNameNS(HP_NS, 'tc'))
  const labelCell = cells.find((cell) => getCellText(cell) === label)
  if (!labelCell) return 0
  const labelAddr = getCellAddr(labelCell)
  if (!labelAddr) return 0

  const rowCells = cells
    .map((cell) => ({ cell, addr: getCellAddr(cell) }))
    .filter((entry): entry is { cell: Element; addr: { row: number; col: number } } =>
      entry.addr !== null && entry.addr.row === labelAddr.row && entry.addr.col > labelAddr.col
    )
    .sort((a, b) => a.addr.col - b.addr.col)

  const placeholderCell = rowCells.find((entry) => getCellText(entry.cell) === '○○○')
  if (placeholderCell && setCellText(doc, placeholderCell.cell, value)) return 1

  const namedEmptyCell = rowCells.find(
    (entry) => getCellText(entry.cell) === '' && entry.cell.getAttribute('name')
  )
  if (namedEmptyCell && setCellText(doc, namedEmptyCell.cell, value)) return 1

  return 0
}

/**
 * HWPX(Blob)의 결재란에서 셀 텍스트가 정확히 "담당자"/"교장"인 라벨을 찾아
 * 같은 행 오른쪽의 빈 칸(○○○ 플레이스홀더 우선, 없으면 name 속성이 있는
 * 빈 칸)을 값으로 채운 새 Blob을 돌려준다. 값이 없는 라벨은 건드리지 않는다.
 */
export async function fillHwpxFooter(hwpxBlob: Blob, values: FooterFillValues): Promise<Blob> {
  const { default: JSZip } = await import('jszip')
  const zip = await JSZip.loadAsync(hwpxBlob)
  const entries = Object.keys(zip.files).filter((name) => /^Contents\/section\d+\.xml$/.test(name))
  const parser = new DOMParser()
  const serializer = new XMLSerializer()

  for (const name of entries) {
    const file = zip.file(name)
    if (!file) continue
    const xmlText = await file.async('string')
    const doc = parser.parseFromString(xmlText, 'application/xml')
    if (doc.getElementsByTagName('parsererror').length > 0) {
      throw new Error(`HWPX 본문 XML 파싱 실패: ${name}`)
    }

    let changed = false
    const tables = Array.from(doc.getElementsByTagNameNS(HP_NS, 'tbl'))
    for (const table of tables) {
      if (values.담당자 && setNameNearLabel(doc, table, '담당자', values.담당자) > 0) changed = true
      if (values.교장 && setNameNearLabel(doc, table, '교장', values.교장) > 0) changed = true
    }

    if (changed) zip.file(name, serializer.serializeToString(doc))
  }

  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' })
}
