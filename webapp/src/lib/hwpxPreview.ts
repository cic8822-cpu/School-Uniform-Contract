import { hwpxTemplateUrl } from './hwpxExport'

export interface HwpxPreviewParagraph {
  kind: 'paragraph'
  parts: HwpxPreviewPart[]
}

export interface HwpxPreviewText {
  kind: 'text'
  text: string
}

export interface HwpxPreviewImage {
  kind: 'image'
  text: string
}

export interface HwpxPreviewCell {
  paragraphs: HwpxPreviewParagraph[]
  colSpan: number
  rowSpan: number
}

export interface HwpxPreviewTable {
  kind: 'table'
  rows: HwpxPreviewCell[][]
}

export type HwpxPreviewPart = HwpxPreviewText | HwpxPreviewImage | HwpxPreviewParagraph | HwpxPreviewTable
export type HwpxPreviewBlock = HwpxPreviewParagraph | HwpxPreviewTable

const sourceCache = new Map<string, Promise<HwpxPreviewBlock[]>>()

function elementChildren(element: Element, localName?: string): Element[] {
  return Array.from(element.childNodes).filter(
    (child): child is Element =>
      child.nodeType === 1 && (!localName || (child as Element).localName === localName)
  )
}

function textElementValue(element: Element): string {
  return Array.from(element.childNodes)
    .map((child) => {
      if (child.nodeType === 3 || child.nodeType === 4) return child.nodeValue ?? ''
      if (child.nodeType !== 1) return ''
      const childElement = child as Element
      if (childElement.localName === 'lineBreak') return '\n'
      if (childElement.localName === 'tab') return '\t'
      return childElement.textContent ?? ''
    })
    .join('')
}

function appendText(parts: HwpxPreviewPart[], text: string): void {
  const previous = parts.at(-1)
  if (previous?.kind === 'text') previous.text += text
  else parts.push({ kind: 'text', text })
}

function parseParagraph(paragraph: Element): HwpxPreviewParagraph {
  const parts: HwpxPreviewPart[] = []

  function visit(element: Element): void {
    if (element.localName === 'p') {
      parts.push(parseParagraph(element))
      return
    }
    if (element.localName === 'tbl') {
      parts.push(parseTable(element))
      return
    }
    if (element.localName === 't') {
      appendText(parts, textElementValue(element))
      return
    }
    if (['pic', 'ole', 'video'].includes(element.localName)) {
      parts.push({ kind: 'image', text: '[원문 삽입 이미지]' })
    }
    elementChildren(element).forEach(visit)
  }

  elementChildren(paragraph).forEach(visit)
  return { kind: 'paragraph', parts }
}

function parseTable(table: Element): HwpxPreviewTable {
  return {
    kind: 'table',
    rows: elementChildren(table, 'tr').map((row) =>
      elementChildren(row, 'tc').map((cell) => {
        const span = elementChildren(cell, 'cellSpan')[0]
        return {
          paragraphs: elementChildren(cell, 'subList').flatMap((subList) =>
            elementChildren(subList, 'p').map(parseParagraph)
          ),
          colSpan: Number(span?.getAttribute('colSpan') ?? '1'),
          rowSpan: Number(span?.getAttribute('rowSpan') ?? '1'),
        }
      })
    ),
  }
}

export function parseHwpxPreviewXml(xml: string): HwpxPreviewBlock[] {
  const document = new DOMParser().parseFromString(xml, 'application/xml')
  if (document.getElementsByTagName('parsererror').length > 0) {
    throw new Error('HWPX 본문 XML을 해석하지 못했습니다.')
  }

  const section = document.documentElement
  return elementChildren(section, 'p').map(parseParagraph)
}

async function loadSource(templateUrl: string): Promise<HwpxPreviewBlock[]> {
  const response = await fetch(templateUrl)
  if (!response.ok) throw new Error('HWPX 원문을 불러오지 못했습니다.')

  const { default: JSZip } = await import('jszip')
  const zip = await JSZip.loadAsync(await response.arrayBuffer())
  const sectionNames = Object.keys(zip.files)
    .filter((name) => /^Contents\/section\d+\.xml$/.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))
  if (sectionNames.length === 0) throw new Error('HWPX 원문 본문이 없습니다.')

  const sections = await Promise.all(
    sectionNames.map(async (name) => {
      const file = zip.file(name)
      if (!file) throw new Error(`HWPX 본문을 찾지 못했습니다: ${name}`)
      return parseHwpxPreviewXml(await file.async('string'))
    })
  )
  return sections.flat()
}

function replaceKnownTokens(value: string, tokenValues: Record<string, string>): string {
  return value.replace(/\{\{([^{}]+)\}\}/g, (match, token: string) => tokenValues[token] ?? match)
}

function replaceParagraphTokens(
  paragraph: HwpxPreviewParagraph,
  tokenValues: Record<string, string>
): HwpxPreviewParagraph {
  return {
    ...paragraph,
    parts: paragraph.parts.map((part) => {
      if (part.kind === 'table') return replaceTableTokens(part, tokenValues)
      if (part.kind === 'paragraph') return replaceParagraphTokens(part, tokenValues)
      return { ...part, text: replaceKnownTokens(part.text, tokenValues) }
    }),
  }
}

function replaceTableTokens(
  table: HwpxPreviewTable,
  tokenValues: Record<string, string>
): HwpxPreviewTable {
  return {
    ...table,
    rows: table.rows.map((row) =>
      row.map((cell) => ({
        ...cell,
        paragraphs: cell.paragraphs.map((paragraph) => replaceParagraphTokens(paragraph, tokenValues)),
      }))
    ),
  }
}

export async function loadHwpxPreview(
  hwpxTemplate: string,
  tokenValues: Record<string, string>
): Promise<HwpxPreviewBlock[]> {
  const templateUrl = hwpxTemplateUrl(hwpxTemplate)
  let source = sourceCache.get(templateUrl)
  if (!source) {
    source = loadSource(templateUrl)
    sourceCache.set(templateUrl, source)
    void source.catch(() => {
      if (sourceCache.get(templateUrl) === source) sourceCache.delete(templateUrl)
    })
  }
  return (await source).map((block) =>
    block.kind === 'paragraph'
      ? replaceParagraphTokens(block, tokenValues)
      : replaceTableTokens(block, tokenValues)
  )
}
