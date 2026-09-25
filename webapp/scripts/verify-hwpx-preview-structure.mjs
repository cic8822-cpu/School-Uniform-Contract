import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFile, readdir } from 'node:fs/promises'
import { basename, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DOMParser } from '@xmldom/xmldom'
import JSZip from 'jszip'
import { createServer } from 'vite'

const root = fileURLToPath(new URL('..', import.meta.url))
const templateDir = join(root, 'public', 'templates')
const server = await createServer({ root, appType: 'custom', server: { middlewareMode: true } })
globalThis.DOMParser = DOMParser

function elements(parent, localName) {
  return Array.from(parent.childNodes).filter(
    (child) => child.nodeType === 1 && (!localName || child.localName === localName)
  )
}

function sourceMetrics(document) {
  const all = Array.from(document.getElementsByTagName('*'))
  const paragraphs = all.filter((element) => element.localName === 'p')
  const tables = all.filter((element) => element.localName === 'tbl')
  const cells = all.filter((element) => element.localName === 'tc')
  const spans = cells.map((cell) => {
    const span = elements(cell, 'cellSpan')[0]
    return [Number(span?.getAttribute('colSpan') ?? '1'), Number(span?.getAttribute('rowSpan') ?? '1')]
  })
  const cellParagraphCounts = cells.map((cell) =>
    elements(cell, 'subList').reduce((count, subList) => count + elements(subList, 'p').length, 0)
  )
  const tableDepths = tables.map((table) => {
    let depth = 1
    let parent = table.parentNode
    while (parent) {
      if (parent.localName === 'tbl') depth += 1
      parent = parent.parentNode
    }
    return depth
  })
  return { paragraphCount: paragraphs.length, tableCount: tables.length, spans, cellParagraphCounts, tableDepths }
}

function previewMetrics(blocks) {
  const result = { paragraphCount: 0, tableCount: 0, spans: [], cellParagraphCounts: [], tableDepths: [] }

  function visitParagraph(paragraph, tableDepth) {
    result.paragraphCount += 1
    for (const part of paragraph.parts) {
      if (part.kind === 'table') visitTable(part, tableDepth + 1)
      if (part.kind === 'paragraph') visitParagraph(part, tableDepth)
    }
  }

  function visitTable(table, depth) {
    result.tableCount += 1
    result.tableDepths.push(depth)
    for (const row of table.rows) {
      for (const cell of row) {
        result.spans.push([cell.colSpan, cell.rowSpan])
        result.cellParagraphCounts.push(cell.paragraphs.length)
        cell.paragraphs.forEach((paragraph) => visitParagraph(paragraph, depth))
      }
    }
  }

  for (const block of blocks) {
    if (block.kind === 'table') visitTable(block, 1)
    else visitParagraph(block, 0)
  }
  return result
}

try {
  const { parseHwpxPreviewXml } = await server.ssrLoadModule('/src/lib/hwpxPreview.ts')
  const names = (await readdir(templateDir)).filter((name) => name.endsWith('.hwpx')).sort()
  assert.equal(names.length, 48, '검증할 HWPX 템플릿은 정확히 48개여야 합니다.')

  const report = []
  for (const name of names) {
    const bytes = await readFile(join(templateDir, name))
    const zip = await JSZip.loadAsync(bytes)
    const sections = Object.keys(zip.files).filter((entry) => /^Contents\/section\d+\.xml$/.test(entry)).sort()
    const totalSource = { paragraphCount: 0, tableCount: 0, spans: [], cellParagraphCounts: [], tableDepths: [] }
    const totalPreview = { paragraphCount: 0, tableCount: 0, spans: [], cellParagraphCounts: [], tableDepths: [] }

    for (const section of sections) {
      const xml = await zip.file(section).async('string')
      const expected = sourceMetrics(new DOMParser().parseFromString(xml, 'application/xml'))
      const actual = previewMetrics(parseHwpxPreviewXml(xml))
      for (const key of Object.keys(totalSource)) {
        if (Array.isArray(totalSource[key])) {
          totalSource[key].push(...expected[key])
          totalPreview[key].push(...actual[key])
        } else {
          totalSource[key] += expected[key]
          totalPreview[key] += actual[key]
        }
      }
    }

    assert.deepEqual(totalPreview, totalSource, `${name}: 문단·표·병합·중첩 구조가 원문과 달라졌습니다.`)
    report.push({
      formId: basename(name).slice(0, 5),
      sha256: createHash('sha256').update(bytes).digest('hex'),
      paragraphs: totalPreview.paragraphCount,
      tables: totalPreview.tableCount,
      multiParagraphCells: totalPreview.cellParagraphCounts.filter((count) => count > 1).length,
      nestedTables: totalPreview.tableDepths.filter((depth) => depth > 1).length,
    })
  }

  console.log(JSON.stringify({ status: 'PASS', templateCount: report.length, templates: report }, null, 2))
} finally {
  await server.close()
}
