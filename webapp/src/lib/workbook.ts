import type { FieldValues, RepeatValues } from '../types'

const COMMON_SHEET = '업무자료'
const REPEAT_SHEET = '반복자료'

interface CommonRow {
  FieldID?: string
  값?: string
}

interface RepeatRowRecord {
  그룹?: string
  행번호?: number
  FieldID?: string
  값?: string
}

/**
 * 기초자료입력 값을 엑셀(.xlsx)로 내려받는다. 공통·사업·문서별 Field는
 * "업무자료" 시트에, 반복그룹(품목·업체·위원·평가) 값은 "반복자료" 시트에
 * 그룹·행번호·FieldID·값 4열로 펼쳐 저장한다. 서버로 전송하지 않는다(ADR-002).
 *
 * xlsx는 엑셀 내보내기/가져오기를 실제로 쓸 때만 필요한 무거운 라이브러리라
 * 동적 import로 분리해 초기 번들 크기를 줄인다(web/performance.md 번들 예산).
 */
export async function downloadWorkbook(values: FieldValues, repeats: RepeatValues): Promise<void> {
  const XLSX = await import('xlsx')
  const book = XLSX.utils.book_new()

  const commonRows: CommonRow[] = Object.entries(values).map(([fieldId, value]) => ({
    FieldID: fieldId,
    값: value,
  }))
  XLSX.utils.book_append_sheet(book, XLSX.utils.json_to_sheet(commonRows), COMMON_SHEET)

  const repeatRows: RepeatRowRecord[] = []
  for (const [group, rows] of Object.entries(repeats)) {
    rows.forEach((row, rowIndex) => {
      for (const [fieldId, value] of Object.entries(row)) {
        if (value === '') continue
        repeatRows.push({ 그룹: group, 행번호: rowIndex + 1, FieldID: fieldId, 값: value })
      }
    })
  }
  XLSX.utils.book_append_sheet(book, XLSX.utils.json_to_sheet(repeatRows), REPEAT_SHEET)

  XLSX.writeFile(book, '업무자료.xlsx', { compression: true })
}

export interface WorkbookReadResult {
  values: FieldValues
  repeats: RepeatValues
}

/** downloadWorkbook으로 내려받은 엑셀을 다시 읽어 값·반복행으로 복원한다. */
export async function readWorkbook(file: File): Promise<WorkbookReadResult> {
  const XLSX = await import('xlsx')
  const book = XLSX.read(await file.arrayBuffer(), { type: 'array' })

  const commonSheet = book.Sheets[COMMON_SHEET]
  if (!commonSheet) throw new Error(`${COMMON_SHEET} 시트를 찾지 못했습니다.`)
  const commonRows = XLSX.utils.sheet_to_json<CommonRow>(commonSheet, { defval: '' })
  const values = commonRows.reduce<FieldValues>((result, row) => {
    if (!row.FieldID) return result
    return { ...result, [row.FieldID]: String(row.값 ?? '') }
  }, {})

  const repeats: RepeatValues = {}
  const repeatSheet = book.Sheets[REPEAT_SHEET]
  if (repeatSheet) {
    const repeatRows = XLSX.utils.sheet_to_json<RepeatRowRecord>(repeatSheet, { defval: '' })
    for (const row of repeatRows) {
      if (!row.그룹 || !row.FieldID || !row.행번호) continue
      const rowIndex = Number(row.행번호) - 1
      const group = row.그룹
      const groupRows = repeats[group] ?? []
      while (groupRows.length <= rowIndex) groupRows.push({})
      groupRows[rowIndex] = { ...groupRows[rowIndex], [row.FieldID]: String(row.값 ?? '') }
      repeats[group] = groupRows
    }
  }

  return { values, repeats }
}
