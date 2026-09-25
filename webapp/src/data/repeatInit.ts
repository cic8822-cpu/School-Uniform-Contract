import type { RepeatGroupDef, RepeatRow, RepeatValues } from '../types'

/** repeatGroups 정의의 excelRows에 적힌 행수(예: "...29~38행(10행)")를 파싱한다. */
function parseRowCount(excelRows: string): number {
  const match = /\((\d+)행\)/.exec(excelRows)
  return match ? Number(match[1]) : 10
}

/** 반복그룹마다 빈 행을 excelRows 행수만큼 만들어 초기 상태를 구성한다. */
export function createInitialRepeatValues(repeatGroups: RepeatGroupDef[]): RepeatValues {
  const initial: RepeatValues = {}
  for (const group of repeatGroups) {
    const rowCount = parseRowCount(group.excelRows)
    const emptyRow: RepeatRow = Object.fromEntries(group.fieldIds.map((fieldId) => [fieldId, '']))
    initial[group.group] = Array.from({ length: rowCount }, () => ({ ...emptyRow }))
  }
  return initial
}
