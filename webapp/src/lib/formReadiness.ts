import type { FieldDef, FieldValues, FormRecord, InputTabId, RepeatValues } from '../types'
import { computeEvaluationTotal, computeItemAmount, computeItemAmountTotal } from './fieldCompute'

const COMPUTED_GROUP = '계산'
const MAX_LISTED_LABELS = 3

/** 계산 Field(K-xx)가 어느 입력 탭의 값에서 나오는지. 계산 Field는 직접 입력하지 않는다. */
const COMPUTED_SOURCE_TAB: Record<string, InputTabId> = {
  'K-01': '품목',
  'K-02': '품목',
  'K-03': '평가',
}

const INPUT_TAB_ORDER: InputTabId[] = ['공통', '사업', '문서별', '품목', '업체', '위원', '평가']

/** Field가 입력되는 탭. 반복그룹 Field는 그룹 이름, 계산 Field는 원천 값이 있는 탭을 돌려준다. */
export function inputTabOf(field: FieldDef): InputTabId {
  if (field.repeatGroup) return field.repeatGroup as InputTabId
  if (field.group === COMPUTED_GROUP) return COMPUTED_SOURCE_TAB[field.fieldId] ?? '공통'
  return field.group as InputTabId
}

function hasAnyRepeatValue(rows: RepeatValues[string] | undefined, fieldId: string): boolean {
  return (rows ?? []).some((row) => (row[fieldId] ?? '') !== '')
}

function isComputedFilled(field: FieldDef, repeats: RepeatValues): boolean {
  switch (field.fieldId) {
    case 'K-01':
      return (repeats['품목'] ?? []).some((row) => computeItemAmount(row) > 0)
    case 'K-02':
      return computeItemAmountTotal(repeats['품목'] ?? []) > 0
    case 'K-03':
      return computeEvaluationTotal(repeats['평가'] ?? []) > 0
    default:
      return false
  }
}

/**
 * 값이 채워졌는지. 반복그룹 Field는 행 하나라도 값이 있으면 채워진 것으로,
 * 계산 Field는 원천 값에서 결과가 나오면 채워진 것으로 본다.
 */
export function isFieldFilled(field: FieldDef, values: FieldValues, repeats: RepeatValues): boolean {
  if (field.repeatGroup) return hasAnyRepeatValue(repeats[field.repeatGroup], field.fieldId)
  if (field.group === COMPUTED_GROUP) return isComputedFilled(field, repeats)
  return (values[field.fieldId] ?? '') !== ''
}

/** 서식이 필수로 요구하는 Field 중 아직 값이 없는 것. 정의가 없는 ID는 건너뛴다. */
export function findMissingFields(
  form: FormRecord,
  fields: FieldDef[],
  values: FieldValues,
  repeats: RepeatValues
): FieldDef[] {
  return form.requiredFieldIds
    .map((fieldId) => fields.find((candidate) => candidate.fieldId === fieldId))
    .filter((field): field is FieldDef => field !== undefined)
    .filter((field) => !isFieldFilled(field, values, repeats))
}

export interface MissingGroup {
  tab: InputTabId
  labels: string[]
}

/** 부족한 Field를 입력 탭별로 묶는다(탭은 기초자료입력 화면의 순서대로). */
export function groupMissingByTab(missing: FieldDef[]): MissingGroup[] {
  return INPUT_TAB_ORDER.map((tab) => ({
    tab,
    labels: missing.filter((field) => inputTabOf(field) === tab).map((field) => field.label),
  })).filter((group) => group.labels.length > 0)
}

export type FormStatusKind = 'ready' | 'missing' | 'unavailable'

export interface FormStatus {
  kind: FormStatusKind
  text: string
}

/** 서식 하나의 작성 가능 상태와 화면에 쓸 문구. 미구현 서식은 자료 유무와 상관없이 별도 표시한다. */
export function formStatusOf(form: FormRecord, missing: FieldDef[]): FormStatus {
  if (form.implementationStatus === 'notImplemented') return { kind: 'unavailable', text: '○ 미구현' }
  if (missing.length > 0) return { kind: 'missing', text: `▲ 자료필요 ${missing.length}건` }
  return { kind: 'ready', text: '● 작성가능' }
}

/** "학교명, 학년도, 구매명 외 2건"처럼 카드에 들어갈 짧은 요약문. */
export function summarizeLabels(labels: string[]): string {
  const shown = labels.slice(0, MAX_LISTED_LABELS).join(', ')
  const rest = labels.length - MAX_LISTED_LABELS
  return rest > 0 ? `${shown} 외 ${rest}건` : shown
}
