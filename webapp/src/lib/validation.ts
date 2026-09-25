import type { FieldDef } from '../types'

const MIN_SCHOOL_YEAR = 2000
const MAX_SCHOOL_YEAR = 2100

/**
 * Field 정의의 validation 서술문에서 `\`값1\`·\`값2\`` 형태로 나열된 승인 목록을
 * 추출한다. 백틱으로 감싼 값이 없으면 목록형이 아니므로 null을 반환하고,
 * 이 경우 호출부는 자유 입력(텍스트)으로 대체한다.
 */
export function extractEnumOptions(validation: string): string[] | null {
  const matches = [...validation.matchAll(/`([^`]+)`/g)].map((match) => match[1])
  return matches.length > 0 ? matches : null
}

/**
 * Field 정의의 validation 서술문에서 `N~M자` 형태의 길이 제한을 추출한다.
 * 없으면 null.
 */
export function extractLengthRange(validation: string): { min: number; max: number } | null {
  const match = /(\d+)~(\d+)자/.exec(validation)
  if (!match) return null
  return { min: Number(match[1]), max: Number(match[2]) }
}

/** 입력값 하나에 대한 사용자 안내용 오류 메시지. 문제가 없으면 null. */
export function validateFieldValue(field: FieldDef, value: string): string | null {
  if (value === '') return null

  if (field.type === 'year') {
    const year = Number(value)
    if (!/^\d{4}$/.test(value) || year < MIN_SCHOOL_YEAR || year > MAX_SCHOOL_YEAR) {
      return `${field.label}은 ${MIN_SCHOOL_YEAR}~${MAX_SCHOOL_YEAR} 사이 4자리 연도여야 합니다.`
    }
    return null
  }

  if (field.type === 'integer' || field.type === 'currency') {
    if (!/^\d+$/.test(value)) return `${field.label}은 0 이상 정수여야 합니다.`
    return null
  }

  if (field.type === 'number') {
    if (Number.isNaN(Number(value))) return `${field.label}은 숫자여야 합니다.`
    return null
  }

  const lengthRange = extractLengthRange(field.validation)
  if (lengthRange && (value.length < lengthRange.min || value.length > lengthRange.max)) {
    return `${field.label}은 ${lengthRange.min}~${lengthRange.max}자여야 합니다.`
  }

  const enumOptions = field.type === 'enum' ? extractEnumOptions(field.validation) : null
  if (enumOptions && !enumOptions.includes(value)) {
    return `${field.label}은 승인된 목록(${enumOptions.join(', ')}) 중에서 선택해야 합니다.`
  }

  return null
}
