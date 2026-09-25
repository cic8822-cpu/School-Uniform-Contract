/** `<input type="date">`의 ISO 값(YYYY-MM-DD)을 공문서 표기(YYYY. MM. DD.)로 변환한다. */
export function formatDotDate(isoDate: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(isoDate)
  if (!match) return isoDate
  const [, year, month, day] = match
  return `${year}. ${month}. ${day}.`
}

/** 4자리 연도 문자열이면 "학년도" 접미사를 붙인다(이미 붙어 있으면 그대로 둔다). */
export function formatSchoolYear(rawYear: string): string {
  if (/^\d{4}$/.test(rawYear)) return `${rawYear}학년도`
  return rawYear
}

export function toNumber(value: string): number {
  const normalized = value.replace(/,/g, '').trim()
  if (normalized === '') return 0
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : 0
}

export function formatCurrency(amount: number): string {
  return `${amount.toLocaleString('ko-KR')}원`
}
