import { toNumber } from './format'
import type { RepeatRow } from '../types'

/** K-01: 품목별 금액 = 수량(R-05) × 단가(R-06). */
export function computeItemAmount(row: RepeatRow): number {
  return toNumber(row['R-05'] ?? '') * toNumber(row['R-06'] ?? '')
}

/** K-02: 합계금액 = 품목별 금액(K-01)의 합. */
export function computeItemAmountTotal(rows: RepeatRow[]): number {
  return rows.reduce((total, row) => total + computeItemAmount(row), 0)
}

/** K-03: 평가 총점 = 평가 점수(R-07)의 합. */
export function computeEvaluationTotal(rows: RepeatRow[]): number {
  return rows.reduce((total, row) => total + toNumber(row['R-07'] ?? ''), 0)
}
