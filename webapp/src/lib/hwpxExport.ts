import type { FieldDef, FieldValues, FormRecord } from '../types'
import { buildTokenFilledHwpx } from './hwpx'
import { fillHwpxFooter } from './hwpxFooterFill'
import { downloadBlob } from './download'
import { formatDotDate, formatSchoolYear } from './format'
import { assetUrl } from './assetUrl'

/** forms.json의 hwpxTemplate 경로(artifacts/hwpx/P2-04/...)를 웹 정적 자산 경로로 바꾼다. */
export function hwpxTemplateUrl(hwpxTemplate: string): string {
  const fileName = hwpxTemplate.split('/').pop() ?? hwpxTemplate
  return assetUrl(`templates/${fileName}`)
}

function formatTokenValue(token: string, rawValue: string): string {
  if (token === '발행일') return formatDotDate(rawValue)
  if (token === '학년도') return formatSchoolYear(rawValue)
  return rawValue
}

export interface TokenCollectionResult {
  tokenValues: Record<string, string>
  missingTokens: string[]
}

/** form.hwpxTokens에 필요한 값을 fields.json의 hwpxToken 매핑으로 모은다. */
export function collectTokenValues(
  form: FormRecord,
  fields: FieldDef[],
  values: FieldValues
): TokenCollectionResult {
  const tokenValues: Record<string, string> = {}
  const missingTokens: string[] = []

  for (const token of form.hwpxTokens) {
    const field = fields.find((candidate) => candidate.hwpxToken === token)
    const rawValue = field ? (values[field.fieldId] ?? '') : ''
    if (!rawValue) {
      missingTokens.push(token)
      continue
    }
    tokenValues[token] = formatTokenValue(token, rawValue)
  }

  return { tokenValues, missingTokens }
}

/**
 * 결재란 라벨 채움 대상 서식(담당자 C-08·교장 C-07을 입력받는 15종,
 * template-mapping.md §3)인지와 채울 값을 함께 돌려준다.
 */
export function collectFooterFillValues(
  form: FormRecord,
  values: FieldValues
): { applicable: boolean; 담당자?: string; 교장?: string } {
  const applicable = form.inputFieldIds.includes('C-07') || form.inputFieldIds.includes('C-08')
  if (!applicable) return { applicable: false }
  return {
    applicable: true,
    담당자: values['C-08'] || undefined,
    교장: values['C-07'] || undefined,
  }
}

export interface HwpxExportResult {
  missingTokens: string[]
}

/**
 * 서식의 HWPX 템플릿에 토큰을 채우고, 결재란 대상 서식이면 담당자·교장 라벨도
 * 채운 뒤 파일로 내려받는다. 필수 토큰이 비어 있으면 다운로드 대신 누락 목록을
 * 돌려주므로, 호출부(UI)가 사용자에게 안내할 수 있다.
 */
export async function exportHwpxDocument(
  form: FormRecord,
  fields: FieldDef[],
  values: FieldValues
): Promise<HwpxExportResult> {
  if (!form.hwpxTemplate) throw new Error(`HWPX 템플릿이 없는 서식입니다: ${form.formId}`)

  const { tokenValues, missingTokens } = collectTokenValues(form, fields, values)
  if (missingTokens.length > 0) return { missingTokens }

  const tokenFilledBlob = await buildTokenFilledHwpx(hwpxTemplateUrl(form.hwpxTemplate), tokenValues)
  const footer = collectFooterFillValues(form, values)
  const finalBlob =
    footer.applicable && (footer.담당자 || footer.교장)
      ? await fillHwpxFooter(tokenFilledBlob, footer)
      : tokenFilledBlob

  downloadBlob(finalBlob, `${form.formId}_${form.title}.hwpx`)
  return { missingTokens: [] }
}
