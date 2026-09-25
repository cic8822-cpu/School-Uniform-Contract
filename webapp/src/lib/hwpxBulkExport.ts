import type { FieldDef, FieldValues, FormRecord } from '../types'
import { buildTokenFilledHwpx } from './hwpx'
import { fillHwpxFooter } from './hwpxFooterFill'
import { downloadBlob } from './download'
import { collectFooterFillValues, collectTokenValues, hwpxTemplateUrl } from './hwpxExport'

export interface BulkExportResult {
  succeeded: string[]
  skipped: { formId: string; reason: string }[]
}

const BULK_ZIP_FILENAME = '교복구매_서식_일괄.zip'

/**
 * HWPX 템플릿이 있는 서식(outputMethod !== 'none') 전체를 현재 기초자료 값으로
 * 채워 하나의 ZIP으로 내려받는다. 필수 토큰이 비어 있는 서식은 건너뛰고
 * 사유를 결과로 돌려주므로, 호출부가 사용자에게 어떤 서식이 빠졌는지 안내할
 * 수 있다. jszip은 이 버튼을 실제로 누를 때만 동적으로 불러온다.
 */
export async function exportAllHwpxAsZip(
  forms: FormRecord[],
  fields: FieldDef[],
  values: FieldValues
): Promise<BulkExportResult> {
  const { default: JSZip } = await import('jszip')
  const archive = new JSZip()
  const succeeded: string[] = []
  const skipped: { formId: string; reason: string }[] = []

  const candidates = forms.filter(
    (form): form is FormRecord & { hwpxTemplate: string } =>
      form.outputMethod !== 'none' && Boolean(form.hwpxTemplate)
  )

  for (const form of candidates) {
    const { tokenValues, missingTokens } = collectTokenValues(form, fields, values)
    if (missingTokens.length > 0) {
      skipped.push({ formId: form.formId, reason: `누락값: ${missingTokens.join(', ')}` })
      continue
    }

    try {
      const tokenFilledBlob = await buildTokenFilledHwpx(hwpxTemplateUrl(form.hwpxTemplate), tokenValues)
      const footer = collectFooterFillValues(form, values)
      const finalBlob =
        footer.applicable && (footer.담당자 || footer.교장)
          ? await fillHwpxFooter(tokenFilledBlob, footer)
          : tokenFilledBlob
      archive.file(`${form.formId}_${form.title}.hwpx`, finalBlob)
      succeeded.push(form.formId)
    } catch (caught: unknown) {
      skipped.push({ formId: form.formId, reason: caught instanceof Error ? caught.message : '생성 실패' })
    }
  }

  if (succeeded.length > 0) {
    const zipBlob = await archive.generateAsync({ type: 'blob' })
    downloadBlob(zipBlob, BULK_ZIP_FILENAME)
  }

  return { succeeded, skipped }
}
