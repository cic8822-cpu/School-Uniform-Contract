import { useEffect, useRef } from 'react'
import { collectFooterFillValues, collectTokenValues } from '../../lib/hwpxExport'
import { inputTabOf } from '../../lib/formReadiness'
import type { FieldDef, FieldValues, FormRecord, InputTabId, RepeatValues } from '../../types'

interface ExportConfirmProps {
  forms: FormRecord[]
  fields: FieldDef[]
  values: FieldValues
  repeats: RepeatValues
  busy: boolean
  onCancel: () => void
  onConfirm: () => void
  onGoInput: (tab: InputTabId) => void
}

const CONFIRM_TABS: InputTabId[] = ['공통', '사업', '문서별']
const REPEAT_TABS: InputTabId[] = ['품목', '업체', '위원', '평가']
const FOOTER_FIELD_IDS = ['C-07', 'C-08']
const MISSING_TEXT = '미입력'

function countFilledRows(repeats: RepeatValues, group: string): number {
  return (repeats[group] ?? []).filter((row) => Object.values(row).some((cell) => cell !== '')).length
}

function exportCandidates(forms: FormRecord[]): FormRecord[] {
  return forms.filter((form) => form.outputMethod !== 'none' && Boolean(form.hwpxTemplate))
}

/** 일괄 ZIP에 담길 서식과, 값이 비어 제외될 서식을 미리 계산한다. */
function planExport(forms: FormRecord[], fields: FieldDef[], values: FieldValues) {
  const plans = exportCandidates(forms).map((form) => ({
    form,
    missingTokens: collectTokenValues(form, fields, values).missingTokens,
  }))
  return {
    ready: plans.filter((plan) => plan.missingTokens.length === 0),
    skipped: plans.filter((plan) => plan.missingTokens.length > 0),
  }
}

/** 서식 토큰과 결재란(담당자·교장)에 실제로 들어가는 Field만 확인 대상으로 고른다. */
function pickFieldsToConfirm(forms: FormRecord[], fields: FieldDef[], values: FieldValues): FieldDef[] {
  const candidates = exportCandidates(forms)
  const tokens = new Set(candidates.flatMap((form) => form.hwpxTokens))
  const needsFooter = candidates.some((form) => collectFooterFillValues(form, values).applicable)
  return fields.filter(
    (field) =>
      !field.repeatGroup &&
      CONFIRM_TABS.includes(inputTabOf(field)) &&
      ((field.hwpxToken !== null && tokens.has(field.hwpxToken)) ||
        (needsFooter && FOOTER_FIELD_IDS.includes(field.fieldId)))
  )
}

/** 일괄 ZIP을 내려받기 직전에, 어떤 값으로 어떤 서식이 만들어지는지 보여주고 확정을 받는 화면. */
export function ExportConfirm({
  forms,
  fields,
  values,
  repeats,
  busy,
  onCancel,
  onConfirm,
  onGoInput,
}: ExportConfirmProps) {
  const headingRef = useRef<HTMLHeadingElement>(null)
  const { ready, skipped } = planExport(forms, fields, values)
  const confirmFields = pickFieldsToConfirm(forms, fields, values)

  useEffect(() => {
    headingRef.current?.focus()
    headingRef.current?.scrollIntoView({ block: 'start', behavior: 'smooth' })
  }, [])

  return (
    <section className="export-confirm" aria-labelledby="export-confirm-title">
      <h3 id="export-confirm-title" ref={headingRef} tabIndex={-1}>
        내려받기 전에 입력 내용을 확인해 주세요
      </h3>
      <p>
        아래 값으로 서식 <b>{ready.length}종</b>을 ZIP에 담습니다.
        {skipped.length > 0 && (
          <>
            {' '}
            값이 비어 있는 <b>{skipped.length}종</b>은 제외됩니다.
          </>
        )}
      </p>

      {CONFIRM_TABS.map((tab) => {
        const tabFields = confirmFields.filter((field) => inputTabOf(field) === tab)
        if (tabFields.length === 0) return null
        return (
          <div className="export-confirm-group" key={tab}>
            <div className="export-confirm-group-head">
              <h4>{tab} 정보</h4>
              <button className="link-button" onClick={() => onGoInput(tab)}>
                {tab} 탭에서 수정 →
              </button>
            </div>
            <dl>
              {tabFields.map((field) => {
                const value = values[field.fieldId] ?? ''
                return (
                  <div key={field.fieldId}>
                    <dt>{field.label}</dt>
                    <dd className={value === '' ? 'is-empty' : undefined}>{value === '' ? MISSING_TEXT : value}</dd>
                  </div>
                )
              })}
            </dl>
          </div>
        )
      })}

      <p className="export-confirm-repeats">
        참고 · 입력된 행 수:{' '}
        {REPEAT_TABS.map((tab) => `${tab} ${countFilledRows(repeats, tab)}행`).join(' · ')}
      </p>

      {skipped.length > 0 && (
        <details className="export-confirm-skipped">
          <summary>제외될 서식 {skipped.length}종 보기</summary>
          <ul>
            {skipped.map((plan) => (
              <li key={plan.form.formId}>
                <b>{plan.form.formId}</b> {plan.form.title} — 부족: {plan.missingTokens.join(', ')}
              </li>
            ))}
          </ul>
        </details>
      )}

      <div className="export-confirm-actions">
        <button className="primary" disabled={busy || ready.length === 0} onClick={onConfirm}>
          {busy ? 'ZIP 생성 중…' : `확인했습니다. ${ready.length}종 ZIP 내려받기`}
        </button>
        <button className="btn-secondary" disabled={busy} onClick={() => onGoInput('공통')}>
          기초자료 수정하러 가기
        </button>
        <button className="btn-secondary" disabled={busy} onClick={onCancel}>
          취소
        </button>
      </div>
      {ready.length === 0 && (
        <p className="field-input-error">내려받을 수 있는 서식이 없습니다. 값을 먼저 입력해 주세요.</p>
      )}
    </section>
  )
}
