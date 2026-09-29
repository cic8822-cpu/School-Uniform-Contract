import { Check } from 'lucide-react'
import { formStatusOf, groupMissingByTab } from '../../lib/formReadiness'
import type { ContractMethod, FieldDef, FormRecord, InputTabId, Workflow, WorkflowStep } from '../../types'

interface ProcedurePageProps {
  method: ContractMethod
  workflow: Workflow
  stepNo: number
  onSelectStep: (stepNo: number) => void
  forms: FormRecord[]
  missingFieldsOf: (form: FormRecord) => FieldDef[]
  onOpenForm: (form: FormRecord) => void
  onGoInput: (tab: InputTabId) => void
}

const SENTENCE_BOUNDARY = /(?<=[.!?。])\s+/

/** 체크리스트가 없으면 설명문을 행동 목록으로 대신 쓴다. */
function actionTexts(step: WorkflowStep): string[] {
  return step.checklistItems.length > 0 ? step.checklistItems.map((item) => item.text) : [step.description]
}

/** 설명문과 같은 문장이 "해야 할 일"에도 있으면 본문 첫 문단으로 또 보여줄 필요가 없다. */
function isDescriptionRepeatedInActions(step: WorkflowStep): boolean {
  const description = step.description.trim()
  return actionTexts(step).some((text) => text.trim() === description)
}

/** 한 문단에 여러 행동이 들어 있으면 문장 단위로 나눠 글머리 목록으로 보여준다. */
function toActionItems(step: WorkflowStep): string[] {
  return actionTexts(step)
    .flatMap((text) => text.split(SENTENCE_BOUNDARY))
    .map((sentence) => sentence.trim())
    .filter((sentence) => sentence.length > 0)
}

export function ProcedurePage({
  method,
  workflow,
  stepNo,
  onSelectStep,
  forms,
  missingFieldsOf,
  onOpenForm,
  onGoInput,
}: ProcedurePageProps) {
  const step = workflow.steps.find((candidate) => candidate.stepNo === stepNo) ?? workflow.steps[0]

  return (
    <main className="page">
      <p className="crumb">
        홈 › 계약절차 › {method.webDisplayLabel}
      </p>
      <h1>계약절차 · {method.webDisplayLabel}</h1>

      <div className="procedure">
        <aside>
          {workflow.steps.map((candidate) => (
            <button
              key={candidate.stepNo}
              className={
                candidate.stepNo === stepNo ? 'active' : candidate.stepNo < stepNo ? 'done' : ''
              }
              onClick={() => onSelectStep(candidate.stepNo)}
            >
              <b>{candidate.stepNo < stepNo ? <Check size={15} /> : candidate.stepNo}</b> STEP{' '}
              {String(candidate.stepNo).padStart(2, '0')} · {candidate.stepName}
            </button>
          ))}
        </aside>

        <article>
          <small>
            STEP {String(step.stepNo).padStart(2, '0')} / {workflow.steps.length}
          </small>
          <h2>{step.stepName}</h2>
          {!isDescriptionRepeatedInActions(step) && <p className="desc">{step.description}</p>}

          <h3>해야 할 일</h3>
          <ul>
            {toActionItems(step).map((text, index) => (
              <li key={index}>{text}</li>
            ))}
          </ul>

          <div className="caution">
            {step.cautions[0]?.text ?? '매뉴얼과 현행 규정을 최종 확인합니다.'}
          </div>

          <h3>관련 서식</h3>
          {step.formIds.map((formId) => {
            const form = forms.find((candidate) => candidate.formId === formId)
            if (!form) return null
            const missing = missingFieldsOf(form)
            const status = formStatusOf(form, missing)
            const missingGroups = status.kind === 'missing' ? groupMissingByTab(missing) : []
            return (
              <div className="row-wrap" key={formId}>
                <button className="row" onClick={() => onOpenForm(form)}>
                  <b>{formId}</b>
                  <span>{form.title}</span>
                  <em className={`status-${status.kind}`}>{status.text}</em>
                </button>
                {missingGroups.length > 0 && (
                  <ul className="row-missing" aria-label={`${form.title} 필요한 자료`}>
                    {missingGroups.map((group) => (
                      <li key={group.tab}>
                        <span>
                          <b>{group.tab}</b> 탭 · {group.labels.join(', ')}
                        </span>
                        <button className="link-button" onClick={() => onGoInput(group.tab)}>
                          입력하러 가기 →
                        </button>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            )
          })}
        </article>
      </div>
    </main>
  )
}
