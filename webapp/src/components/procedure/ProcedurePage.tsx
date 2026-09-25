import { Check } from 'lucide-react'
import type { ContractMethod, FormRecord, Workflow } from '../../types'

interface ProcedurePageProps {
  method: ContractMethod
  workflow: Workflow
  stepNo: number
  onSelectStep: (stepNo: number) => void
  forms: FormRecord[]
  missingFieldCount: (form: FormRecord) => number
  onOpenForm: (form: FormRecord) => void
}

export function ProcedurePage({
  method,
  workflow,
  stepNo,
  onSelectStep,
  forms,
  missingFieldCount,
  onOpenForm,
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
          <p className="desc">{step.description}</p>

          <h3>해야 할 일</h3>
          <ul>
            {(step.checklistItems.length > 0 ? step.checklistItems : [{ text: step.description }]).map(
              (item, index) => (
                <li key={index}>{item.text}</li>
              )
            )}
          </ul>

          <div className="caution">
            {step.cautions[0]?.text ?? '매뉴얼과 현행 규정을 최종 확인합니다.'}
          </div>

          <h3>관련 서식</h3>
          {step.formIds.map((formId) => {
            const form = forms.find((candidate) => candidate.formId === formId)
            if (!form) return null
            const missing = missingFieldCount(form)
            return (
              <button className="row" key={formId} onClick={() => onOpenForm(form)}>
                <b>{formId}</b>
                <span>{form.title}</span>
                <em>{missing > 0 ? '▲ 자료필요' : '● 작성가능'}</em>
              </button>
            )
          })}
        </article>
      </div>
    </main>
  )
}
