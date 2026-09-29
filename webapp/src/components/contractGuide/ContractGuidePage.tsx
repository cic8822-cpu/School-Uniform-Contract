import type { ContractMethod, FormRecord, Workflow, WorkflowStep } from '../../types'
import {
  COMMON_FLOW,
  CONTRACT_GLOSSARY,
  METHOD_EXPLAINERS,
  type FlowGroup,
  type MethodExplainer,
} from './contractGuideContent'
import './contractGuide.css'

interface ContractGuidePageProps {
  methods: ContractMethod[]
  workflows: Workflow[]
  forms: FormRecord[]
  onOpenProcedure: (methodId: string) => void
}

type Assignee = '학교' | '업체' | '공동'
const ASSIGNEE_ORDER: Assignee[] = ['학교', '업체', '공동']
const ASSIGNEE_TITLE: Record<Assignee, string> = {
  학교: '학교가 하는 일',
  업체: '업체가 하는 일',
  공동: '학교와 업체가 함께 하는 일',
}

interface FlowItem {
  label: string
  note: string
}

/** 단계를 하나씩 보여주거나, 묶음이 있으면 큰 흐름으로 묶어 보여준다. */
function buildFlow(steps: WorkflowStep[], groups?: FlowGroup[]): FlowItem[] {
  if (!groups) return steps.map((step) => ({ label: step.stepName, note: '' }))
  return groups.map((group) => {
    const first = Math.min(...group.stepNos)
    const last = Math.max(...group.stepNos)
    return { label: group.label, note: first === last ? `${first}단계` : `${first}~${last}단계` }
  })
}

function stepsByAssignee(steps: WorkflowStep[], assignee: Assignee): string[] {
  return steps.filter((step) => step.assignee === assignee).map((step) => step.stepName)
}

export function ContractGuidePage({ methods, workflows, forms, onOpenProcedure }: ContractGuidePageProps) {
  const sortedMethods = [...methods].sort((a, b) => a.order - b.order)
  const formTitle = (formId: string): string => forms.find((form) => form.formId === formId)?.title ?? formId

  return (
    <main className="page">
      <p className="crumb">홈 › 계약방법안내</p>
      <h1>계약방법안내</h1>
      <p className="lede">
        교복을 사려면 학교가 업체와 계약을 맺어야 합니다. 어떤 업체를 어떤 방식으로 고르느냐에 따라
        계약방법이 4가지로 나뉩니다. 계약방법은 시스템이 자동으로 정하지 않고 담당자가 직접 선택합니다.
      </p>

      <section className="cg-common" aria-labelledby="cg-common-title">
        <h2 id="cg-common-title">먼저 이것만 기억하세요</h2>
        <p>
          어떤 방법이든 큰 흐름은 같습니다. <b>다른 점은 첫 번째, 업체를 정하는 방식뿐</b>입니다.
        </p>
        <ol className="cg-common-flow">
          {COMMON_FLOW.map((item, index) => (
            <li key={item.label} className={item.differs ? 'differs' : undefined}>
              <span className="cg-common-no">{index + 1}</span>
              <b>{item.label}</b>
              <small>{item.note}</small>
            </li>
          ))}
        </ol>
      </section>

      <div className="cg-grid">
        {sortedMethods.map((method) => {
          const explainer: MethodExplainer | undefined = METHOD_EXPLAINERS[method.id]
          const workflow = workflows.find((candidate) => candidate.workflowId === method.workflowId)
          if (!explainer || !workflow) return null
          const flow = buildFlow(workflow.steps, explainer.flowGroups)
          return (
            <article
              key={method.id}
              className="cg-card"
              style={{ '--accent': method.webThemeColor } as React.CSSProperties}
              aria-labelledby={`cg-${method.id}`}
            >
              <div className="cg-card-head">
                <small>
                  {method.stepCount}단계 · 서식 {method.workflowFormCount}종
                </small>
                <h2 id={`cg-${method.id}`}>{method.webDisplayLabel}</h2>
                <p className="cg-one-line">{explainer.oneLine}</p>
              </div>

              <h3>어떻게 진행되나요</h3>
              <p>{explainer.howItWorks}</p>

              <h3>큰 흐름</h3>
              <ol className="cg-flow">
                {flow.map((item) => (
                  <li key={item.label}>
                    <b>{item.label}</b>
                    {item.note && <small>{item.note}</small>}
                  </li>
                ))}
              </ol>

              <h3>알아둘 점</h3>
              <ul className="cg-points">
                {explainer.keyPoints.map((point) => (
                  <li key={point}>{point}</li>
                ))}
              </ul>

              <h3>누가 무엇을 하나요</h3>
              <dl className="cg-roles">
                {ASSIGNEE_ORDER.map((assignee) => {
                  const stepNames = stepsByAssignee(workflow.steps, assignee)
                  if (stepNames.length === 0) return null
                  return (
                    <div key={assignee}>
                      <dt>{ASSIGNEE_TITLE[assignee]}</dt>
                      <dd>{stepNames.join(' · ')}</dd>
                    </div>
                  )
                })}
              </dl>

              <h3>처음 준비할 서식</h3>
              <ul className="cg-forms">
                {explainer.coreFormIds.map((formId) => (
                  <li key={formId}>
                    {formTitle(formId)} <small>{formId}</small>
                  </li>
                ))}
              </ul>
              {method.workflowFormCount > explainer.coreFormIds.length && (
                <p className="cg-more">
                  그 밖의 서식까지 모두 {method.workflowFormCount}종은 계약절차에서 단계별로 확인합니다.
                </p>
              )}

              <p className="cg-when">
                <b>이런 경우를 떠올려 보세요(참고)</b>
                {explainer.goodWhen}
              </p>

              <button className="primary" onClick={() => onOpenProcedure(method.id)}>
                이 방법의 절차 보기 →
              </button>
            </article>
          )
        })}
      </div>

      <section className="cg-compare" aria-labelledby="cg-compare-title">
        <h2 id="cg-compare-title">한눈에 비교</h2>
        <div className="cg-compare-scroll">
          <table>
            <thead>
              <tr>
                <th scope="col">비교</th>
                {sortedMethods.map((method) => (
                  <th key={method.id} scope="col">
                    {method.webDisplayLabel}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                <th scope="row">업체를 정하는 방식</th>
                {sortedMethods.map((method) => (
                  <td key={method.id}>{METHOD_EXPLAINERS[method.id]?.selection}</td>
                ))}
              </tr>
              <tr>
                <th scope="row">절차 단계</th>
                {sortedMethods.map((method) => (
                  <td key={method.id}>{method.stepCount}단계</td>
                ))}
              </tr>
              <tr>
                <th scope="row">관련 서식</th>
                {sortedMethods.map((method) => (
                  <td key={method.id}>{method.workflowFormCount}종</td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <section className="cg-glossary" aria-labelledby="cg-glossary-title">
        <h2 id="cg-glossary-title">처음 보는 용어 쉬운 풀이</h2>
        <dl>
          {CONTRACT_GLOSSARY.map((item) => (
            <div key={item.term}>
              <dt>{item.term}</dt>
              <dd>{item.meaning}</dd>
            </div>
          ))}
        </dl>
      </section>

      <p className="contract-guide-note">
        이 화면은 이해를 돕는 참고 안내입니다. 계약방법별 적용 요건·금액 기준·기한·법령은 해당 매뉴얼과
        현행 규정을 최종 확인하세요.
      </p>
    </main>
  )
}
