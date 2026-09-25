import type { ContractMethod } from '../../types'
import './contractGuide.css'

interface ContractGuidePageProps {
  methods: ContractMethod[]
  onOpenProcedure: (methodId: string) => void
}

export function ContractGuidePage({ methods, onOpenProcedure }: ContractGuidePageProps) {
  const sortedMethods = [...methods].sort((a, b) => a.order - b.order)

  return (
    <main className="page">
      <p className="crumb">홈 › 계약방법안내</p>
      <h1>계약방법안내</h1>
      <p className="lede">
        계약방법은 시스템이 금액 등으로 자동 판정하지 않습니다. 아래 비교표를 참고해 담당자가 직접
        선택합니다.
      </p>

      <div className="contract-guide-table" role="table">
        <div className="contract-guide-row contract-guide-head" role="row">
          <div className="contract-guide-cell contract-guide-label" role="columnheader">
            비교 항목
          </div>
          {sortedMethods.map((method) => (
            <div
              key={method.id}
              className="contract-guide-cell"
              role="columnheader"
              style={{ '--accent': method.webThemeColor } as React.CSSProperties}
            >
              <b>{method.webDisplayLabel}</b>
            </div>
          ))}
        </div>

        <GuideRow label="적용 조건(참고)" methods={sortedMethods} render={(m) => m.applicabilityNote} />
        <GuideRow label="절차 단계 수" methods={sortedMethods} render={(m) => `${m.stepCount}단계`} />
        <GuideRow
          label="공통 필수 서식"
          methods={sortedMethods}
          render={(m) => m.commonRequiredForms?.join(', ') ?? '매뉴얼 확인 필요'}
        />
        <GuideRow
          label="전용 핵심 서식"
          methods={sortedMethods}
          render={(m) => (m.dedicatedCoreForms.length > 0 ? m.dedicatedCoreForms.join(', ') : '—')}
        />
        <GuideRow
          label="절차 내 서식 수"
          methods={sortedMethods}
          render={(m) => `${m.workflowFormCount}종`}
        />
        <GuideRow label="담당자 안내" methods={sortedMethods} render={(m) => m.excelGuideText} />
        <div className="contract-guide-row" role="row">
          <div className="contract-guide-cell contract-guide-label" role="rowheader">
            절차 확인
          </div>
          {sortedMethods.map((method) => (
            <div key={method.id} className="contract-guide-cell" role="cell">
              <button className="primary" onClick={() => onOpenProcedure(method.id)}>
                절차 보기 →
              </button>
            </div>
          ))}
        </div>
      </div>

      <p className="contract-guide-note">
        참고: 적용 조건은 참고용 안내이며 금액 기준 등 법적 판단 근거가 아닙니다. 2단계 입찰(규격·가격
        동시)과 일반경쟁입찰은 절차 안내상 같은 입찰계약 흐름으로 연결되며, 일반경쟁입찰 전용 절차는
        매뉴얼 확인이 필요합니다.
      </p>
    </main>
  )
}

interface GuideRowProps {
  label: string
  methods: ContractMethod[]
  render: (method: ContractMethod) => string
}

function GuideRow({ label, methods, render }: GuideRowProps) {
  return (
    <div className="contract-guide-row" role="row">
      <div className="contract-guide-cell contract-guide-label" role="rowheader">
        {label}
      </div>
      {methods.map((method) => (
        <div key={method.id} className="contract-guide-cell" role="cell">
          {render(method)}
        </div>
      ))}
    </div>
  )
}
