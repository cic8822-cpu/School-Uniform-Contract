import { ArrowRight } from 'lucide-react'
import type { ContractMethod } from '../../types'

interface HomePageProps {
  methods: ContractMethod[]
  onSelectMethod: (methodId: string) => void
}

const DEFAULT_HIGHLIGHT_METHOD_ID = 'CM-03'

export function HomePage({ methods, onSelectMethod }: HomePageProps) {
  return (
    <main className="page">
      <p className="eyebrow">SCHOOL UNIFORM PROCUREMENT GUIDE</p>
      <h1>교복 학교주관구매 길라잡이</h1>
      <p className="lede">계약방법 안내부터 서식 작성·출력까지 한 곳에서</p>

      <section className="banner">
        <b>한눈에 보는 계약 업무 길잡이</b>
        <button onClick={() => onSelectMethod(DEFAULT_HIGHLIGHT_METHOD_ID)}>
          GO <ArrowRight size={17} />
        </button>
      </section>

      <h2>계약 방식을 선택해 주세요</h2>
      <p className="muted">계약방법은 시스템이 자동 판정하지 않습니다. 담당자가 직접 선택합니다.</p>

      <div className="cards">
        {methods.map((method, index) => (
          <button
            key={method.id}
            className="card"
            style={{ '--accent': method.webThemeColor } as React.CSSProperties}
            onClick={() => onSelectMethod(method.id)}
          >
            <i />
            <small>
              STEP {index + 1} · {method.stepCount}단계
            </small>
            <h2>{method.webDisplayLabel}</h2>
            <p>{method.applicabilityNote}</p>
            <b>절차 확인 →</b>
          </button>
        ))}
      </div>
    </main>
  )
}
