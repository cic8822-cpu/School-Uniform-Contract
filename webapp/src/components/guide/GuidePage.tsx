import './guide.css'

interface GuideStep {
  title: string
  description: string
}

interface FaqItem {
  question: string
  answer: string
}

interface GlossaryItem {
  term: string
  definition: string
}

const STEPS: GuideStep[] = [
  {
    title: '① 계약방법 선택',
    description:
      '홈 화면 또는 계약방법안내에서 계약방법 4종(1인견적·2인견적·2단계 계약·일반입찰)을 비교하고 담당자가 직접 선택합니다. 시스템이 금액 등으로 자동 판정하지 않습니다.',
  },
  {
    title: '② 계약절차 확인',
    description: '선택한 계약방법의 단계별 절차와 각 단계에서 필요한 서식을 확인합니다.',
  },
  {
    title: '③ 기초자료 입력',
    description:
      '기초자료입력에서 공통·사업·문서별 정보와 품목·업체·위원·평가 반복그룹을 한 번만 입력합니다. 입력값은 이 브라우저에만 저장되며 서버로 전송되지 않습니다.',
  },
  {
    title: '④ 서식함에서 출력',
    description:
      '서식함에서 서식을 찾아 미리보기를 열고, 지원되는 서식은 HWPX 또는 PDF로 출력합니다. 개인정보 보호 서식(F-029·F-047·F-050)은 빈 양식으로만 출력됩니다.',
  },
  {
    title: '⑤ 내보내기로 백업',
    description:
      '입력한 기초자료를 엑셀 파일로 내려받아 보관하거나, 다른 브라우저에서 다시 불러올 수 있습니다. 완성된 서식은 ZIP으로 한 번에 내려받을 수도 있습니다.',
  },
]

const FAQ_ITEMS: FaqItem[] = [
  {
    question: '입력한 학교명·담당자 정보가 서버에 저장되나요?',
    answer:
      '아니요. 이 웹앱은 서버·DB가 없는 정적 웹앱입니다. 입력값·엑셀 업로드 결과·생성된 문서는 모두 이 브라우저 안에서만 처리되며 네트워크로 전송되지 않습니다.',
  },
  {
    question: '다른 컴퓨터에서 이어서 작업하려면 어떻게 하나요?',
    answer:
      '내보내기 화면 또는 기초자료입력 화면에서 "엑셀로 내려받기"를 눌러 업무자료.xlsx를 저장한 뒤, 다른 컴퓨터의 기초자료입력 화면에서 "엑셀에서 불러오기"로 같은 파일을 불러오면 됩니다.',
  },
  {
    question: '학생·학부모 개인정보를 입력하는 칸이 왜 없나요?',
    answer:
      '개인정보 보호 원칙(F-029·F-047·F-050)에 따라 학생·학부모 성명·연락처·신체 치수 등은 이 웹앱에서 입력·저장·자동반영하지 않습니다. 해당 서식은 빈 양식으로 출력해 손으로 작성·제출받는 것이 정상적인 업무 방식입니다.',
  },
  {
    question: 'HWPX 출력 버튼이 안 보이는 서식이 있어요.',
    answer:
      'HWPX 템플릿이 아직 없는 서식(F-026·F-028·F-051·F-052)이거나, 필수 토큰 값(학교명·학년도 등)이 비어 있는 경우입니다. 기초자료입력에서 해당 값을 먼저 채워 주세요.',
  },
  {
    question: 'HWPX 파일을 한컴오피스가 아닌 곳에서 열면 깨지나요?',
    answer:
      '이 웹앱은 원본 HWPX 템플릿의 텍스트 토큰만 바꾸고 나머지 구조는 그대로 두는 방식으로 생성하므로, 한컴오피스(또는 호환 뷰어)에서 원본과 동일하게 열립니다.',
  },
]

const GLOSSARY_ITEMS: GlossaryItem[] = [
  { term: 'Form ID', definition: '서식 하나를 가리키는 고유 번호(예: F-007). 서식함·계약절차·미리보기에서 같은 서식을 가리킬 때 공통으로 씁니다.' },
  { term: 'Field ID', definition: '기초자료입력 값 하나를 가리키는 고유 번호(예: C-01 학교명). 공통(C)·사업(B)·문서별(D)·반복(R)·계산(K) 그룹으로 나뉩니다.' },
  { term: '계약방법', definition: '1인견적 수의계약·2인견적 수의계약·2단계 계약(규격·가격 동시)·일반입찰계약 4종. 담당자가 직접 선택하며 시스템이 자동 판정하지 않습니다.' },
  { term: 'HWPX', definition: '한컴오피스 한글의 XML 기반 문서 형식. 이 웹앱은 원본 HWPX 템플릿의 허용된 토큰만 채워 새 HWPX를 생성합니다.' },
  { term: '토큰', definition: 'HWPX 문서 안의 {{학교명}}처럼 중괄호로 감싼 자리표시자. 허용 목록에 있는 토큰만 값으로 채워지고, 금지 패턴(성명·전화·주소 등)에 걸리는 토큰은 채우지 않습니다.' },
  { term: '반복그룹', definition: '품목·업체·위원·평가처럼 여러 행을 입력하는 Field 묶음. 기초자료입력에서 표 형태로 여러 줄을 입력합니다.' },
]

export function GuidePage() {
  return (
    <main className="page">
      <h1>사용안내</h1>
      <p className="lede">교복 학교주관구매 길라잡이는 서버·DB 없이 이 브라우저에서만 동작합니다.</p>

      <h2>이용 순서</h2>
      <div className="cards guide">
        {STEPS.map((step) => (
          <article className="card" key={step.title}>
            <h3>{step.title}</h3>
            <p>{step.description}</p>
          </article>
        ))}
      </div>

      <h2 className="guide-section-title">자주 묻는 질문</h2>
      <div className="guide-faq">
        {FAQ_ITEMS.map((item) => (
          <details key={item.question} className="guide-faq-item">
            <summary>{item.question}</summary>
            <p>{item.answer}</p>
          </details>
        ))}
      </div>

      <h2 className="guide-section-title">용어집</h2>
      <dl className="guide-glossary">
        {GLOSSARY_ITEMS.map((item) => (
          <div className="guide-glossary-item" key={item.term}>
            <dt>{item.term}</dt>
            <dd>{item.definition}</dd>
          </div>
        ))}
      </dl>
    </main>
  )
}
