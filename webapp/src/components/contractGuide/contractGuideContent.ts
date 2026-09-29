// 계약방법안내 화면에 쓰는 초보자용 설명 문구.
// 근거: workflows.json 단계 설명(교복 매뉴얼 기반)과 contract-methods.json의 참고 문구.
// 금액 기준 등 법령 판단은 새로 만들지 않고 "매뉴얼·법령 확인"으로 남긴다.

export interface FlowGroup {
  label: string
  stepNos: number[]
}

export interface MethodExplainer {
  /** 한 줄 정의 */
  oneLine: string
  /** 업체를 어떻게 정하는지 풀어 쓴 설명 */
  howItWorks: string
  /** 비교표에 쓰는 짧은 선정 방식 */
  selection: string
  /** 이런 경우를 떠올려 보세요(참고용) */
  goodWhen: string
  /** 초보자가 꼭 알아둘 점 */
  keyPoints: string[]
  /** 처음 준비할 서식 Form ID */
  coreFormIds: string[]
  /** 큰 흐름으로 묶을 단계. 없으면 단계를 하나씩 보여준다 */
  flowGroups?: FlowGroup[]
}

const BID_FLOW_GROUPS: FlowGroup[] = [
  { label: '공고', stepNos: [1, 2] },
  { label: '업체 입찰', stepNos: [3] },
  { label: '규격 심사', stepNos: [4] },
  { label: '가격 개찰·낙찰', stepNos: [5, 6] },
  { label: '계약 체결', stepNos: [7] },
  { label: '납품·검수·대금', stepNos: [8, 9, 10] },
  { label: '변경계약', stepNos: [11] },
]

const BID_CORE_FORM_IDS = ['F-009', 'F-010', 'F-011', 'F-008', 'F-012']

export const METHOD_EXPLAINERS: Record<string, MethodExplainer> = {
  'CM-01': {
    oneLine: '한 업체에서 견적서를 받아, 그 견적을 기준으로 계약하는 가장 간단한 방법입니다.',
    howItWorks:
      '학교가 업체 1곳에 교복 사양서를 주고 견적서를 받습니다. 제출된 견적을 기준으로 계약을 맺습니다. 여러 업체를 비교하지 않아서 절차가 짧습니다.',
    selection: '업체 1곳의 견적',
    goodWhen: '추정가격이 소액인 경우 (금액 기준은 지방계약법을 따르므로 매뉴얼에서 확인)',
    keyPoints: [
      '견적서를 요청할 때 교복 사양서를 함께 줍니다.',
      '계약할 때 계약 특수조건을 첨부합니다.',
      '견적서와 수의계약 사유를 확인합니다.',
    ],
    coreFormIds: ['F-008', 'F-012', 'F-043', 'F-031'],
  },
  'CM-02': {
    oneLine: '2곳 이상 업체에서 견적서를 받아 비교하고, 가장 유리한 조건의 업체와 계약합니다.',
    howItWorks:
      '같은 교복 사양서를 2개 이상 업체에 주고 견적서를 받습니다. 받은 견적서를 비교해 가장 유리한 조건을 낸 업체를 고르고, 비교한 근거는 서류로 남겨 둡니다.',
    selection: '2곳 이상 견적을 비교',
    goodWhen: '2개 이상 업체의 비교견적이 필요한 경우',
    keyPoints: [
      '모든 업체에 똑같은 교복 사양서를 주어 같은 조건으로 견적을 받습니다.',
      '견적을 비교한 근거를 관련 서류로 남겨 둡니다.',
      '계약금액은 견적가 × 예정수량으로 정합니다.',
    ],
    coreFormIds: ['F-008', 'F-012', 'F-043', 'F-031'],
  },
  'CM-03': {
    oneLine:
      '먼저 품질(규격)이 기준을 넘는 업체만 가려내고, 그 업체들끼리 가격으로 경쟁하게 하는 입찰입니다.',
    howItWorks:
      '학교가 입찰을 공고하면 업체가 규격제안서와 가격입찰서를 냅니다. 교복선정위원회가 규격제안서만 먼저 평가해 기준(적격) 이상인 업체를 가려내고, 통과한 업체의 가격입찰서만 열어 예정가격 이하 최저가 업체를 낙찰자로 정합니다.',
    selection: '규격 심사를 통과한 업체 중 최저가',
    goodWhen: '품질 적격심사 후 가격경쟁이 필요한 경우 (지방계약법 시행령 제18조③)',
    keyPoints: [
      '매뉴얼 절차에서는 2단계 입찰(규격·가격 동시)이 원칙입니다.',
      '규격제안서는 학교에 직접 제출하고, 가격입찰서는 나라장터(G2B)로 제출합니다.',
      '규격제안서는 블라인드 심사를 위해 업체명과 표시를 지웁니다.',
      '가격입찰서는 규격을 통과한 업체 것만 열어 봅니다.',
    ],
    coreFormIds: BID_CORE_FORM_IDS,
    flowGroups: BID_FLOW_GROUPS,
  },
  'CM-04': {
    oneLine: '입찰 공고를 내서 참여를 원하는 업체끼리 공개적으로 경쟁하게 하는 방법입니다.',
    howItWorks:
      '학교가 입찰 공고를 게시하면 참여할 수 있는 업체가 입찰합니다. 이 앱에는 일반경쟁입찰만의 별도 절차 자료가 없어서, 2단계 입찰과 같은 입찰 절차(11단계)를 참고용으로 보여 드립니다.',
    selection: '공개경쟁 (전용 절차는 매뉴얼 확인)',
    goodWhen: '공개경쟁이 원칙인 경우',
    keyPoints: [
      '절차와 서식은 2단계 입찰과 같은 체계를 참고합니다.',
      '일반경쟁입찰만의 절차는 이 앱에 없으니 매뉴얼에서 다시 확인하세요.',
      '나라장터(G2B)에 이용자 등록을 마친 업체만 참여할 수 있습니다.',
    ],
    coreFormIds: BID_CORE_FORM_IDS,
    flowGroups: BID_FLOW_GROUPS,
  },
}

/** 모든 방법에 공통인 큰 흐름. 다른 점은 첫 단계(업체를 정하는 방식)뿐이다. */
export const COMMON_FLOW: { label: string; note: string; differs: boolean }[] = [
  { label: '업체 정하기', note: '방법마다 다릅니다', differs: true },
  { label: '계약 체결', note: '선정된 업체와 계약', differs: false },
  { label: '납품·검사검수', note: '학교가 규격서와 맞는지 확인', differs: false },
  { label: '대금 지급', note: '검수 완료 후 지급', differs: false },
]

export interface GlossaryTerm {
  term: string
  meaning: string
}

export const CONTRACT_GLOSSARY: GlossaryTerm[] = [
  { term: '수의계약', meaning: '입찰 공고 없이 업체로부터 견적을 받아 맺는 계약입니다. 이 앱의 1인견적·2인견적이 여기에 해당합니다.' },
  { term: '견적서', meaning: '업체가 "이 조건이면 이 가격에 납품하겠다"고 적어 내는 문서입니다.' },
  { term: '추정가격', meaning: '계약 방법을 정하는 기준이 되는 예상 금액입니다. 정확한 산정과 기준 금액은 매뉴얼과 법령에서 확인하세요.' },
  { term: '교복 사양서', meaning: '학교가 원하는 교복의 품목·규격을 적은 문서입니다. 모든 업체가 같은 조건으로 견적·입찰하도록 함께 제공합니다.' },
  { term: '입찰', meaning: '공고를 보고 업체가 계약 조건(규격·가격)을 적어 내는 것입니다.' },
  { term: '나라장터(G2B)', meaning: '국가종합전자조달시스템입니다. 입찰 공고를 게재하고 가격입찰서를 받는 곳입니다.' },
  { term: '규격(적격)심사', meaning: '교복선정위원회가 제안서를 평가항목·배점기준에 따라 점수로 평가해, 종합평점이 기준 이상인 업체만 다음 단계로 보내는 심사입니다.' },
  { term: '블라인드 심사', meaning: '업체명과 표시를 지우고 평가해, 업체를 알지 못한 채 제안서만 보고 심사하는 방식입니다.' },
  { term: '개찰', meaning: '제출된 입찰서를 열어 내용을 확인하는 것입니다. 2단계 입찰에서는 규격을 통과한 업체의 가격입찰서만 엽니다.' },
  { term: '예정가격', meaning: '낙찰 여부를 가르는 기준 가격입니다. 예정가격 이하 최저가 입찰자가 낙찰자가 됩니다.' },
  { term: '낙찰자', meaning: '입찰에서 계약 상대방으로 결정된 업체입니다.' },
  { term: '검사·검수', meaning: '납품된 교복이 규격서와 일치하는지 학교가 확인하는 절차입니다. 검수 합격 후에 대금을 지급합니다.' },
]
