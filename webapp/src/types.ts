// 웹앱 마스터 데이터(_workspace/05_web/data/*.json)의 타입 정의.
// 이 타입들은 마스터 데이터 스키마의 부분집합이며, 화면 구현에서 실제로 쓰는
// 필드만 선언한다(WP-02 산출물 자체는 이 타입보다 더 많은 필드를 가질 수 있음).

export interface ContractMethod {
  id: string
  order: number
  webDisplayLabel: string
  workflowId: string
  stepCount: number
  applicabilityNote: string
  applicabilityIsReferenceOnly: boolean
  autoDecision: boolean
  excelGuideText: string
  commonRequiredForms: string[] | null
  dedicatedCoreForms: string[]
  workflowFormIds: string[]
  workflowFormCount: number
  webThemeColor: string
  webThemeColorName: string
}

export interface ContractMethodsFile {
  policy: string
  notes: string[]
  contractMethods: ContractMethod[]
}

export interface SourceRef {
  document: string
  locator: string
  status: string
}

export interface ChecklistItem {
  id: string
  text: string
  sourceRefs: SourceRef[]
  evidenceStatus: string
}

export interface CautionItem {
  text: string
  sourceRefs: SourceRef[]
  evidenceStatus: string
}

export interface ManualReference {
  document: string
  locator: string
  status: string
  note: string
}

export interface WorkflowStep {
  stepNo: number
  stepName: string
  stepNameRaw: string
  description: string
  assignee: string
  assigneeBasis: string
  formIds: string[]
  formLabelsExcel: string[]
  checklistItems: ChecklistItem[]
  cautions: CautionItem[]
  manualReferences: ManualReference
}

export interface Workflow {
  workflowId: string
  steps: WorkflowStep[]
}

export interface StandardPhase {
  phase: string
  todo: string
  formRangeRaw: string
  formIds: string[]
  representativeFormId: string
}

export interface WorkflowsFile {
  workflows: Workflow[]
  standardPhases: {
    source: string
    phases: StandardPhase[]
  }
}

export interface HwpxCoverage {
  level: string
  supportedFieldIds: string[]
  unsupportedFieldIds: string[]
  displayLabel: string
  outputCompletion: string
}

export interface ContractMethodScope {
  status: string
  displayLabel: string
  reason: string
  filterBehavior: string
}

export type OutputMethod = 'both' | 'hwpx' | 'none'
export type ImplementationStatus = 'implemented' | 'hwpxPriority' | 'notImplemented'

export interface FormRecord {
  formId: string
  manualNo: string
  title: string
  stage: string
  docType: string
  author: string
  owner: string
  difficulty: string
  contractMethods: string[]
  contractMethodsNote: string | null
  inputFieldIds: string[]
  requiredFieldIds: string[]
  computedFieldIds: string[]
  fieldGroups: string[]
  boundary: string
  boundaryNote: string | null
  implementationStatus: ImplementationStatus
  outputMethod: OutputMethod
  hwpxTemplate: string | null
  hwpxTokens: string[]
  privacyProtected: boolean
  privacyAllowedFieldIds: string[] | null
  hwpxCoverage: HwpxCoverage
  contractMethodScope: ContractMethodScope
}

export interface FormsFile {
  forms: FormRecord[]
}

export type FieldType =
  | 'string'
  | 'year'
  | 'date'
  | 'dateOrPeriod'
  | 'datetime'
  | 'enum'
  | 'currency'
  | 'integer'
  | 'number'
  | 'stringArray'
  | 'text'

export type FieldRequirement = 'required' | 'conditional' | 'computed'

export interface FieldDef {
  fieldId: string
  label: string
  group: string
  repeatGroup: string | null
  type: FieldType
  required: FieldRequirement
  validation: string
  hwpxToken: string | null
  hwpxFooterFill: string | null
}

export interface RepeatGroupDef {
  group: string
  fieldIds: string[]
  excelRows: string
  total?: string
  extraColumnsWithoutId?: string[]
}

export interface DerivedLookup {
  label: string
  source: string
  usage: string
}

export interface ProtectedFormForbiddenField {
  label: string
  inputUiForbidden: boolean
  storeForbidden: boolean
  tokenForbidden: boolean
  networkForbidden: boolean
}

export interface ProtectedForm {
  formId: string
  allowedFieldIds: string[]
  outputPolicy: string
  forbiddenFields: ProtectedFormForbiddenField[]
}

export interface GlobalForbiddenItem {
  label: string
  inputUiForbidden: boolean
}

export interface FieldsFile {
  fields: FieldDef[]
  repeatGroups: RepeatGroupDef[]
  derivedLookups: DerivedLookup[]
  protectedForms: ProtectedForm[]
  globalForbidden: GlobalForbiddenItem[]
}

export interface AppData {
  methods: ContractMethod[]
  workflows: Workflow[]
  standardPhases: StandardPhase[]
  forms: FormRecord[]
  fields: FieldDef[]
  repeatGroups: RepeatGroupDef[]
  protectedForms: ProtectedForm[]
  globalForbidden: GlobalForbiddenItem[]
}

/** 공통·사업·문서별(비반복) Field 값. Field ID -> 입력값(문자열). */
export type FieldValues = Record<string, string>

/** 반복그룹 한 행의 값. Field ID -> 입력값(문자열). */
export type RepeatRow = Record<string, string>

/** 반복그룹 이름 -> 행 배열. */
export type RepeatValues = Record<string, RepeatRow[]>

export type RouteName =
  | 'home'
  | 'contractGuide'
  | 'procedure'
  | 'input'
  | 'forms'
  | 'export'
  | 'guide'
