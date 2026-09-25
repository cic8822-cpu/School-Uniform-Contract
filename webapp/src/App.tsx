import { useMemo, useState } from 'react'
import './App.css'
import { useAppData } from './data/useAppData'
import { createInitialRepeatValues } from './data/repeatInit'
import { Header } from './components/layout/Header'
import { Footer } from './components/layout/Footer'
import { HomePage } from './components/home/HomePage'
import { ContractGuidePage } from './components/contractGuide/ContractGuidePage'
import { ProcedurePage } from './components/procedure/ProcedurePage'
import { InputPage } from './components/input/InputPage'
import { FormsPage } from './components/forms/FormsPage'
import { DocumentPreviewModal } from './components/modal/DocumentPreviewModal'
import { GuidePage } from './components/guide/GuidePage'
import { ExportPage } from './components/exportPage/ExportPage'
import type { FieldValues, FormRecord, RepeatValues, RouteName } from './types'

const DEFAULT_METHOD_ID = 'CM-03'
const SCHOOL_YEAR_FIELD_ID = 'C-02'
const DEFAULT_SCHOOL_YEAR = '2026'

/** 반복행이 아닌 Field(공통·사업·문서별)의 초기값. 학년도만 편의상 기본값을 채워 둔다. */
function createInitialValues(fieldIds: string[]): FieldValues {
  return Object.fromEntries(
    fieldIds.map((fieldId) => [fieldId, fieldId === SCHOOL_YEAR_FIELD_ID ? DEFAULT_SCHOOL_YEAR : ''])
  )
}

/** 불러온 반복행을 그룹별 기본 행수(초기 상태)에 맞춰 채우거나 잘라낸다. */
function mergeImportedRepeats(initial: RepeatValues, imported: RepeatValues): RepeatValues {
  const merged: RepeatValues = {}
  for (const [group, initialRows] of Object.entries(initial)) {
    const importedRows = imported[group]
    if (!importedRows) {
      merged[group] = initialRows
      continue
    }
    merged[group] = initialRows.map((emptyRow, index) => ({ ...emptyRow, ...importedRows[index] }))
  }
  return merged
}

export default function App() {
  const { data, loading, error } = useAppData()

  const [route, setRoute] = useState<RouteName>('home')
  const [selectedMethodId, setSelectedMethodId] = useState(DEFAULT_METHOD_ID)
  const [stepNo, setStepNo] = useState(1)
  const [searchQuery, setSearchQuery] = useState('')
  const [activeForm, setActiveForm] = useState<FormRecord | null>(null)

  const nonRepeatFieldIds = useMemo(
    () => (data ? data.fields.filter((field) => !field.repeatGroup && field.group !== '계산').map((f) => f.fieldId) : []),
    [data]
  )
  const [values, setValues] = useState<FieldValues>({})
  const [repeats, setRepeats] = useState<RepeatValues>({})
  const [initialized, setInitialized] = useState(false)

  if (data && !initialized) {
    setValues(createInitialValues(nonRepeatFieldIds))
    setRepeats(createInitialRepeatValues(data.repeatGroups))
    setInitialized(true)
  }

  if (loading || !data) return <main className="loading">길라잡이를 불러오는 중입니다.</main>
  if (error) return <main className="loading">{error}</main>

  const selectedMethod = data.methods.find((method) => method.id === selectedMethodId) ?? data.methods[0]
  const selectedWorkflow = data.workflows.find((workflow) => workflow.workflowId === selectedMethod.workflowId)

  function missingFieldCount(form: FormRecord): number {
    return form.requiredFieldIds.filter((fieldId) => !values[fieldId]).length
  }

  function missingFieldIds(form: FormRecord): string[] {
    return form.requiredFieldIds.filter((fieldId) => !values[fieldId])
  }

  function openProcedure(methodId: string): void {
    setSelectedMethodId(methodId)
    setStepNo(1)
    setRoute('procedure')
  }

  function handleChangeValue(fieldId: string, value: string): void {
    setValues((previous) => ({ ...previous, [fieldId]: value }))
  }

  function handleChangeRepeatCell(group: string, rowIndex: number, fieldId: string, value: string): void {
    setRepeats((previous) => {
      const rows = previous[group] ?? []
      const nextRows = rows.map((row, index) => (index === rowIndex ? { ...row, [fieldId]: value } : row))
      return { ...previous, [group]: nextRows }
    })
  }

  function handleImport(result: { values: FieldValues; repeats: RepeatValues }): void {
    setValues((previous) => ({ ...previous, ...result.values }))
    setRepeats((previous) => mergeImportedRepeats(previous, result.repeats))
  }

  return (
    <div className="app">
      <a className="skip-link" href="#main-content">
        본문으로 바로가기
      </a>
      <Header onNavigate={setRoute} />

      <div id="main-content" tabIndex={-1}>
      {route === 'home' && <HomePage methods={data.methods} onSelectMethod={openProcedure} />}

      {route === 'contractGuide' && (
        <ContractGuidePage methods={data.methods} onOpenProcedure={openProcedure} />
      )}

      {route === 'procedure' && selectedWorkflow && (
        <ProcedurePage
          method={selectedMethod}
          workflow={selectedWorkflow}
          stepNo={stepNo}
          onSelectStep={setStepNo}
          forms={data.forms}
          missingFieldCount={missingFieldCount}
          onOpenForm={setActiveForm}
        />
      )}

      {route === 'input' && (
        <InputPage
          fields={data.fields}
          repeatGroups={data.repeatGroups}
          values={values}
          onChangeValue={handleChangeValue}
          repeats={repeats}
          onChangeRepeatCell={handleChangeRepeatCell}
          onImport={handleImport}
        />
      )}

      {route === 'forms' && (
        <FormsPage
          forms={data.forms}
          query={searchQuery}
          onQueryChange={setSearchQuery}
          missingFieldCount={missingFieldCount}
          onOpenForm={setActiveForm}
        />
      )}

      {route === 'export' && (
        <ExportPage values={values} repeats={repeats} forms={data.forms} fields={data.fields} />
      )}

      {route === 'guide' && <GuidePage />}

      {activeForm && (
        <DocumentPreviewModal
          form={activeForm}
          fields={data.fields}
          values={values}
          repeats={repeats}
          missingFieldIds={missingFieldIds(activeForm)}
          onClose={() => setActiveForm(null)}
        />
      )}
      </div>

      <Footer />
    </div>
  )
}
