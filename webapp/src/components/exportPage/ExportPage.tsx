import { useState } from 'react'
import type { FieldDef, FieldValues, FormRecord, RepeatValues } from '../../types'
import { downloadWorkbook } from '../../lib/workbook'
import { exportAllHwpxAsZip } from '../../lib/hwpxBulkExport'
import './exportPage.css'

interface ExportPageProps {
  values: FieldValues
  repeats: RepeatValues
  forms: FormRecord[]
  fields: FieldDef[]
}

function countReadyForms(forms: FormRecord[], missingFieldCount: (form: FormRecord) => number): number {
  return forms.filter(
    (form) => form.implementationStatus !== 'notImplemented' && missingFieldCount(form) === 0
  ).length
}

interface ExportSummaryProps {
  forms: FormRecord[]
  missingFieldCount: (form: FormRecord) => number
}

function ExportSummary({ forms, missingFieldCount }: ExportSummaryProps) {
  const implemented = forms.filter((form) => form.implementationStatus !== 'notImplemented').length
  const ready = countReadyForms(forms, missingFieldCount)

  return (
    <ul className="export-summary">
      <li>
        전체 서식 <b>{forms.length}</b>종
      </li>
      <li>
        구현된 서식 <b>{implemented}</b>종
      </li>
      <li>
        지금 바로 작성 가능한 서식 <b>{ready}</b>종
      </li>
    </ul>
  )
}

export function ExportPage({ values, repeats, forms, fields }: ExportPageProps) {
  const [workbookError, setWorkbookError] = useState<string | null>(null)
  const [bulkBusy, setBulkBusy] = useState(false)
  const [bulkMessage, setBulkMessage] = useState<string | null>(null)

  function missingFieldCount(form: FormRecord): number {
    return form.requiredFieldIds.filter((fieldId) => !values[fieldId]).length
  }

  async function handleDownloadWorkbook(): Promise<void> {
    try {
      setWorkbookError(null)
      await downloadWorkbook(values, repeats)
    } catch (caught: unknown) {
      setWorkbookError(caught instanceof Error ? caught.message : '엑셀 파일을 만들지 못했습니다.')
    }
  }

  async function handleBulkHwpxExport(): Promise<void> {
    setBulkBusy(true)
    setBulkMessage(null)
    try {
      const result = await exportAllHwpxAsZip(forms, fields, values)
      const skippedNote =
        result.skipped.length > 0
          ? ` (${result.skipped.length}종 제외: ${result.skipped.map((item) => item.formId).join(', ')})`
          : ''
      setBulkMessage(
        result.succeeded.length > 0
          ? `${result.succeeded.length}종을 ZIP으로 내려받았습니다.${skippedNote}`
          : `내려받은 서식이 없습니다. 기초자료입력에서 필수값을 먼저 채워 주세요.${skippedNote}`
      )
    } catch (caught: unknown) {
      setBulkMessage(caught instanceof Error ? caught.message : '일괄 ZIP 생성 중 오류가 발생했습니다.')
    } finally {
      setBulkBusy(false)
    }
  }

  return (
    <main className="page">
      <h1>내보내기</h1>
      <p className="lede">입력한 기초자료를 엑셀로 백업하거나, 완성된 서식을 한 번에 내려받을 수 있습니다.</p>
      <div className="privacy">
        내려받은 엑셀·ZIP 파일은 이 브라우저에서만 만들어지며 서버로 전송되지 않습니다. 파일 보관 책임은
        담당자에게 있습니다.
      </div>

      <section className="export-panel">
        <h2>기초자료 백업</h2>
        <p>
          공통·사업·문서별 정보와 품목·업체·위원·평가 반복그룹 전체를 하나의 엑셀 파일(.xlsx)로
          내려받습니다. 같은 파일을 기초자료입력 화면의 "엑셀에서 불러오기"로 다시 불러올 수 있습니다.
        </p>
        <button className="primary" onClick={() => void handleDownloadWorkbook()}>
          기초자료 엑셀 내려받기
        </button>
        {workbookError && <p className="field-input-error">{workbookError}</p>}
      </section>

      <section className="export-panel">
        <h2>서식 일괄 출력(ZIP)</h2>
        <p>
          HWPX 템플릿이 있는 서식 중 필수 토큰 값이 채워진 서식만 모아 ZIP 하나로 내려받습니다. 값이
          비어 있는 서식은 건너뛰고 목록으로 안내합니다.
        </p>
        <button disabled={bulkBusy} onClick={() => void handleBulkHwpxExport()}>
          {bulkBusy ? 'ZIP 생성 중…' : '완성된 서식 전체 ZIP으로 내려받기'}
        </button>
        {bulkMessage && <p className="export-bulk-message">{bulkMessage}</p>}
      </section>

      <section className="export-panel">
        <h2>서식별 출력 현황</h2>
        <p>서식 하나씩 다시 확인하며 출력하려면 서식함에서 서식을 열어 개별로 진행합니다.</p>
        <ExportSummary forms={forms} missingFieldCount={missingFieldCount} />
      </section>
    </main>
  )
}
