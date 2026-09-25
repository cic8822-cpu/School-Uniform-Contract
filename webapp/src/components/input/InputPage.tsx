import { useRef, useState } from 'react'
import type { FieldDef, FieldValues, RepeatGroupDef, RepeatValues } from '../../types'
import { CommonFieldsTab } from './CommonFieldsTab'
import { RepeatGroupTab } from './RepeatGroupTab'
import { computeEvaluationTotal, computeItemAmount, computeItemAmountTotal } from '../../lib/fieldCompute'
import { formatCurrency } from '../../lib/format'
import { downloadWorkbook, readWorkbook } from '../../lib/workbook'
import './input.css'

interface InputPageProps {
  fields: FieldDef[]
  repeatGroups: RepeatGroupDef[]
  values: FieldValues
  onChangeValue: (fieldId: string, value: string) => void
  repeats: RepeatValues
  onChangeRepeatCell: (group: string, rowIndex: number, fieldId: string, value: string) => void
  onImport: (result: { values: FieldValues; repeats: RepeatValues }) => void
}

type TabId = '공통' | '사업' | '문서별' | '품목' | '업체' | '위원' | '평가'
const TAB_ORDER: TabId[] = ['공통', '사업', '문서별', '품목', '업체', '위원', '평가']

export function InputPage({
  fields,
  repeatGroups,
  values,
  onChangeValue,
  repeats,
  onChangeRepeatCell,
  onImport,
}: InputPageProps) {
  const [activeTab, setActiveTab] = useState<TabId>('공통')
  const [importError, setImportError] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const fieldsByGroup = (group: string) => fields.filter((field) => field.group === group)
  const fieldsByRepeatGroup = (group: string) => fields.filter((field) => field.repeatGroup === group)
  const repeatGroupDef = (group: string) => repeatGroups.find((candidate) => candidate.group === group)

  async function handleDownloadWorkbook(): Promise<void> {
    try {
      setImportError(null)
      await downloadWorkbook(values, repeats)
    } catch (caught: unknown) {
      setImportError(caught instanceof Error ? caught.message : '엑셀 파일을 만들지 못했습니다.')
    }
  }

  async function handleFileSelected(event: React.ChangeEvent<HTMLInputElement>): Promise<void> {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    try {
      setImportError(null)
      const result = await readWorkbook(file)
      onImport(result)
    } catch (caught: unknown) {
      setImportError(caught instanceof Error ? caught.message : '엑셀 파일을 읽지 못했습니다.')
    }
  }

  return (
    <main className="page">
      <h1>기초자료입력</h1>
      <p className="lede">한 번 입력하면 전체 서식에 자동으로 반영됩니다.</p>
      <div className="privacy">입력한 업무자료는 서버 DB에 저장하지 않습니다. 이 브라우저에서만 사용합니다.</div>

      <div className="input-toolbar">
        <button className="primary" onClick={() => void handleDownloadWorkbook()}>
          엑셀로 내려받기
        </button>
        <button onClick={() => fileInputRef.current?.click()}>엑셀에서 불러오기</button>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx"
          className="input-file-hidden"
          onChange={handleFileSelected}
        />
      </div>
      {importError && <p className="field-input-error">{importError}</p>}

      <div className="input-tabs" role="tablist">
        {TAB_ORDER.map((tab) => (
          <button
            key={tab}
            role="tab"
            aria-selected={activeTab === tab}
            className={activeTab === tab ? 'input-tab-button active' : 'input-tab-button'}
            onClick={() => setActiveTab(tab)}
          >
            {tab}
          </button>
        ))}
      </div>

      {activeTab === '공통' && (
        <CommonFieldsTab
          title="① 공통 정보"
          fields={fieldsByGroup('공통')}
          values={values}
          onChange={onChangeValue}
        />
      )}
      {activeTab === '사업' && (
        <CommonFieldsTab
          title="② 사업 정보"
          fields={fieldsByGroup('사업')}
          values={values}
          onChange={onChangeValue}
        />
      )}
      {activeTab === '문서별' && (
        <CommonFieldsTab
          title="③ 문서별 정보"
          fields={fieldsByGroup('문서별')}
          values={values}
          onChange={onChangeValue}
        />
      )}

      {activeTab === '품목' && (
        <RepeatGroupTab
          title={`④ 품목 (최대 ${repeatGroupDef('품목')?.excelRows ?? ''})`}
          fields={fieldsByRepeatGroup('품목')}
          rows={repeats['품목'] ?? []}
          onChangeCell={(rowIndex, fieldId, value) => onChangeRepeatCell('품목', rowIndex, fieldId, value)}
          computedColumn={{
            label: 'K-01 품목별 금액',
            compute: computeItemAmount,
            format: formatCurrency,
          }}
          totalRow={{
            label: 'K-02 합계금액',
            compute: computeItemAmountTotal,
            format: formatCurrency,
          }}
        />
      )}
      {activeTab === '업체' && (
        <RepeatGroupTab
          title="⑤ 업체"
          fields={fieldsByRepeatGroup('업체')}
          rows={repeats['업체'] ?? []}
          onChangeCell={(rowIndex, fieldId, value) => onChangeRepeatCell('업체', rowIndex, fieldId, value)}
        />
      )}
      {activeTab === '위원' && (
        <RepeatGroupTab
          title="⑥ 위원"
          fields={fieldsByRepeatGroup('위원')}
          rows={repeats['위원'] ?? []}
          onChangeCell={(rowIndex, fieldId, value) => onChangeRepeatCell('위원', rowIndex, fieldId, value)}
        />
      )}
      {activeTab === '평가' && (
        <RepeatGroupTab
          title="⑦ 평가"
          fields={fieldsByRepeatGroup('평가')}
          rows={repeats['평가'] ?? []}
          onChangeCell={(rowIndex, fieldId, value) => onChangeRepeatCell('평가', rowIndex, fieldId, value)}
          totalRow={{
            label: 'K-03 평가 총점',
            compute: computeEvaluationTotal,
            format: (amount) => `${amount}점`,
          }}
          note="평가항목·배점 열은 Field 마스터에 별도 ID가 없어 Excel 원본에서만 관리되며, 이 화면에는 평가 점수(R-07)만 입력합니다."
        />
      )}
    </main>
  )
}
