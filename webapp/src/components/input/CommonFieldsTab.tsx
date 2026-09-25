import type { FieldDef, FieldValues } from '../../types'
import { validateFieldValue } from '../../lib/validation'
import { FieldInput } from './FieldInput'

interface CommonFieldsTabProps {
  title: string
  fields: FieldDef[]
  values: FieldValues
  onChange: (fieldId: string, value: string) => void
}

/** 공통·사업·문서별처럼 반복되지 않는 Field 그룹을 2열 그리드로 렌더링한다. */
export function CommonFieldsTab({ title, fields, values, onChange }: CommonFieldsTabProps) {
  return (
    <section className="input-tab">
      <h2>{title}</h2>
      <div className="input-grid">
        {fields.map((field) => {
          const value = values[field.fieldId] ?? ''
          return (
            <FieldInput
              key={field.fieldId}
              field={field}
              value={value}
              onChange={(nextValue) => onChange(field.fieldId, nextValue)}
              error={validateFieldValue(field, value)}
            />
          )
        })}
      </div>
    </section>
  )
}
