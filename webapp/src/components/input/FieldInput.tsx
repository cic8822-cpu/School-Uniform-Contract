import type { FieldDef } from '../../types'
import { extractEnumOptions } from '../../lib/validation'

interface FieldInputProps {
  field: FieldDef
  value: string
  onChange: (value: string) => void
  error?: string | null
  compact?: boolean
}

const MULTILINE_TYPES = new Set<FieldDef['type']>(['text', 'stringArray'])
const NUMERIC_TYPES = new Set<FieldDef['type']>(['integer', 'currency', 'number', 'year'])

function inputTypeFor(field: FieldDef): string {
  if (field.type === 'date') return 'date'
  if (NUMERIC_TYPES.has(field.type)) return 'number'
  return 'text'
}

/** Field 정의 하나에 맞는 입력 컨트롤을 렌더링한다(텍스트·숫자·날짜·선택·여러줄). */
export function FieldInput({ field, value, onChange, error, compact }: FieldInputProps) {
  const enumOptions = field.type === 'enum' ? extractEnumOptions(field.validation) : null
  const isRequired = field.required === 'required'

  return (
    <label className={compact ? 'field-input field-input-compact' : 'field-input'}>
      <span className="field-input-label">
        {field.label}
        {isRequired && <span className="field-input-required">*</span>}
      </span>

      {enumOptions ? (
        <select value={value} onChange={(event) => onChange(event.target.value)}>
          <option value="">선택 안 함</option>
          {enumOptions.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
      ) : MULTILINE_TYPES.has(field.type) ? (
        <textarea rows={3} value={value} onChange={(event) => onChange(event.target.value)} />
      ) : (
        <input
          type={inputTypeFor(field)}
          value={value}
          onChange={(event) => onChange(event.target.value)}
        />
      )}

      <span className="field-input-hint">{field.validation}</span>
      {error && <span className="field-input-error">{error}</span>}
    </label>
  )
}
