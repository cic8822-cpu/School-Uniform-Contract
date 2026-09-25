import type { FieldDef, RepeatRow } from '../../types'
import { extractEnumOptions } from '../../lib/validation'

export interface ComputedColumn {
  label: string
  compute: (row: RepeatRow) => number
  format: (amount: number) => string
}

export interface TotalRow {
  label: string
  compute: (rows: RepeatRow[]) => number
  format: (amount: number) => string
}

interface RepeatGroupTabProps {
  title: string
  fields: FieldDef[]
  rows: RepeatRow[]
  onChangeCell: (rowIndex: number, fieldId: string, value: string) => void
  computedColumn?: ComputedColumn
  totalRow?: TotalRow
  note?: string
}

/** 품목·업체·위원·평가 4개 반복그룹을 excelRows 행수만큼 표로 렌더링한다. */
export function RepeatGroupTab({
  title,
  fields,
  rows,
  onChangeCell,
  computedColumn,
  totalRow,
  note,
}: RepeatGroupTabProps) {
  return (
    <section className="input-tab">
      <h2>{title}</h2>
      {note && <p className="repeat-tab-note">{note}</p>}

      <div className="repeat-table-scroll">
        <table className="repeat-table">
          <thead>
            <tr>
              <th>행</th>
              {fields.map((field) => (
                <th key={field.fieldId}>{field.label}</th>
              ))}
              {computedColumn && <th>{computedColumn.label}</th>}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rowIndex) => (
              <tr key={rowIndex}>
                <td className="repeat-table-row-no">{rowIndex + 1}</td>
                {fields.map((field) => (
                  <td key={field.fieldId}>
                    <RepeatCell
                      field={field}
                      value={row[field.fieldId] ?? ''}
                      onChange={(value) => onChangeCell(rowIndex, field.fieldId, value)}
                    />
                  </td>
                ))}
                {computedColumn && (
                  <td className="repeat-table-computed">
                    {computedColumn.format(computedColumn.compute(row))}
                  </td>
                )}
              </tr>
            ))}
          </tbody>
          {totalRow && (
            <tfoot>
              <tr>
                <td colSpan={fields.length + 1}>{totalRow.label}</td>
                <td className="repeat-table-computed">{totalRow.format(totalRow.compute(rows))}</td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </section>
  )
}

interface RepeatCellProps {
  field: FieldDef
  value: string
  onChange: (value: string) => void
}

function RepeatCell({ field, value, onChange }: RepeatCellProps) {
  const enumOptions = field.type === 'enum' ? extractEnumOptions(field.validation) : null
  if (enumOptions) {
    return (
      <select value={value} onChange={(event) => onChange(event.target.value)}>
        <option value="">선택 안 함</option>
        {enumOptions.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
    )
  }

  const inputType = field.type === 'integer' || field.type === 'currency' || field.type === 'number'
    ? 'number'
    : 'text'
  return <input type={inputType} value={value} onChange={(event) => onChange(event.target.value)} />
}
