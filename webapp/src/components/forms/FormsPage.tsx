import { Search } from 'lucide-react'
import { formStatusOf, summarizeLabels } from '../../lib/formReadiness'
import type { FieldDef, FormRecord } from '../../types'
import './forms.css'

interface FormsPageProps {
  forms: FormRecord[]
  query: string
  onQueryChange: (query: string) => void
  missingFieldsOf: (form: FormRecord) => FieldDef[]
  onOpenForm: (form: FormRecord) => void
}

export function FormsPage({ forms, query, onQueryChange, missingFieldsOf, onOpenForm }: FormsPageProps) {
  const filtered = forms.filter((form) =>
    (form.formId + form.title).toLowerCase().includes(query.toLowerCase())
  )

  return (
    <main className="page">
      <h1>서식함</h1>
      <label className="search">
        <Search size={20} />
        <input
          placeholder="서식명 또는 번호 검색"
          value={query}
          onChange={(event) => onQueryChange(event.target.value)}
        />
      </label>

      <div className="forms">
        {filtered.map((form) => {
          const missing = missingFieldsOf(form)
          const status = formStatusOf(form, missing)
          return (
            <button className="form" key={form.formId} onClick={() => onOpenForm(form)}>
              <span className="form-head">
                <b>{form.formId}</b>
                <small>매뉴얼 {form.manualNo}</small>
              </span>
              <h2>{form.title}</h2>
              <p>{form.author}</p>
              {form.privacyProtected && <span className="form-badge">개인정보 보호 서식</span>}
              <span className="form-status">
                <em className={`status-${status.kind}`}>{status.text}</em>
                {status.kind === 'missing' && (
                  <span className="form-missing">
                    부족: {summarizeLabels(missing.map((field) => field.label))}
                  </span>
                )}
              </span>
            </button>
          )
        })}
      </div>
    </main>
  )
}
