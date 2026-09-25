import { Search } from 'lucide-react'
import type { FormRecord } from '../../types'
import './forms.css'

interface FormsPageProps {
  forms: FormRecord[]
  query: string
  onQueryChange: (query: string) => void
  missingFieldCount: (form: FormRecord) => number
  onOpenForm: (form: FormRecord) => void
}

export function FormsPage({ forms, query, onQueryChange, missingFieldCount, onOpenForm }: FormsPageProps) {
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
        {filtered.map((form) => (
          <button className="form" key={form.formId} onClick={() => onOpenForm(form)}>
            <b>{form.formId}</b>
            <small>매뉴얼 {form.manualNo}</small>
            <h2>{form.title}</h2>
            <p>{form.author}</p>
            {form.privacyProtected && <span className="form-badge">개인정보 보호 서식</span>}
            <em>
              {form.implementationStatus === 'notImplemented'
                ? '○ 미구현'
                : missingFieldCount(form) > 0
                  ? '▲ 기초자료 필요'
                  : '● 작성가능'}
            </em>
          </button>
        ))}
      </div>
    </main>
  )
}
