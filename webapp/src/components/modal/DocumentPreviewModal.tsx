import { useEffect, useMemo, useRef, useState } from 'react'
import { X } from 'lucide-react'
import type { FieldDef, FieldValues, FormRecord, RepeatValues } from '../../types'
import { collectTokenValues, exportHwpxDocument } from '../../lib/hwpxExport'
import {
  loadHwpxPreview,
  type HwpxPreviewBlock,
  type HwpxPreviewParagraph,
  type HwpxPreviewTable,
} from '../../lib/hwpxPreview'
import { exportElementToPdf } from '../../lib/pdf'
import './modal.css'

interface DocumentPreviewModalProps {
  form: FormRecord
  fields: FieldDef[]
  values: FieldValues
  repeats: RepeatValues
  missingFieldIds: string[]
  onClose: () => void
}

function labelFor(fields: FieldDef[], fieldId: string): string {
  return fields.find((field) => field.fieldId === fieldId)?.label ?? fieldId
}

/** Field ID 목록 중 이 form에서 실제로 보여줄 대상만 남긴다(개인정보 보호 서식 경계 포함). */
function visibleFieldIds(form: FormRecord): string[] {
  return form.privacyAllowedFieldIds ?? form.inputFieldIds
}

function repeatGroupsUsedBy(form: FormRecord, fields: FieldDef[]): string[] {
  const allowedIds = new Set(visibleFieldIds(form))
  const groups = new Set<string>()
  for (const field of fields) {
    if (field.repeatGroup && allowedIds.has(field.fieldId)) groups.add(field.repeatGroup)
  }
  return [...groups]
}

function PreviewTable({ table }: { table: HwpxPreviewTable }) {
  return (
    <table className="paper-repeat-table paper-source-table">
      <tbody>
        {table.rows.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {row.map((cell, cellIndex) => (
              <td colSpan={cell.colSpan} rowSpan={cell.rowSpan} key={cellIndex}>
                {cell.paragraphs.map((paragraph, paragraphIndex) => (
                  <PreviewParagraph paragraph={paragraph} key={paragraphIndex} />
                ))}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

function PreviewParagraph({ paragraph }: { paragraph: HwpxPreviewParagraph }) {
  if (paragraph.parts.length === 0) return <div className="paper-source-paragraph">　</div>
  return (
    <div className="paper-source-paragraph">
      {paragraph.parts.map((part, index) =>
        part.kind === 'table' ? (
          <PreviewTable table={part} key={`table-${index}`} />
        ) : part.kind === 'paragraph' ? (
          <PreviewParagraph paragraph={part} key={`paragraph-${index}`} />
        ) : (
          <span key={`${part.kind}-${index}`}>{part.text}</span>
        )
      )}
    </div>
  )
}

function PreviewBlock({ block }: { block: HwpxPreviewBlock }) {
  return block.kind === 'paragraph' ? (
    <PreviewParagraph paragraph={block} />
  ) : (
    <PreviewTable table={block} />
  )
}

export function DocumentPreviewModal({
  form,
  fields,
  values,
  repeats,
  missingFieldIds,
  onClose,
}: DocumentPreviewModalProps) {
  const paperRef = useRef<HTMLDivElement>(null)
  const closeButtonRef = useRef<HTMLButtonElement>(null)
  const onCloseRef = useRef(onClose)
  const [busy, setBusy] = useState(false)
  const [statusMessage, setStatusMessage] = useState<string | null>(null)
  const [source, setSource] = useState<{ key: string; blocks: HwpxPreviewBlock[] } | null>(null)
  const [sourceFailure, setSourceFailure] = useState<{ key: string; message: string } | null>(null)

  useEffect(() => {
    onCloseRef.current = onClose
  }, [onClose])

  useEffect(() => {
    const trigger = document.activeElement instanceof HTMLElement ? document.activeElement : null
    closeButtonRef.current?.focus()
    function handleKeyDown(event: KeyboardEvent): void {
      if (event.key === 'Escape') onCloseRef.current()
    }
    document.addEventListener('keydown', handleKeyDown)
    return () => {
      document.removeEventListener('keydown', handleKeyDown)
      if (trigger?.isConnected) trigger.focus()
    }
  }, [])

  const tokenValues = useMemo(() => collectTokenValues(form, fields, values).tokenValues, [fields, form, values])
  const previewKey = form.hwpxTemplate ? `${form.hwpxTemplate}:${JSON.stringify(tokenValues)}` : null
  const sourceBlocks = source?.key === previewKey ? source.blocks : null
  const sourceError = sourceBlocks === null && sourceFailure?.key === previewKey ? sourceFailure.message : null
  const sourcePending = Boolean(form.hwpxTemplate) && sourceBlocks === null && !sourceError
  const pdfOutputBlocked = Boolean(form.hwpxTemplate) && sourceBlocks === null
  const pdfOutputMessage = sourcePending
    ? '원문 HWPX를 불러오는 동안 PDF 출력은 사용할 수 없습니다.'
    : sourceError
      ? '원문 HWPX를 불러오지 못해 PDF 출력할 수 없습니다.'
      : null

  useEffect(() => {
    let cancelled = false
    if (!form.hwpxTemplate || !previewKey) return

    void loadHwpxPreview(form.hwpxTemplate, tokenValues)
      .then((blocks) => {
        if (!cancelled) setSource({ key: previewKey, blocks })
      })
      .catch((caught: unknown) => {
        if (!cancelled) {
          setSourceFailure({
            key: previewKey,
            message: caught instanceof Error ? caught.message : 'HWPX 원문을 표시하지 못했습니다.',
          })
        }
      })
    return () => {
      cancelled = true
    }
  }, [form.hwpxTemplate, previewKey, tokenValues])

  const nonRepeatFieldIds = visibleFieldIds(form).filter(
    (fieldId) => !fields.find((field) => field.fieldId === fieldId)?.repeatGroup
  )
  const usedRepeatGroups = repeatGroupsUsedBy(form, fields)
  const canOutputHwpx = form.outputMethod !== 'none' && Boolean(form.hwpxTemplate)
  const canOutputPdf = form.outputMethod === 'both'

  async function handleHwpxExport(): Promise<void> {
    setBusy(true)
    setStatusMessage(null)
    try {
      const result = await exportHwpxDocument(form, fields, values)
      if (result.missingTokens.length > 0) {
        setStatusMessage(`다음 값을 먼저 입력해야 HWPX를 생성할 수 있습니다: ${result.missingTokens.join(', ')}`)
      } else {
        setStatusMessage('HWPX 파일을 내려받았습니다.')
      }
    } catch (caught: unknown) {
      setStatusMessage(caught instanceof Error ? caught.message : 'HWPX 생성 중 오류가 발생했습니다.')
    } finally {
      setBusy(false)
    }
  }

  async function handlePdfExport(): Promise<void> {
    if (!paperRef.current) return
    setBusy(true)
    setStatusMessage(null)
    try {
      await exportElementToPdf(paperRef.current, `${form.formId}_${form.title}.pdf`)
      setStatusMessage('PDF 파일을 내려받았습니다.')
    } catch (caught: unknown) {
      setStatusMessage(caught instanceof Error ? caught.message : 'PDF 생성 중 오류가 발생했습니다.')
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="modal" onClick={onClose}>
      <section
        role="dialog"
        aria-modal="true"
        aria-labelledby="document-preview-title"
        onClick={(event) => event.stopPropagation()}
        style={{ boxSizing: 'border-box', maxHeight: '90vh', overflowY: 'auto' }}
      >
        <button className="close" ref={closeButtonRef} onClick={onClose} aria-label="닫기">
          <X />
        </button>
        <h2 id="document-preview-title">
          {form.formId}　{form.title}
        </h2>

        {form.privacyProtected && (
          <p className="modal-privacy-note">
            개인정보 보호 서식입니다. 학생·학부모 등 개인 응답란은 빈 양식으로 유지되며, 이 화면에서
            입력·저장·자동반영하지 않습니다.
          </p>
        )}

        <div className="paper" ref={paperRef} aria-busy={sourcePending}>
          {sourcePending && <p>원문 HWPX를 불러오는 중입니다.</p>}
          {sourceError && <p className="modal-status">{sourceError}</p>}
          {sourceBlocks?.map((block, index) => <PreviewBlock block={block} key={index} />)}

          {!form.hwpxTemplate && (
            <>
              <h1>{form.title}</h1>
              <dl className="paper-fields">
                {nonRepeatFieldIds.map((fieldId) => (
                  <div className="paper-field" key={fieldId}>
                    <dt>{labelFor(fields, fieldId)}</dt>
                    <dd>{values[fieldId] || '　'}</dd>
                  </div>
                ))}
              </dl>

              {usedRepeatGroups.map((group) => {
                const groupFields = fields.filter((field) => field.repeatGroup === group)
                const rows = (repeats[group] ?? []).filter((row) =>
                  groupFields.some((field) => (row[field.fieldId] ?? '') !== '')
                )
                if (rows.length === 0) return null
                return (
                  <table className="paper-repeat-table" key={group}>
                    <caption>{group}</caption>
                    <thead>
                      <tr>
                        {groupFields.map((field) => (
                          <th key={field.fieldId}>{field.label}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                          {groupFields.map((field) => (
                            <td key={field.fieldId}>{row[field.fieldId] || ''}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )
              })}
            </>
          )}
        </div>

        <p>누락 입력값: {missingFieldIds.map((id) => labelFor(fields, id)).join(', ') || '없음'}</p>

        <div className="modal-actions">
          {canOutputHwpx && (
            <button className="primary" disabled={busy} onClick={handleHwpxExport}>
              HWPX 작성
            </button>
          )}
          {canOutputPdf && (
            <button
              aria-describedby={pdfOutputMessage ? 'pdf-output-status' : undefined}
              disabled={busy || pdfOutputBlocked}
              onClick={handlePdfExport}
            >
              PDF 출력
            </button>
          )}
          {!canOutputHwpx && !canOutputPdf && (
            <span className="modal-not-implemented">이 서식은 아직 자동 출력을 지원하지 않습니다.</span>
          )}
        </div>
        {pdfOutputMessage && <p className="modal-status" id="pdf-output-status" role="status">{pdfOutputMessage}</p>}
        {statusMessage && <p className="modal-status">{statusMessage}</p>}
      </section>
    </div>
  )
}
