import { useEffect, useState } from 'react'
import type { AppData, ContractMethodsFile, FieldsFile, FormsFile, WorkflowsFile } from '../types'
import { assetUrl } from '../lib/assetUrl'

interface AppDataState {
  data: AppData | null
  loading: boolean
  error: string | null
}

/**
 * 마스터 데이터 4종(contract-methods·workflows·forms·fields)을 정적 JSON에서
 * 불러온다. 이 JSON은 서버가 아니라 빌드에 포함된 정적 자산(`public/data/`)이며,
 * 여기서 읽은 값이 다시 네트워크로 나가는 곳은 없다(ADR-002).
 */
export function useAppData(): AppDataState {
  const [data, setData] = useState<AppData | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false

    async function load(): Promise<void> {
      try {
        const [contractMethods, workflows, forms, fields] = await Promise.all([
          fetchJson<ContractMethodsFile>(assetUrl('data/contract-methods.json')),
          fetchJson<WorkflowsFile>(assetUrl('data/workflows.json')),
          fetchJson<FormsFile>(assetUrl('data/forms.json')),
          fetchJson<FieldsFile>(assetUrl('data/fields.json')),
        ])
        if (cancelled) return
        setData({
          methods: contractMethods.contractMethods,
          workflows: workflows.workflows,
          standardPhases: workflows.standardPhases.phases,
          forms: forms.forms,
          fields: fields.fields,
          repeatGroups: fields.repeatGroups,
          protectedForms: fields.protectedForms,
          globalForbidden: fields.globalForbidden,
        })
      } catch (caught: unknown) {
        if (cancelled) return
        setError(caught instanceof Error ? caught.message : '데이터를 불러오지 못했습니다.')
      } finally {
        if (!cancelled) setLoading(false)
      }
    }

    void load()
    return () => {
      cancelled = true
    }
  }, [])

  return { data, loading, error }
}

async function fetchJson<T>(url: string): Promise<T> {
  const response = await fetch(url)
  if (!response.ok) throw new Error(`${url}을 불러오지 못했습니다. (${response.status})`)
  return (await response.json()) as T
}
