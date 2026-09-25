import { downloadBlob } from './download'

// 허용/금지 토큰 목록은 scripts/hwpx_token_fill.ps1 39~40행, 그리고
// _workspace/05_web/data/template-mapping.md §1과 문자 단위로 정확히 일치해야
// 한다. 개인정보 경계와 직결되므로 이 두 값은 그 근거를 재확인하지 않고는
// 임의로 바꾸지 않는다(웹앱_trd.md §4 Critical).
const allowed = new Set(['학교명', '학년도', '문서번호', '발행일', '제목', '수신기관', '시행일', '공개구분', '교장명', '관련문서', '제출기한'])
const forbidden = /성명|이름|담당자|전화|휴대|주소|이메일|서명|직인|계좌|주민|생년|학번|설문/

/**
 * HWPX 템플릿을 내려받아 {{토큰}}을 값으로 치환한 결과를 Blob으로 돌려준다.
 * 다운로드를 트리거하지 않으므로, 결재란 라벨 채움(hwpxFooterFill)처럼 추가
 * 후처리가 필요한 호출부에서 이어서 사용할 수 있다.
 */
export async function buildTokenFilledHwpx(templateUrl: string, values: Record<string, string>): Promise<Blob> {
  // jszip은 HWPX 출력을 실제로 누를 때만 필요하므로 동적 import로 분리한다
  // (web/performance.md 번들 예산).
  const { default: JSZip } = await import('jszip')
  const response = await fetch(templateUrl)
  if (!response.ok) throw new Error('HWPX 템플릿을 불러오지 못했습니다.')
  const zip = await JSZip.loadAsync(await response.arrayBuffer())
  const entries = Object.keys(zip.files).filter((name) => /^Contents\/section.*\.xml$/.test(name))
  if (!entries.length) throw new Error('HWPX 본문 XML이 없습니다.')
  for (const name of entries) {
    const file = zip.file(name)
    if (!file) continue
    const xml = await file.async('string')
    const replaced = xml.replace(/\{\{([^{}]+)\}\}/g, (_match, key: string) => {
      if (!allowed.has(key) || forbidden.test(key)) throw new Error(`허용되지 않은 HWPX 토큰: ${key}`)
      const value = values[key]
      if (value === undefined || value === '') throw new Error(`필수 HWPX 값이 비어 있습니다: ${key}`)
      return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    })
    if (/\{\{[^{}]+\}\}/.test(replaced)) throw new Error(`치환되지 않은 토큰이 남았습니다: ${name}`)
    zip.file(name, replaced)
  }
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' })
}

/** 토큰을 채운 HWPX를 즉시 파일로 내려받는다(후처리가 필요 없는 단순 호출용). */
export async function createHwpx(templateUrl: string, values: Record<string, string>, outputName: string): Promise<void> {
  const blob = await buildTokenFilledHwpx(templateUrl, values)
  downloadBlob(blob, outputName)
}
