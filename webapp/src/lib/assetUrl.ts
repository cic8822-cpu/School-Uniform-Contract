/**
 * 배포 방식(정적 서버 vs `file://` 직접 열기)에 상관없이 동작하도록
 * Vite의 `import.meta.env.BASE_URL`(vite.config.ts의 `base: './'`) 기준으로
 * 정적 자산 경로를 만든다. 항상 루트("/")가 아닌 상대경로를 반환한다.
 */
export function assetUrl(relativePath: string): string {
  const base = import.meta.env.BASE_URL
  const normalizedBase = base.endsWith('/') ? base : `${base}/`
  const normalizedPath = relativePath.startsWith('/') ? relativePath.slice(1) : relativePath
  return `${normalizedBase}${normalizedPath}`
}
