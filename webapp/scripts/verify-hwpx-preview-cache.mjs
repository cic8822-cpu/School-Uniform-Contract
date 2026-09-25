import assert from 'node:assert/strict'
import { fileURLToPath } from 'node:url'
import { createServer } from 'vite'

const root = fileURLToPath(new URL('..', import.meta.url))
const server = await createServer({
  root,
  appType: 'custom',
  server: { middlewareMode: true },
})
const originalFetch = globalThis.fetch
let fetchCalls = 0

try {
  const { loadHwpxPreview } = await server.ssrLoadModule('/src/lib/hwpxPreview.ts')
  globalThis.fetch = async () => {
    fetchCalls += 1
    throw new Error('일시적 HWPX 원문 요청 실패')
  }

  const templateName = `cache-retry-${Date.now()}.hwpx`
  await assert.rejects(() => loadHwpxPreview(templateName, {}))
  await assert.rejects(() => loadHwpxPreview(templateName, {}))
  assert.equal(fetchCalls, 2, '실패한 HWPX 원문 요청은 캐시에서 제거되어 재시도해야 합니다.')
  console.log(JSON.stringify({ status: 'PASS', fetchCalls }))
} finally {
  globalThis.fetch = originalFetch
  await server.close()
}
