# RED — 원문 미리보기 차단 결함

- 실행 환경: `webapp`의 Vite 개발 서버 `http://127.0.0.1:4174/`, Codex 인앱 브라우저, 2026-09-25.
- Esc 포커스 시나리오: `서식함 → F-008 → Esc`.
  - 관찰값: `document.activeElement.tagName = "BODY"`.
  - 판정: 트리거 버튼으로 포커스가 돌아가지 않아 FAIL.
- 템플릿 없음 상태 시나리오: `서식함 → F-026`.
  - 관찰값: `.paper[aria-busy] = "true"`, 로딩 문구 없음, PDF 버튼 없음.
  - 판정: HWPX를 요청하지 않는 서식이 영구 로딩 상태로 노출되어 FAIL.
- 원문 실패 후 재시도 시나리오: `node scripts/verify-hwpx-preview-cache.mjs`.
  - 관찰값: `AssertionError`, 실제 `fetchCalls = 1`, 기대값 2.
  - 판정: 거부된 Promise가 캐시에 남아 재시도하지 못해 FAIL. 원본 실행 로그: `hwpx-cache-retry.log`.
