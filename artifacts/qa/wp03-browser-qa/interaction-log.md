# WP03 browser action log

- Commit checked: `8f8fffeb4300b96853e760da87a8a51b9fdb6404`
- Server: `npm run dev -- --host 127.0.0.1 --port 4173` (PID recorded in QA session: 29916)
- Browser: Codex In-app Browser, URL `http://127.0.0.1:4173/`, measured viewport `1280x720`
- Input values used: 학교명 `테스트학교`, 학년도 `2026` (default), 문서 발행일 `2026-09-25`, 문서번호 `문서-001`; no personal data
- F-008 open: full paragraphs and all three source tables visible in dialog
- F-008 HWPX 작성: status paragraph `HWPX 파일을 내려받았습니다.` appeared
- F-050 open: privacy banner and blank response structure visible
- F-050 modal metrics: `scrollHeight=1527`, `clientHeight=648`, `overflowY=auto`; modal is scrollable
- Esc: dialog closed, but active element was `BODY` instead of the invoking F-050 button (focus restoration failure)
- Console query: `tab.dev.logs({levels:['error','warn']})` returned `[]`
- 320px: not executable in this browser surface; no viewport capability is advertised by the in-app browser. This is a blocker, not an inferred pass.

