# GREEN — 원문 미리보기 차단 결함 브라우저 QA

- 환경: `webapp` Vite `http://127.0.0.1:4174/`, 로컬 프록시 `4175`(F-008 HWPX 응답 1.2초 지연)·`4176`(F-008 HWPX 응답 503), Codex 인앱 브라우저의 Playwright DOM 검사. Browser 플러그인은 이 세션에 없음.
- F-008 원문 로딩 시나리오: `4175 → 서식함 → F-008` 직후.
  - 관찰값: `PDF 출력.disabled = true`, `.paper[aria-busy] = "true"`, `원문 HWPX를 불러오는 동안 PDF 출력은 사용할 수 없습니다.` 표시.
  - 1.8초 뒤: `PDF 출력.disabled = false`, `.paper[aria-busy] = "false"`, 원문 표 `3`개.
- 원문 오류 시나리오: `4176 → 서식함 → F-008`.
  - 관찰값: `HWPX 원문을 불러오지 못했습니다.`, `원문 HWPX를 불러오지 못해 PDF 출력할 수 없습니다.`, `PDF 출력.disabled = true`, `.paper[aria-busy] = "false"`.
  - 판정: 원문 로딩 전·오류 시 PDF 성공 저장 경로가 UI에서 차단됨.
- Esc 포커스 시나리오: 로드 완료한 F-008 모달에서 `Escape`.
  - 관찰값: `document.activeElement = BUTTON.form`, 텍스트가 `F-008 … 교복 사양서`.
- 비템플릿 시나리오: `서식함 → F-026`.
  - 관찰값: `.paper[aria-busy] = "false"`, 원문 로딩 문구 없음.
- 개인정보 경계 시나리오: `서식함 → F-050`(개인식별·연락처·서명·설문 원자료 미입력).
  - 관찰값: 개인정보 보호 안내 존재, 모달의 `input, textarea, select = 0`, 원문 표 `3`개, `.paper[aria-busy] = "false"`.
  - 판정: F-050 개인 응답 입력·저장·자동반영 경계 유지.
- 콘솔: 최종 오류/경고 `[]`.
