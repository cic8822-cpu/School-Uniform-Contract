# WP03 최종 브라우저 수동 QA

- 기준 커밋: `3800c6e95583e98a67f7a6d0f9e7e1bfd83d93f7`
- 표면: Codex In-app Browser, `http://127.0.0.1:4176/`, Vite 개발 서버
- 입력: 학교명 `테스트학교`, 학년도 `2026`, 구매명 `교복 구매`, 품목명 `재킷`, 문서번호 `문서-001`, 발행일 `2026-09-25`. 개인정보는 입력하지 않음.
- 결론: **BLOCK** — 구현·정상 경로는 PASS했으나, 정적 로컬 HWPX 원문 fetch의 지연/실패 상태를 브라우저에서 유도할 수 있는 fault-injection 표면이 없어 해당 필수 적대 시나리오를 실행하지 못함. 추정 PASS로 처리하지 않음.

## surfaceEvidence

| 시나리오 | 기준 | 표면·정확한 호출 | 판정 | 증빙 |
|---|---|---|---|---|
| S-01 | F-008 선택→원문 전체 문단/표 | `tab.playwright.getByRole('button',{name:/^F-008/}).click()` 후 `tab.playwright.domSnapshot()` 및 브라우저 스크린샷 | PASS | AR-01, AR-02 |
| S-02 | F-008 HWPX 작성 | F-008 미리보기에서 `HWPX 작성` 클릭, 상태 문구 확인 | PASS | AR-03 |
| S-03 | 지연 중 PDF disabled | F-008 재진입 직후 `PDF 출력.isEnabled()` 확인 | BLOCK | AR-04 |
| S-04 | 오류 중 PDF disabled | 원문 fetch 오류 유도 후 PDF 버튼 상태 확인 | BLOCK | AR-04 |
| S-05 | F-050 빈 양식·개인정보 0 | F-050 미리보기, `dialog input/textarea/select` 읽기, DOM·스크린샷 확인 | PASS | AR-05, AR-06 |
| S-06 | Esc 포커스 복귀 | F-050 호출 버튼 클릭→`dialog.press('Escape')`→`document.activeElement` 확인 | PASS | AR-07 |
| S-07 | 1280 반응형/스크롤 | viewport 1280×720, F-008 미리보기, `innerWidth/body.scrollWidth/dialog` 측정 | PASS | AR-08 |
| S-08 | 320 반응형/조작 가능 | viewport 320×720, F-008 미리보기, modal 내 가로·세로 스크롤과 레이아웃 측정 | PASS | AR-09 |
| S-09 | 브라우저 콘솔 | `tab.dev.logs({levels:['error','warn']})` at 320×720; 1280 흐름 포함 | PASS | AR-10 |
| S-10 | lint/build | `npm run lint`; `npm run build` | PASS | AR-11, AR-12 |

## adversarialCases

| 시나리오 | 기준 | 적대 클래스 | 기대 동작 | 판정 | 증빙 |
|---|---|---|---|---|---|
| A-01 | F-008 | 원문이 입력 요약으로 축약 | 전체 설명 문단과 3개 표가 렌더링 | PASS | AR-01 |
| A-02 | F-008 | 필수값 일부 누락 후 HWPX 작성 | 생성하지 않고 누락값을 안내 | PASS | AR-03 |
| A-03 | F-008 | HWPX 생성 busy 중 중복 클릭 | busy 동안 HWPX/PDF 버튼 비활성 | PASS | AR-13 |
| A-04 | PDF | 원문 fetch 지연 | 지연 중 PDF disabled 및 안내 | BLOCK | AR-04 |
| A-05 | PDF | 원문 fetch 오류 | 오류 중 PDF disabled 및 안내 | BLOCK | AR-04 |
| A-06 | F-050 | 개인정보 응답값 자동 유입 | 응답 셀·기타 의견·개인 입력 컨트롤이 빈 상태 | PASS | AR-05, AR-06 |
| A-07 | 모달 | 긴 원문이 viewport 초과 | 내부 세로 스크롤로 접근 가능 | PASS | AR-08, AR-09 |
| A-08 | 모달 | Escape 종료/포커스 손실 | 닫힌 뒤 호출 버튼으로 포커스 복귀 | PASS | AR-07 |
| A-09 | 브라우저 런타임 | console error/warn | 0건 | PASS | AR-10 |
| A-10 | 반응형 | 320px 폭 | 잘림 없이 모달 조작 가능(내부 가로·세로 스크롤 허용) | PASS | AR-09 |

## artifactRefs

| id | kind | description | path |
|---|---|---|---|
| AR-01 | text | F-008 DOM: 전체 문단, 섬유혼용률/디자인·규격/교복사진 3개 표 | `artifacts/qa/wp03-final-browser-qa/f008-dom.txt` |
| AR-02 | screenshot | F-008 미리보기 화면 캡처(1280×720) | `artifacts/qa/wp03-final-browser-qa/f008-screenshot-capture.md` |
| AR-03 | log | F-008 HWPX 작성 및 누락값 안내 결과 | `artifacts/qa/wp03-final-browser-qa/f008-hwpx-log.txt` |
| AR-04 | log | 지연·오류 fault-injection 실행 불가 사유 및 미실행 기록 | `artifacts/qa/wp03-final-browser-qa/pdf-delay-error-blocker.txt` |
| AR-05 | text | F-050 DOM 및 privacy 안내 | `artifacts/qa/wp03-final-browser-qa/f050-dom.txt` |
| AR-06 | screenshot | F-050 빈 양식 화면 캡처(320/기본 viewport) | `artifacts/qa/wp03-final-browser-qa/f050-screenshot-capture.md` |
| AR-07 | log | Escape 후 activeElement가 F-050 호출 버튼임을 확인 | `artifacts/qa/wp03-final-browser-qa/focus-log.txt` |
| AR-08 | log | 1280×720 실측 및 modal scroll metrics | `artifacts/qa/wp03-final-browser-qa/responsive-1280.txt` |
| AR-09 | log | 320×720 실측, overflow/scroll metrics 및 캡처 관찰 | `artifacts/qa/wp03-final-browser-qa/responsive-320.txt` |
| AR-10 | log | 브라우저 console error/warn 결과 | `artifacts/qa/wp03-final-browser-qa/console-errors.txt` |
| AR-11 | log | npm lint 출력 및 종료코드 | `artifacts/qa/wp03-final-browser-qa/npm-lint.txt` |
| AR-12 | log | npm build 출력 및 종료코드 | `artifacts/qa/wp03-final-browser-qa/npm-build.txt` |
| AR-13 | log | PDF busy 중 disabled→완료 후 enabled 결과 | `artifacts/qa/wp03-final-browser-qa/pdf-busy-log.txt` |

