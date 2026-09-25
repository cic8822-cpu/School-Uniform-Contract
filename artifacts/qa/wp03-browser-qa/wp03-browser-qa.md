# WP03 브라우저 수동 QA

- 기준 커밋: `8f8fffeb4300b96853e760da87a8a51b9fdb6404`
- 표면: `http://127.0.0.1:4173/`, Codex In-app Browser
- 사용 값: `테스트학교`, 학년도 `2026`, 문서번호 `문서-001`, 발행일 `2026-09-25`만 입력. 개인정보 가능 값은 입력하지 않음.
- 결론: **BLOCK**

## surfaceEvidence

| 시나리오 | 기준 | 표면·정확한 호출 | 판정 | 증빙 |
|---|---|---|---|---|
| S-01 | F-008 전체 원문 미리보기 | 서식함 → F-008 버튼 → 미리보기, 1280x720 | PASS | `f008-dom.txt` |
| S-02 | F-008 HWPX 작성 흐름 | F-008 미리보기 → `HWPX 작성` 클릭 | PASS | `f008-dom.txt`, `interaction-log.md` |
| S-03 | F-050 빈 양식·개인정보 비자동표시 | 서식함 → F-050 버튼 → 미리보기 | PASS | `f050-dom.txt` |
| S-04 | 1280px 모달 스크롤 | F-050 모달 DOM metrics: scrollHeight 1527, clientHeight 648, overflowY auto | PASS | `interaction-log.md` |
| S-05 | 1280px 콘솔 무오류 | `tab.dev.logs({levels:['error','warn']})` | PASS | `console-errors.txt` |
| S-06 | lint/build | `npm run lint`; `npm run build` | PASS | `npm-results.txt` |
| S-07 | Esc 및 포커스 복귀 | F-050 모달에서 `Escape` 입력 후 activeElement 확인 | FAIL | `interaction-log.md` — 모달은 닫혔으나 activeElement가 BODY이며 호출 버튼으로 복귀하지 않음 |
| S-08 | 320px 반응형 | 320px viewport에서 동일 시나리오 | BLOCK | `interaction-log.md` — 현재 브라우저 표면에 viewport 변경 capability가 없어 실행 불가. 추정 PASS로 처리하지 않음. |

## adversarialCases

| 시나리오 | 기준 | 적대 클래스 | 기대 동작 | 판정 | 증빙 |
|---|---|---|---|---|---|
| A-01 | F-008 | 원문이 입력 요약으로 축약되는 경우 | 전체 문단과 표가 보여야 함 | PASS | `f008-dom.txt` |
| A-02 | F-008 | 필수 일부 누락 상태에서 작성 | 버튼이 동작하고 상태가 사용자에게 보여야 함 | PASS | `f008-dom.txt`, `interaction-log.md` |
| A-03 | F-050 | 개인정보 응답값 자동 유입 | 응답 셀·기타 의견이 빈 상태여야 함 | PASS | `f050-dom.txt` |
| A-04 | 모달 | 긴 원문이 viewport를 넘는 경우 | 내부 스크롤로 접근 가능해야 함 | PASS | `interaction-log.md` |
| A-05 | 모달 | 키보드 Escape 종료 및 포커스 | 닫힌 뒤 호출 컨트롤로 포커스 복귀 | FAIL | `interaction-log.md` |
| A-06 | 브라우저 런타임 | console error/warn | 0건 | PASS | `console-errors.txt` |
| A-07 | 반응형 | 320px 레이아웃 | 잘림 없이 조작 가능 | BLOCK | `interaction-log.md` — viewport capability 부재 |

## artifactRefs

| id | kind | description | path |
|---|---|---|---|
| AR-01 | text | F-008 원문 문단·표 DOM 관찰 | `artifacts/qa/wp03-browser-qa/f008-dom.txt` |
| AR-02 | text | F-050 빈 양식·개인정보 보호 DOM 관찰 | `artifacts/qa/wp03-browser-qa/f050-dom.txt` |
| AR-03 | log | 브라우저 호출·viewport·스크롤·Esc·제약 기록 | `artifacts/qa/wp03-browser-qa/interaction-log.md` |
| AR-04 | log | 콘솔 error/warn 결과 | `artifacts/qa/wp03-browser-qa/console-errors.txt` |
| AR-05 | log | npm lint/build 종료코드 | `artifacts/qa/wp03-browser-qa/npm-results.txt` |

## 잔여 차단

1. Esc 종료 후 포커스가 호출 버튼으로 복귀하지 않음.
2. 현재 연결된 in-app browser가 viewport 변경을 제공하지 않아 320px 시나리오를 실행할 수 없음. 실제 320px 검증 전에는 반응형 PASS로 기록하지 않음.

