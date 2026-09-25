# 교복구매 웹앱 입력·인쇄·Windows 설치형·정적 배포 계획

## TL;DR
> Summary:      현재 React/Vite 클라이언트 앱에 B-02/B-03/C-05 입력 계약을 바로잡고, PDF 저장과 분리된 Form 직접 인쇄를 추가한 뒤, 동일 소스를 보안이 강화된 Electron NSIS 설치본과 Cloudflare Pages 정적 사이트로 제공한다. 두 채널 모두 서버·DB·업무자료 전송 없이 브라우저 메모리와 File/Blob API만 사용한다.
> Deliverables:
> - B-02 자유 서술 길이 UX, B-03 4개 승인 계약방식 선택, C-05 실제 달력 날짜 검증
> - `DocumentPreviewModal`의 `PDF 저장` + `바로 인쇄` 분리, 원문 전체만 인쇄하는 print CSS
> - Electron x64 + NSIS 설치 EXE, 오프라인 실행·설치·제거 검증과 SHA-256
> - Cloudflare Pages Direct Upload용 정적 빌드, preview 배포·네트워크 무전송 검증
> - 개인정보 보호 서식(F-029·F-047·F-050), 특히 F-050 빈 설문 양식 인쇄 회귀 게이트
> Effort:       XL
> Risk:         High - 외부 호스팅 금지였던 기존 결정을 이중 배포로 명시적으로 개정해야 하고, 네이티브 인쇄·설치·개인정보 무전송을 실제 Windows/브라우저에서 함께 입증해야 함

## Scope
### Must have
- B-02는 원천 계약대로 “학년 또는 대상 설명” 자유 문자열을 유지하되 2~100자 제약, `maxlength`, 글자수·오류 연결을 실제 UI와 XLSX 재입력 경로 모두에서 강제한다(`입력데이터_사전.md:36`, `webapp/public/data/fields.json:567-575`). 고정 학년 enum으로 축소하지 않는다.
- B-03은 자유 입력 폴백을 없애고 `contract-methods.json`의 4개 `excelDropdownLabel`을 단일 승인 목록으로 사용한다. 금액에 따른 자동판정은 계속 금지한다(`webapp/public/data/contract-methods.json:15-21`, `:48-52`, `:79-83`, 정책 `:10-13`).
- C-05는 `<input type="date">` UX와 별개로 XLSX import를 포함한 모든 값에 `YYYY-MM-DD` 및 실제 존재하는 날짜(윤년 포함)를 엄격 검증하고, 출력 시에만 `YYYY. MM. DD.`로 표시한다(`webapp/public/data/fields.json:339-348`, `webapp/src/lib/format.ts:1-7`).
- 누락값뿐 아니라 잘못된 입력값도 Form의 작성가능 상태와 HWPX/PDF/직접 인쇄를 차단한다. 반복 필드는 해당 Form이 요구할 때 최소 한 행의 유효값을 판정한다(`webapp/src/App.tsx:71-77`, `webapp/src/components/forms/FormsPage.tsx:38-44`).
- `DocumentPreviewModal`에 기존 PDF 다운로드와 별도의 `바로 인쇄` 버튼을 제공한다. 버튼 이름은 각각 `PDF 저장`, `바로 인쇄`로 구분한다(`webapp/src/components/modal/DocumentPreviewModal.tsx:176-187`, `:264-284`).
- 직접 인쇄는 브라우저와 Electron이 공유하는 `window.print()` 경로를 기본으로 하고, print media에서는 `.paper`의 원문 전체만 출력한다. 모달 제목·버튼·누락 안내·배경 앱은 인쇄하지 않고, 화면의 `max-height`/스크롤 때문에 원문이 잘리지 않게 한다(`webapp/src/components/modal/DocumentPreviewModal.tsx:190-285`). 별도 Electron IPC로 문서값을 전달하지 않는다.
- HWPX 원문 로딩 중이거나 로딩 오류가 있으면 PDF 저장과 바로 인쇄를 모두 차단하며 정확한 이유를 `role="status"`로 알린다(`webapp/src/components/modal/DocumentPreviewModal.tsx:119-150`). 출력 미지원/HWPX 전용 Form에는 가짜 인쇄 버튼을 노출하지 않는다(`웹앱_prd.md:46-48`).
- F-029·F-047·F-050은 허용 공통값 외 입력·저장·자동반영·로그·네트워크 전송을 금지한다. F-050은 응답·소계·총점·평균·자유서술을 모두 빈칸으로 유지한 인쇄 전용 설문만 미리보기·PDF·바로 인쇄한다(`입력데이터_사전.md:63-67`, `webapp/public/data/fields.json:1476-1595`).
- 정적 웹은 Cloudflare Pages의 정적 asset 배포만 사용한다. Pages Functions/Workers, DB binding, 인증, 분석·원격 로깅, API endpoint, localStorage/sessionStorage/IndexedDB를 도입하지 않는다. JSON/HWPX fetch는 동봉된 same-origin asset으로 제한한다(`webapp/src/data/useAppData.ts:11-15`, `:24-31`, `webapp/src/lib/download.ts:1-8`).
- Windows 설치형은 Electron x64 + electron-builder NSIS로 만들고, `nodeIntegration:false`, `contextIsolation:true`, `sandbox:true`, `webSecurity:true`, CSP, navigation/new-window 차단을 적용한다. renderer에 Node/Electron/범용 IPC를 노출하지 않는다.
- 로컬 설치본은 인터넷 없이 홈→입력→미리보기→HWPX/PDF/직접 인쇄→XLSX 백업이 동작해야 한다(`웹앱_로드맵.md:16-21`). Cloudflare 채널은 온라인 정적 편의 채널이며 입력은 탭 메모리에서만 유지하고 새로고침 시 사라진다.
- 정적 preview/production 배포와 외부 전송은 실행 시점에 사용자의 명시적 확인을 다시 받은 뒤 수행한다. 승인 전에는 로컬 build/검증까지만 실행한다.
- 원본 XLSM·HWPX의 SHA-256과 내용은 변경하지 않고, 마스킹 데이터만 테스트한다. 생성 설치본은 `artifacts/desktop/`, 정적 배포 후보는 `webapp/dist-web/`, 증빙은 `<attemptDir>`에 둔다.

### Must NOT have (guardrails, anti-slop, scope boundaries)
- 계약방법 금액 자동판정, 법령 기준 하드코딩, B-03 자유 입력, B-02의 근거 없는 학년 고정 목록을 추가하지 않는다.
- PDF 저장 버튼을 직접 인쇄로 이름만 바꾸거나, PDF를 자동 생성한 뒤 인쇄하는 우회 구현을 하지 않는다. 두 액션은 독립적이어야 한다.
- 미리보기 전체 모달·앱 chrome·스크롤된 일부 화면을 인쇄하지 않는다. 오직 준비 완료된 `.paper` 원문 전체만 인쇄한다.
- HWPX-only/미구현 Form, 원문 로딩 실패 Form, 누락·유효성 오류 Form에 인쇄 가능처럼 보이는 버튼을 제공하지 않는다.
- F-029/F-047/F-050의 개인식별·연락처·서명·개별 응답·자유서술을 fixture, screenshot, console, IPC, crash report, Pages 요청, 설치 로그, Git에 넣지 않는다.
- Electron에서 `nodeIntegration:true`, `contextIsolation:false`, `sandbox:false`, `webSecurity:false`, `shell.openExternal` 무제한 호출, 임의 URL navigation, raw `ipcRenderer` 노출을 허용하지 않는다.
- 로컬 HTTP 서버·작업 스케줄러를 설치형의 런타임 전제로 남기지 않는다. 기존 `npx serve` 방식은 개발/검증 보조 수단으로만 유지한다(`webapp/DEPLOY.md:5-20`).
- Cloudflare Pages Functions/Workers, KV/D1/R2, 서버 로그에 입력값을 남기는 API, Google Fonts 등 런타임 제3자 요청을 추가하지 않는다(`webapp/src/index.css:1`).
- 코드서명 인증서가 없는데 서명된 배포본이라고 주장하지 않는다. unsigned installer는 내부 QA용으로만 표시하고, 외부 배포 전 `Get-AuthenticodeSignature` 결과와 SmartScreen 위험을 명시한다.
- Electron 자동 업데이트, 사용자 계정/로그인, 클라우드 동기화, 다중 사용자 협업, 서비스워커 기반 입력 캐시는 이번 범위에 넣지 않는다.

## Verification strategy
> Zero human intervention - all verification is agent-executed.
- Test decision: TDD + Vitest/React Testing Library(jsdom), Playwright(real Chromium), PowerShell, Windows computer-use
- QA policy: every task has agent-executed scenarios. 기능 작업은 먼저 명시된 RED assertion 실패를 캡처하고 같은 assertion의 GREEN 통과를 캡처한다.
- Evidence: `<attemptDir>/task-<N>-<slug>.<ext>` — under ulw-loop, `<attemptDir>` is the `currentAttemptDir` from `omo ulw-loop status --json` (`.omo/evidence/ulw/<session>/<goalId>/a<attempt>`); outside ulw-loop use `.omo/evidence/`
- 공통 회귀 명령: `cd webapp; npm ci; npm run lint; npm run test:unit; npm run build:web; npm run build:desktop; npm run verify:preview-structure`.
- 개인정보 fixture는 `검증초등학교`, `2026`, `검증업체` 같은 마스킹 업무값만 사용하며 실제 사람·연락처·서명·설문 원자료는 사용하지 않는다.

## Execution strategy
### Parallel execution waves
> Target 5-8 tasks per wave. <3 per wave (except final) = under-splitting.
> Extract shared dependencies as Wave-1 tasks to maximize parallelism.

Wave 1 (no dependencies):
- Task 1: 테스트 하네스와 canonical Field 검증 계약
- Task 2: 웹/데스크톱 이중 Vite 빌드 경계
- Task 3: CSP·오프라인 폰트·무외부요청 기반
- Task 4: 보안 Electron 셸과 로컬 asset protocol
- Task 5: Cloudflare Pages 정적 배포 골격

Wave 2 (after Wave 1):
- Task 6: B-02/B-03/C-05 입력 UX — depends [1]
- Task 7: Form 작성가능·출력 차단 통합 — depends [1]
- Task 8: PDF 저장과 분리된 바로 인쇄 — depends [1, 3]
- Task 9: Electron NSIS 패키징 — depends [1, 2, 3, 4]
- Task 10: 정적 빌드 개인정보·배포 검증기 — depends [1, 2, 3, 5]

Wave 3 (after Wave 2):
- Task 11: 브라우저/로컬 정적 E2E — depends [6, 7, 8, 10]
- Task 12: Windows 설치·오프라인·인쇄·제거 E2E — depends [6, 7, 8, 9]
- Task 13: 교차 채널 개인정보·보안 독립 감사 — depends [6, 7, 8, 9, 10]

Wave 4 (final implementation/release closure):
- Task 14: 승인 기반 Cloudflare preview 배포, 회귀 확인, 프로젝트 문서/게이트 갱신 — depends [11, 12, 13]

Critical path: Task 1 -> Task 6 -> Task 11 -> Task 14

### Dependency matrix
| Task | Depends on | Blocks | Can parallelize with |
|------|------------|--------|----------------------|
| 1 | none | 6, 7, 8, 9, 10 | 2, 3, 4, 5 |
| 2 | none | 9, 10 | 1, 3, 4, 5 |
| 3 | none | 8, 9, 10 | 1, 2, 4, 5 |
| 4 | none | 9 | 1, 2, 3, 5 |
| 5 | none | 10 | 1, 2, 3, 4 |
| 6 | 1 | 11, 12, 13 | 7, 8, 9, 10 |
| 7 | 1 | 11, 12, 13 | 6, 8, 9, 10 |
| 8 | 1, 3 | 11, 12, 13 | 6, 7, 9, 10 |
| 9 | 1, 2, 3, 4 | 12, 13 | 6, 7, 8, 10 |
| 10 | 1, 2, 3, 5 | 11, 13 | 6, 7, 8, 9 |
| 11 | 6, 7, 8, 10 | 14 | 12, 13 |
| 12 | 6, 7, 8, 9 | 14 | 11, 13 |
| 13 | 6, 7, 8, 9, 10 | 14 | 11, 12 |
| 14 | 11, 12, 13 | F1-F4 | none |

## Todos
> Implementation + Test = ONE task. Never separate.
> Every task MUST have: References + Acceptance Criteria + QA Scenarios + Commit.

- [ ] 1. 테스트 하네스와 canonical Field 검증 계약 수립

  What to do: `package.json`/lockfile에 Vitest, React Testing Library, jsdom, Playwright, Electron, electron-builder, Wrangler를 dev dependency로 잠그고 `test:unit`, `test:e2e`, `build:web`, `build:desktop`, `package:win`, `verify:static` 스크립트를 정의한다. `FieldDef`에 명시적 enum source 계약을 추가하고 B-03은 `contractMethods[].excelDropdownLabel`을 source로 삼는다. `validation.ts`에 B-02 길이 및 실제 달력 C-05 검증, enum source resolve를 구현한다. 최초 RED test는 B-03 자유값 허용, C-05 `2026-02-29`, B-02 1/101자 허용을 각각 실패로 고정한다.
  Must NOT do: validation 설명문을 백틱으로 조작해 우연히 select가 되게 하지 말 것. 계약방식 문자열을 UI 여러 곳에 중복 하드코딩하지 말 것. 실제 개인정보 fixture를 만들지 말 것.

  Parallelization: Can parallel: YES | Wave 1 | Blocks: [6, 7, 8, 9, 10] | Blocked by: []

  References (executor has NO interview context - be exhaustive):
  - Pattern:  `webapp/package.json:6-30` - 현재 build/lint/preview만 있고 테스트·Electron·Wrangler가 없는 스크립트/의존성 기준
  - API/Type: `webapp/src/types.ts:139-164` - Field type과 `FieldDef` 확장 지점
  - Pattern:  `webapp/src/lib/validation.ts:7-23` - 현 enum/길이 파서와 자유입력 폴백 원인
  - Pattern:  `webapp/src/lib/validation.ts:27-58` - 현재 날짜 분기가 없고 enum option이 있을 때만 검사하는 검증기
  - API/Type: `webapp/public/data/fields.json:567-606` - B-02/B-03 현재 계약
  - API/Type: `webapp/public/data/contract-methods.json:15-21` - B-03 canonical 값의 첫 항목; 나머지도 `:48-52`, `:79-83`, `:138` 이후 동일 구조
  - External: `https://vitest.dev/guide/` - Vitest 설정/실행
  - External: `https://playwright.dev/docs/intro` - Playwright 설치/브라우저 실행

  Acceptance criteria (agent-executable only):
  - [ ] 변경 전 `npm run test:unit -- validation`에서 B-02/B-03/C-05 RED assertion이 실패한 증빙과 변경 후 동일 명령 0 exit GREEN 증빙이 있다.
  - [ ] `npm ci`가 lockfile만으로 성공하고 `npm run lint && npm run test:unit`이 0 exit다.
  - [ ] B-03 resolved options가 정확히 4개이고 각 value가 contract method의 `excelDropdownLabel`과 순서까지 일치한다.
  - [ ] `2028-02-29`는 통과하고 `2026-02-29`, `2026-13-01`, `2026.09.25.`는 입력값으로 거부된다. 출력 formatter만 dot date를 만든다.
  - [ ] B-02는 2자와 100자는 통과, 1자와 101자는 거부된다.

  QA scenarios (MANDATORY - task incomplete without these):
  ```
  Scenario: canonical Field 규칙 GREEN
    Tool:     bash
    Steps:    `cd webapp && npm ci && npm run test:unit -- validation field-options`
    Expected: B-02 경계 4건, B-03 승인목록 4개, C-05 실제 날짜/윤년 assertion이 모두 PASS하고 exit 0
    Evidence: <attemptDir>/task-1-field-contract.txt

  Scenario: 승인되지 않은 값 RED
    Tool:     bash
    Steps:    `cd webapp && npm run test:unit -- validation -t "invalid imported values"`
    Expected: `임의계약`, `2026-02-29`, 101자 B-02 각각 정확한 한글 오류 메시지를 반환하고 유효값으로 판정되지 않음
    Evidence: <attemptDir>/task-1-field-contract-error.txt
  ```

  Commit: YES | Message: `test(webapp): establish canonical field validation contracts` | Files: [`webapp/package.json`, `webapp/package-lock.json`, `webapp/vitest.config.ts`, `webapp/playwright.config.ts`, `webapp/src/types.ts`, `webapp/src/lib/validation.ts`, `webapp/src/lib/fieldOptions.ts`, `webapp/src/lib/*.test.ts`]

- [ ] 2. 웹/데스크톱 이중 Vite 빌드 경계 확정

  What to do: `vite.config.ts`를 mode 기반으로 분기해 Pages용 web build는 root base(`/`)와 `dist-web`, Electron renderer는 상대 base(`./`)와 `dist-desktop`을 사용한다. 같은 React 소스를 사용하고 target별 업무 로직 fork를 금지한다. Node/npm preflight에서 버전과 OS를 기록한다.
  Must NOT do: Electron을 위해 `webSecurity`를 끄거나 웹 빌드에 절대 로컬 경로를 넣지 말 것. 기존 `npm run build`를 모호하게 남기지 말고 web build alias로 고정할 것.

  Parallelization: Can parallel: YES | Wave 1 | Blocks: [9, 10] | Blocked by: []

  References:
  - Pattern:  `webapp/vite.config.ts:1-8` - 현재 단일 `base:'./'` 설정
  - Pattern:  `webapp/DEPLOY.md:5-7` - 현재 dist 정적성 및 일반 브라우저 `file://` 실패 실측
  - Pattern:  `webapp/src/lib/assetUrl.ts:1-10` - BASE_URL 기반 상대 asset 계산
  - API/Type: `webapp/src/data/useAppData.ts:24-31` - 두 target에서 반드시 읽혀야 할 정적 JSON 4종
  - Test:     `webapp/scripts/verify-hwpx-preview-structure.mjs` - 빌드 후 유지할 HWPX 구조검사 패턴
  - External: `https://vite.dev/config/shared-options.html#base` - Vite base 계약

  Acceptance criteria:
  - [ ] `npm run build:web`이 `webapp/dist-web/index.html`을 만들고 asset URL은 `/assets/...` 또는 root 기준이다.
  - [ ] `npm run build:desktop`이 `webapp/dist-desktop/index.html`을 만들고 asset URL은 `./assets/...` 상대경로다.
  - [ ] 두 결과 모두 data JSON 4종과 `public/templates/*.hwpx` 전체를 포함하고, 절대 `C:\Users\...` 문자열이 0건이다.
  - [ ] `npm run build`는 명시적으로 `build:web`과 동일하다.

  QA scenarios:
  ```
  Scenario: 두 target 빌드
    Tool:     bash
    Steps:    `cd webapp && node --version && npm --version && npm run build:web && npm run build:desktop`
    Expected: 두 명령 exit 0, dist-web/dist-desktop에 index·assets·data·templates 존재
    Evidence: <attemptDir>/task-2-dual-build.txt

  Scenario: 로컬 절대경로 누출
    Tool:     bash
    Steps:    `rg -n "C:\\\\Users|나연수_26\.9\.10" webapp/dist-web webapp/dist-desktop`
    Expected: 검색 결과 0건(exit 1은 정상 무매치로 기록)
    Evidence: <attemptDir>/task-2-dual-build-error.txt
  ```

  Commit: YES | Message: `build(webapp): separate web and desktop renderer targets` | Files: [`webapp/vite.config.ts`, `webapp/tsconfig*.json`]

- [ ] 3. CSP·오프라인 폰트·무외부요청 기반 적용

  What to do: Google Fonts runtime import를 제거하고 시스템 한글 font stack 또는 저장소에 포함된 로컬 font만 사용한다. `index.html`에 `default-src 'self'`, `script-src 'self'`, `style-src 'self'`, `font-src 'self'`, `connect-src 'self'`, 필요한 `img-src 'self' data: blob:`, `object-src 'none'`, `base-uri 'self'`, `form-action 'none'`, `frame-ancestors 'none'` CSP를 적용한다. 브라우저/`app:` Electron 양쪽에서 필요한 Blob 다운로드·canvas PDF가 막히지 않는 최소 허용만 유지한다.
  Must NOT do: `unsafe-eval`, wildcard source, Google Fonts/CDN, telemetry domain을 허용하지 말 것. CSP 오류를 피하려고 정책을 제거하지 말 것.

  Parallelization: Can parallel: YES | Wave 1 | Blocks: [8, 9, 10] | Blocked by: []

  References:
  - Pattern:  `webapp/src/index.css:1-2` - 현재 Google Fonts 외부 요청과 fallback stack
  - Pattern:  `webapp/src/lib/pdf.ts:12-37` - canvas/data URL/Blob이 필요한 PDF 경로
  - Pattern:  `webapp/src/lib/download.ts:1-8` - Blob URL 다운로드 경로
  - External: `https://www.electronjs.org/docs/latest/tutorial/security` - CSP, webSecurity, 원격 콘텐츠 제한
  - External: `https://developer.mozilla.org/docs/Web/HTTP/CSP` - CSP directive 의미

  Acceptance criteria:
  - [ ] `rg -n "fonts.googleapis.com|fonts.gstatic.com|@import url\(https?" webapp/src webapp/index.html` 결과가 0건이다.
  - [ ] unit/browser smoke에서 HWPX/PDF/XLSX Blob 다운로드가 CSP 위반 없이 시작된다.
  - [ ] Playwright `request` 수집에서 same-origin 정적 GET 외 요청이 0건이다.

  QA scenarios:
  ```
  Scenario: CSP 하 정상 렌더·다운로드
    Tool:     playwright(real Chrome)
    Steps:    `cd webapp && npm run build:web && npm run preview:web`; Chrome으로 root→기초자료→서식 F-008 미리보기→PDF 저장을 실행하고 console/securitypolicyviolation을 수집
    Expected: UI 정상, 다운로드 발생, CSP violation·외부 도메인 요청 0건
    Evidence: <attemptDir>/task-3-csp-network.json

  Scenario: 외부 요청 차단
    Tool:     playwright(real Chrome)
    Steps:    페이지에서 `fetch('https://example.com/telemetry',{method:'POST',body:'masked'})`를 evaluate하고 결과/console을 캡처
    Expected: CSP `connect-src 'self'`로 요청이 차단되고 실제 전송 성공이 아님
    Evidence: <attemptDir>/task-3-csp-network-error.json
  ```

  Commit: YES | Message: `security(webapp): enforce offline-safe content policy` | Files: [`webapp/index.html`, `webapp/src/index.css`, `webapp/public/fonts/**`]

- [ ] 4. 보안 Electron 셸과 로컬 asset protocol 구현

  What to do: `webapp/electron/main.mjs`에 production `app://bundle/` 표준·보안 custom protocol을 등록해 `dist-desktop`의 읽기 전용 resource만 안전하게 매핑한다. path traversal을 거부하고 dev에서는 loopback Vite URL만 허용한다. BrowserWindow는 `nodeIntegration:false`, `contextIsolation:true`, `sandbox:true`, `webSecurity:true`; navigation과 `window.open`은 allowlist 없이 거부한다. renderer API/IPC/preload는 만들지 않는다. 직접 인쇄는 renderer `window.print()`가 네이티브 dialog를 열도록 둔다.
  Must NOT do: 임의 파일시스템 경로, shell, raw IPC, 업무자료 payload를 main process로 넘기지 말 것. `file://` CORS 우회를 위해 보안 플래그를 끄지 말 것.

  Parallelization: Can parallel: YES | Wave 1 | Blocks: [9] | Blocked by: []

  References:
  - Pattern:  `webapp/DEPLOY.md:5-16` - 일반 `file://` 대신 안전한 로컬 실행 경로가 필요한 근거
  - Pattern:  `webapp/src/lib/assetUrl.ts:6-10` - renderer asset URL 계약
  - External: `https://www.electronjs.org/docs/latest/tutorial/security` - Electron 보안 체크리스트
  - External: `https://www.electronjs.org/docs/latest/api/browser-window` - BrowserWindow 보안 옵션
  - External: `https://www.electronjs.org/docs/latest/api/protocol` - custom protocol 등록/처리
  - External: `https://www.electronjs.org/docs/latest/tutorial/sandbox` - sandbox 계약

  Acceptance criteria:
  - [ ] unit test가 `..`, `%2e%2e`, 절대경로 요청을 403/404로 거부하고 `data/*.json`, `templates/*.hwpx`만 resource root 안에서 해석한다.
  - [ ] Electron smoke에서 `process`, `require`, `fs`, `ipcRenderer`가 renderer global에 없다.
  - [ ] 외부 navigation/new-window 시도가 열리지 않고 app 화면은 유지된다.
  - [ ] main/console 로그에 field value, workbook row, HWPX token value를 쓰는 코드가 없다.

  QA scenarios:
  ```
  Scenario: 설치 전 Electron shell smoke
    Tool:     bash
    Steps:    `cd webapp && npm run build:desktop && npm run electron:smoke`
    Expected: app://bundle 홈·JSON·F-008 템플릿 HTTP-equivalent 응답 성공, renderer Node globals 부재
    Evidence: <attemptDir>/task-4-electron-shell.txt

  Scenario: path traversal·외부 창 공격
    Tool:     bash
    Steps:    `cd webapp && npm run test:unit -- electron-security`
    Expected: traversal 3종과 `window.open('https://example.com')`이 모두 거부되고 테스트 exit 0
    Evidence: <attemptDir>/task-4-electron-shell-error.txt
  ```

  Commit: YES | Message: `feat(desktop): add sandboxed electron application shell` | Files: [`webapp/electron/main.mjs`, `webapp/electron/protocol.mjs`, `webapp/electron/*.test.mjs`]

- [ ] 5. Cloudflare Pages 정적 배포 골격 추가

  What to do: Direct Upload을 선택하고 `wrangler.jsonc`에 `pages_build_output_dir`만 선언한다. 현재 앱은 내부 state 라우팅이므로 root가 기준이나, 향후 새로고침 안전성을 위해 `public/_redirects`의 `/* /index.html 200` 규칙을 명시하고 실제 asset 경로가 SPA fallback에 가려지지 않는 검증을 둔다. Functions 디렉터리는 만들지 않는다. project name은 `uniform-purchase-guide`로 고정한다.
  Must NOT do: Git integration과 Direct Upload을 동시에 설정하지 말 것. Cloudflare binding/secret/API를 app bundle에 넣지 말 것. 승인 전 `wrangler pages deploy`를 실행하지 말 것.

  Parallelization: Can parallel: YES | Wave 1 | Blocks: [10] | Blocked by: []

  References:
  - Pattern:  `webapp/DEPLOY.md:5-16` - 현재 정적 dist/HTTP 서빙 계약
  - Pattern:  `webapp/src/App.tsx:42-49` - URL router가 아닌 내부 route state
  - External: `https://developers.cloudflare.com/pages/framework-guides/deploy-a-vite3-project/` - Vite Pages build
  - External: `https://developers.cloudflare.com/pages/get-started/direct-upload/` - Direct Upload 명령과 제한
  - External: `https://developers.cloudflare.com/pages/configuration/redirects/` - `_redirects` 규칙

  Acceptance criteria:
  - [ ] `wrangler pages deploy` 문자열은 package script/문서에만 있고 자동 postbuild/postinstall hook에는 없다.
  - [ ] `webapp/functions`, `_worker.js`, Pages binding 설정이 존재하지 않는다.
  - [ ] `dist-web/_redirects`가 존재하고 root/index/real asset 요청과 unknown path 동작이 자동 test로 구분된다.

  QA scenarios:
  ```
  Scenario: 로컬 Pages 호환 smoke
    Tool:     bash
    Steps:    `cd webapp && npm run build:web && npx wrangler pages dev dist-web --ip 127.0.0.1 --port 4174`
    Expected: root 200, data JSON·HWPX 200/정상 MIME, unknown route는 index fallback, 존재하지 않는 asset은 검증기가 실패로 탐지
    Evidence: <attemptDir>/task-5-pages-scaffold.txt

  Scenario: 서버 기능 혼입 검사
    Tool:     bash
    Steps:    `cd webapp && npm run verify:static -- --assert-no-functions`
    Expected: Functions/Workers/binding/API/secret 후보 0건, exit 0
    Evidence: <attemptDir>/task-5-pages-scaffold-error.txt
  ```

  Commit: YES | Message: `build(webapp): add static pages deployment scaffold` | Files: [`webapp/wrangler.jsonc`, `webapp/public/_redirects`]

- [ ] 6. B-02/B-03/C-05 입력 UX 구현

  What to do: B-02는 자유 text+`minLength=2`+`maxLength=100`+현재 글자수/오류 안내, B-03은 contract method source 기반 `<select>` 4개, C-05는 date input+엄격 오류 안내로 렌더한다. 모든 input에 stable id, label `htmlFor`, hint/error `aria-describedby`, 오류 시 `aria-invalid`를 연결한다. XLSX import 후 잘못된 값도 동일 오류를 즉시 표시한다.
  Must NOT do: B-02를 학년 checkbox/enum으로 바꾸지 말 것. B-03에 임의 옵션을 추가하거나 계약방법 자동판정을 넣지 말 것. native date UI만 믿고 imported string 검증을 생략하지 말 것.

  Parallelization: Can parallel: YES | Wave 2 | Blocks: [11, 12, 13] | Blocked by: [1]

  References:
  - Pattern:  `webapp/src/components/input/FieldInput.tsx:15-18` - 현재 input type 결정
  - Pattern:  `webapp/src/components/input/FieldInput.tsx:23-49` - enum option이 없으면 text로 폴백하는 현재 B-03 원인
  - Pattern:  `webapp/src/components/input/CommonFieldsTab.tsx:18-27` - 실시간 오류 전달
  - Pattern:  `webapp/src/components/input/InputPage.tsx:62-120` - 7탭 및 사업/공통 필드 렌더 연결
  - Pattern:  `webapp/src/lib/workbook.ts:55-83` - XLSX import가 임의 문자열을 복원하는 경로
  - Test:     `webapp/src/lib/validation.test.ts` - Task 1의 RED/GREEN 계약

  Acceptance criteria:
  - [ ] component test에서 B-02 `maxlength=100`, 글자수, 1/101자 오류와 2/100자 성공이 확인된다.
  - [ ] B-03은 정확히 placeholder+4개 option이고 자유 text input이 아니다.
  - [ ] C-05 `type=date`이며 XLSX로 주입된 비윤년 2월 29일은 오류로 표시된다.
  - [ ] 세 input 모두 accessible name, description, invalid state를 갖는다.

  QA scenarios:
  ```
  Scenario: 사업/공통 입력 happy path
    Tool:     browser:control-in-app-browser
    Steps:    `http://127.0.0.1:4173`에서 기초자료→사업 탭, B-02=`1학년 신입생`, B-03=`2단계 입찰(규격·가격 동시)` 선택; 공통 탭 C-05=`2028-02-29` 입력
    Expected: 오류 0건, B-03 자유입력 불가, 값이 미리보기/백업 state에 보존
    Evidence: <attemptDir>/task-6-input-ux.png

  Scenario: 경계·잘못된 import
    Tool:     playwright(real Chrome)
    Steps:    fixture XLSX에 B-02 101자, B-03 `임의계약`, C-05 `2026-02-29`를 넣어 업로드
    Expected: 세 필드 인접 오류, `aria-invalid=true`, 작성가능/출력으로 승격되지 않음
    Evidence: <attemptDir>/task-6-input-ux-error.png
  ```

  Commit: YES | Message: `feat(input): enforce guided contract and date entry` | Files: [`webapp/src/components/input/FieldInput.tsx`, `webapp/src/components/input/CommonFieldsTab.tsx`, `webapp/src/components/input/InputPage.tsx`, `webapp/src/components/input/input.css`, `webapp/src/components/input/*.test.tsx`]

- [ ] 7. Form 작성가능 판정과 출력 차단 통합

  What to do: `formReadiness.ts` 하나에서 필수값 누락, 단일값 오류, 필요한 반복행 유효성을 판정하고 App/Procedure/Forms/Export/Modal 모두 같은 결과를 사용한다. 상태에는 구체적인 field ID+한글 메시지를 유지하되 개인정보 값 자체는 포함하지 않는다. 잘못된 값은 `▲ 기초자료 필요`로 표시하고 모든 생성/인쇄 액션을 비활성화한다.
  Must NOT do: 화면별로 readiness를 중복 구현하거나 오류 로그에 원래 값을 넣지 말 것. 개인정보 보호 Form에 금지 field를 새로 요구하지 말 것.

  Parallelization: Can parallel: YES | Wave 2 | Blocks: [11, 12, 13] | Blocked by: [1]

  References:
  - Pattern:  `webapp/src/App.tsx:71-77` - 현재 공란만 세는 판정
  - Pattern:  `webapp/src/App.tsx:116-162` - Procedure/Forms/Modal로 판정을 전달하는 연결점
  - Pattern:  `webapp/src/components/forms/FormsPage.tsx:38-44` - ●/▲/○ 상태 표시
  - Pattern:  `webapp/src/components/exportPage/ExportPage.tsx:14-27` - 별도 readiness 계산 중복 지점
  - API/Type: `webapp/src/types.ts:120-131` - Form required/computed/privacy 계약
  - Test:     `웹앱_test.md:60-63` - 기존 WF-01/WF-02 완료조건

  Acceptance criteria:
  - [ ] 같은 Form/values/repeats에 대해 모든 화면의 readiness 결과가 동일하다.
  - [ ] B-02/B-03/C-05 오류가 있는 Form은 `▲`이고 HWPX/PDF/인쇄가 disabled다.
  - [ ] error object와 console output에 사용자가 입력한 원문 값이 없다.
  - [ ] F-029/F-047/F-050의 허용 목록은 기존 master와 동일하고 금지 field를 readiness가 요구하지 않는다.

  QA scenarios:
  ```
  Scenario: 유효 입력으로 작성가능 전환
    Tool:     playwright(real Chrome)
    Steps:    F-011 필수값을 마스킹 정상값으로 채우고 서식함·절차·내보내기·모달 상태를 순서대로 조회
    Expected: 네 화면 모두 `● 작성가능`, blocking issue 0
    Evidence: <attemptDir>/task-7-readiness.json

  Scenario: invalid 값 출력 차단
    Tool:     playwright(real Chrome)
    Steps:    C-05를 fixture import로 `2026-02-29`로 주입하고 F-011을 연다
    Expected: `▲ 기초자료 필요`, 오류명 표시, HWPX/PDF/바로 인쇄 disabled, 입력 원문 console 미출력
    Evidence: <attemptDir>/task-7-readiness-error.json
  ```

  Commit: YES | Message: `fix(webapp): centralize form readiness and output blocking` | Files: [`webapp/src/lib/formReadiness.ts`, `webapp/src/App.tsx`, `webapp/src/components/procedure/ProcedurePage.tsx`, `webapp/src/components/forms/FormsPage.tsx`, `webapp/src/components/exportPage/ExportPage.tsx`, `webapp/src/lib/formReadiness.test.ts`]

- [ ] 8. PDF 저장과 분리된 Form 바로 인쇄 구현

  What to do: Modal의 기존 `PDF 출력`을 `PDF 저장`으로 명확히 바꾸고 독립 `바로 인쇄` 버튼을 추가한다. `printDocumentElement`는 인쇄 대상 marker/body class를 적용한 뒤 `window.print()`를 호출하고 `afterprint`에서 반드시 정리한다. `@media print`는 `.paper` 전체만 보이게 하고 modal/app chrome, 상태/액션, overflow/max-height를 제거하며 A4 margin/page-break/table header를 정의한다. HWPX source가 pending/error거나 readiness issue가 있으면 양쪽 출력 모두 차단한다. `outputMethod`가 `both` 또는 `pdf`인 Form만 인쇄 가능하다.
  Must NOT do: PDF 저장을 호출한 뒤 인쇄하지 말 것. 캡처 이미지 한 장을 프린트하지 말 것. Electron 전용 fork/IPC에 문서 내용을 보내지 말 것. HWPX-only/none Form에 인쇄 버튼을 보이지 말 것.

  Parallelization: Can parallel: YES | Wave 2 | Blocks: [11, 12, 13] | Blocked by: [1, 3]

  References:
  - Pattern:  `webapp/src/components/modal/DocumentPreviewModal.tsx:119-150` - HWPX pending/error 상태
  - Pattern:  `webapp/src/components/modal/DocumentPreviewModal.tsx:152-188` - 출력 capability와 PDF handler
  - Pattern:  `webapp/src/components/modal/DocumentPreviewModal.tsx:213-284` - `.paper` 및 action bar
  - Pattern:  `webapp/src/lib/pdf.ts:4-37` - PDF 저장은 별도 유지할 기존 구현
  - Pattern:  `webapp/src/components/modal/modal.css` - print media 추가 위치
  - API/Type: `webapp/public/data/fields.json:1580-1595` - F-050 빈 설문 output policy
  - External: `https://developer.mozilla.org/docs/Web/API/Window/print` - `window.print()` 계약
  - External: `https://developer.mozilla.org/docs/Web/CSS/CSS_media_queries/Printing` - print media CSS

  Acceptance criteria:
  - [ ] 모달에 `PDF 저장`과 `바로 인쇄`가 서로 다른 버튼/handler로 존재하고 component test에서 `window.print` 1회 호출을 확인한다.
  - [ ] print media screenshot/PDF에 `.paper` 원문 전체만 있고 modal heading/action/status/background app은 없다.
  - [ ] HWPX pending/error 및 invalid readiness에서 `바로 인쇄`가 disabled되고 `window.print` 호출은 0회다.
  - [ ] F-050 인쇄 결과에서 빈 리커트/합계/평균/기타 의견 칸이 유지되고 개인 응답 데이터가 0건이다.

  QA scenarios:
  ```
  Scenario: F-008 원문 전체 직접 인쇄
    Tool:     playwright(real Chrome)
    Steps:    F-008 미리보기 source 로딩 완료 후 `바로 인쇄`를 click하되 `window.print`를 spy하고 `page.emulateMedia({media:'print'})`로 full-page PDF를 저장
    Expected: print 1회, `.paper` 첫/마지막 원문 블록 모두 존재, modal chrome 0, PDF 저장 handler 미호출
    Evidence: <attemptDir>/task-8-direct-print.pdf

  Scenario: 로딩 실패/F-050 개인정보 경계
    Tool:     playwright(real Chrome)
    Steps:    F-008 template fetch를 500으로 intercept해 인쇄 차단 확인; 이어 정상 F-050을 열어 마스킹 공통값만 둔 채 print media PDF 생성
    Expected: 실패 시 버튼 disabled+정확한 오류/print 0회; F-050은 응답·소계·평균·자유서술 값이 모두 빈칸
    Evidence: <attemptDir>/task-8-direct-print-error.pdf
  ```

  Commit: YES | Message: `feat(print): add document-only direct printing` | Files: [`webapp/src/lib/print.ts`, `webapp/src/components/modal/DocumentPreviewModal.tsx`, `webapp/src/components/modal/modal.css`, `webapp/src/components/modal/DocumentPreviewModal.test.tsx`]

- [ ] 9. Electron x64 NSIS 패키징 구성

  What to do: `electron-builder.yml`에 고유 `appId`, productName `교복구매 길라잡이`, x64 NSIS, per-user 설치, 변경 가능한 설치 위치, 시작 메뉴/바탕화면 바로가기, `deleteAppDataOnUninstall:false`, `asar:true`, deterministic artifact name을 설정한다. `dist-desktop`은 `extraResources`의 읽기 전용 `app` 디렉터리로 포함한다. neutral favicon에서 256px 이상 ICO를 재현 가능하게 생성한다. output은 루트 `artifacts/desktop/`으로 고정한다.
  Must NOT do: nsis-web(설치 중 payload 다운로드), 자동 업데이트, 관리자 전용 per-machine 설치, uninstall 시 사용자 출력물 삭제를 넣지 말 것. 인증서 없이는 signed라고 표시하지 말 것.

  Parallelization: Can parallel: YES | Wave 2 | Blocks: [12, 13] | Blocked by: [1, 2, 3, 4]

  References:
  - Pattern:  `webapp/package.json:1-32` - package metadata/scripts 확장 위치
  - Pattern:  `webapp/public/favicon.svg` - 설치 아이콘의 기존 source
  - Pattern:  `webapp/public/data/*.json` - 패키지에 포함할 master data
  - Pattern:  `webapp/public/templates/*.hwpx` - 패키지에 포함할 template set
  - External: `https://www.electron.build/nsis/` - NSIS 옵션
  - External: `https://www.electron.build/docs/configuration/` - files/extraResources/asar/output 설정

  Acceptance criteria:
  - [ ] `npm run package:win`이 `artifacts/desktop/교복구매_길라잡이_Setup_<version>_x64.exe`와 unpacked 앱을 만든다.
  - [ ] installer는 x64 NSIS이며 네트워크 다운로드 없이 설치 가능하다.
  - [ ] packaged resources에 JSON 4종과 기대 HWPX template 수가 있고 원본 파일 SHA-256은 변경되지 않는다.
  - [ ] `Get-AuthenticodeSignature`가 Valid가 아니면 artifact manifest에 `INTERNAL_QA_UNSIGNED`가 명시된다.

  QA scenarios:
  ```
  Scenario: 패키지 생성·resource 검사
    Tool:     bash
    Steps:    `cd webapp && npm run package:win`; 이어 PowerShell `Get-FileHash ..\artifacts\desktop\*Setup*x64.exe -Algorithm SHA256`
    Expected: EXE 1개, SHA-256 64 hex, unpacked app smoke exit 0, template/data 누락 0
    Evidence: <attemptDir>/task-9-nsis-build.txt

  Scenario: 서명/원격 payload 안전 실패
    Tool:     bash
    Steps:    PowerShell `Get-AuthenticodeSignature`와 installer resource/config 검사를 실행
    Expected: Valid이면 subject 기록; NotSigned이면 내부 QA 라벨 존재. nsis-web/update URL/remote payload 0건
    Evidence: <attemptDir>/task-9-nsis-build-error.txt
  ```

  Commit: YES | Message: `build(desktop): package x64 nsis installer` | Files: [`webapp/electron-builder.yml`, `webapp/build/icon.ico`, `webapp/scripts/build-icon.mjs`, `webapp/package.json`, `webapp/package-lock.json`]

- [ ] 10. 정적 빌드 개인정보·배포 검증기 구현

  What to do: `verify-static-build.mjs`로 dist manifest, JSON/template 완전성, sourcemap 금지, 절대경로/외부 URL/API/Functions/secret 후보, CSP/redirect 존재를 검사한다. runtime Playwright test는 요청 method/URL/body를 기록하되 body 값은 저장하지 않고 same-origin GET만 허용한다. XLSX/HWPX/PDF/File API 경로가 POST 없이 완료되는지 확인한다.
  Must NOT do: 검증 로그에 form 값, workbook cell 값, 생성 HWPX XML 본문을 덤프하지 말 것. 검증 편의를 위해 Pages API를 추가하지 말 것.

  Parallelization: Can parallel: YES | Wave 2 | Blocks: [11, 13] | Blocked by: [1, 2, 3, 5]

  References:
  - Pattern:  `webapp/src/data/useAppData.ts:60-63` - 허용되는 same-origin fetch
  - Pattern:  `webapp/src/lib/hwpx.ts:15-43` - template fetch/Blob download
  - Pattern:  `webapp/src/lib/workbook.ts:18-58` - client-only XLSX 경로
  - Test:     `웹앱_test.md:29-38` - WQ-01~04 개인정보·서버 미전송 게이트
  - Pattern:  `webapp/public/data/forms.json:4724-4775` - F-047 보호 계약
  - Pattern:  `webapp/public/data/forms.json:5012-5060` - F-050 보호 계약

  Acceptance criteria:
  - [ ] `npm run verify:static`이 clean build에서 0 exit, intentional external URL/`.map`/`functions` fixture에서 nonzero다.
  - [ ] 기초자료 입력→XLSX import/export→HWPX→PDF→print 전 과정의 network audit가 same-origin GET만 보고한다.
  - [ ] dist/로그/증빙에 테스트 금지 marker(`010-`, 주민번호형, 실제 성명 fixture)가 0건이다.

  QA scenarios:
  ```
  Scenario: clean static artifact 검사
    Tool:     bash
    Steps:    `cd webapp && npm run build:web && npm run verify:static -- dist-web`
    Expected: CSP/redirect/data/templates PASS, source map·외부 endpoint·secret·절대경로·server code 0건
    Evidence: <attemptDir>/task-10-static-guard.json

  Scenario: 검증기 자기시험
    Tool:     bash
    Steps:    `cd webapp && npm run test:unit -- verify-static-build`
    Expected: fake source map, external POST URL, Functions 파일, Windows 절대경로 fixture를 각각 탐지하고 test exit 0
    Evidence: <attemptDir>/task-10-static-guard-error.txt
  ```

  Commit: YES | Message: `test(webapp): guard static artifacts and private data boundaries` | Files: [`webapp/scripts/verify-static-build.mjs`, `webapp/scripts/verify-static-build.test.mjs`, `webapp/tests/network-boundary.spec.ts`]

- [ ] 11. 브라우저·로컬 정적 E2E 및 출력 회귀 완료

  What to do: `dist-web`을 localhost로 서빙하고 신규 담당자 흐름, B-02/B-03/C-05, XLSX 왕복, F-008/F-011, PDF 저장, 바로 인쇄, HWPX, F-050 빈 설문을 real Chromium에서 실행한다. viewport 320/768/1440/1920, keyboard-only, print media PDF를 증빙한다. 네트워크는 same-origin 정적 GET만 허용한다.
  Must NOT do: mocked component test만으로 E2E를 대체하지 말 것. 실제 개인정보를 넣지 말 것. PDF 다운로드 성공만 보고 열기/페이지 검사를 생략하지 말 것.

  Parallelization: Can parallel: YES | Wave 3 | Blocks: [14] | Blocked by: [6, 7, 8, 10]

  References:
  - Test:     `웹앱_test.md:16-27` - WE-01~08 골든 흐름
  - Test:     `웹앱_test.md:49-63` - WA/WF 접근성·반응형·상태 검사
  - Pattern:  `webapp/src/App.tsx:102-169` - 전체 7개 화면과 modal 연결
  - Pattern:  `_workspace/05_web/학교행정업무길라잡이_웹판_UX_벤치마크.md:34-38` - 점진 입력, 근거, 안전 출력 원칙

  Acceptance criteria:
  - [ ] WE-01~08, 신규 입력 3건, PDF 저장/바로 인쇄 분리 시나리오가 real browser에서 PASS한다.
  - [ ] F-008/F-011 print PDF의 첫/마지막 블록이 있고 잘림/겹침/모달 chrome이 없다.
  - [ ] F-050 출력 3종(미리보기, PDF, print PDF)에 개인 응답값 0건이고 빈 양식이다.
  - [ ] 4개 viewport와 keyboard-only 흐름에 critical accessibility failure가 없다.

  QA scenarios:
  ```
  Scenario: 신규 담당자 전체 흐름
    Tool:     playwright(real Chrome)
    Steps:    `cd webapp && npm run test:e2e -- browser-golden.spec.ts --project=chromium`
    Expected: 계약방법 수동선택→절차→입력→서식→미리보기→PDF/HWPX/print→XLSX 왕복 전부 PASS, 외부 요청 0
    Evidence: <attemptDir>/task-11-browser-e2e.zip

  Scenario: invalid/보호 서식 회귀
    Tool:     playwright(real Chrome)
    Steps:    `cd webapp && npm run test:e2e -- browser-negative.spec.ts --project=chromium`
    Expected: invalid B-02/B-03/C-05 출력 차단, HWPX 500 차단, F-050 빈 양식 유지, 개인정보/console leak 0
    Evidence: <attemptDir>/task-11-browser-e2e-error.zip
  ```

  Commit: YES | Message: `test(webapp): verify input and print flows in real chromium` | Files: [`webapp/tests/browser-golden.spec.ts`, `webapp/tests/browser-negative.spec.ts`, `webapp/playwright.config.ts`]

- [ ] 12. Windows 설치·오프라인·네이티브 인쇄·제거 E2E 완료

  What to do: clean Windows VM/테스트 사용자에서 NSIS를 GUI 설치하고 네트워크를 끊은 상태로 앱을 실행한다. 한글+공백 설치경로, JSON/HWPX load, 입력, XLSX/HWPX/PDF 저장, `바로 인쇄`→Microsoft Print to PDF를 확인한다. output은 설치 폴더 밖 QA temp에 두고 uninstall 후 보존을 검사한다. 프로세스/console에 Node API·외부 요청·민감값이 없어야 한다.
  Must NOT do: 개발 서버가 켜진 상태로 offline PASS를 주장하지 말 것. 기존 사용자 문서나 원본을 제거하지 말 것. unsigned QA EXE를 외부 배포하지 말 것.

  Parallelization: Can parallel: YES | Wave 3 | Blocks: [14] | Blocked by: [6, 7, 8, 9]

  References:
  - Pattern:  `웹앱_로드맵.md:16-21` - M4 설치본 offline 완료조건
  - Pattern:  `웹앱_task.md:38-43` - 현재 offline/설치 잔여 상태
  - Pattern:  `webapp/DEPLOY.md:18-20` - 기존 현장 방식 미확정 상태
  - External: `https://www.electron.build/nsis/` - NSIS 설치/제거 옵션
  - External: `https://www.electronjs.org/docs/latest/api/web-contents#contentsprintoptions-callback` - Electron의 native print 동작 참고; 구현은 renderer `window.print()` 유지

  Acceptance criteria:
  - [ ] 한글+공백 경로 설치/실행/제거가 agent computer-use로 성공하고 screenshot/terminal evidence가 있다.
  - [ ] 네트워크 어댑터 차단 또는 요청 차단 상태에서 핵심 flow와 모든 bundled asset이 정상이다.
  - [ ] `바로 인쇄`가 native dialog를 열고 Microsoft Print to PDF 결과에 `.paper` 전체만 존재한다.
  - [ ] F-050 native print PDF가 빈 설문이며 individual response/free text가 없다.
  - [ ] uninstall 후 앱 파일은 제거되고 외부 QA output 파일은 그대로다.

  QA scenarios:
  ```
  Scenario: clean install·offline·print·uninstall
    Tool:     computer-use
    Steps:    PowerShell `Start-Process -FilePath <Setup_x64.exe>`로 visible installer 실행→`C:\Users\Public\Documents\교복 길라잡이 QA\` 선택→앱 실행→네트워크 차단 상태 F-008 입력/미리보기/바로 인쇄→Microsoft Print to PDF로 `<attemptDir>\task-12-electron-print.pdf` 저장→uninstaller GUI 실행
    Expected: 설치·실행·인쇄·제거 성공, 인쇄 PDF 원문 전체, output 보존, 외부 요청/개발 서버 의존 0
    Evidence: <attemptDir>/task-12-electron-install.mp4

  Scenario: F-050·보안 음성 경로
    Tool:     computer-use
    Steps:    설치 앱에서 F-050 열기→빈 양식 바로 인쇄; devtools/자동 probe로 `require/process` 부재와 외부 navigation 차단 확인
    Expected: F-050 빈칸 유지, Node API 미노출, 외부 창 0, app 안정 유지
    Evidence: <attemptDir>/task-12-electron-error.pdf
  ```

  Commit: YES | Message: `test(desktop): verify nsis install offline print and uninstall` | Files: [`webapp/tests/electron-smoke.mjs`, `webapp/scripts/verify-installed-app.ps1`]

- [ ] 13. 교차 채널 개인정보·보안 독립 감사

  What to do: 웹과 Electron 산출물을 독립 검토해 F-029/F-047/F-050 허용 field, HWPX token whitelist, direct print, XLSX, logs, IPC, network, dist/installer content를 대조한다. 원본 5개 파일 SHA-256도 `.claude/quality-gate.json` 값과 재대조한다. Critical/High 발견 시 해당 task를 GREEN으로 되돌리지 말고 수정 task로 재개한다.
  Must NOT do: 실제 개인정보를 주입해 차단 시험하지 말 것. 원본 HWPX/XLSM을 열어 저장하거나 installer에 포함하지 말 것.

  Parallelization: Can parallel: YES | Wave 3 | Blocks: [14] | Blocked by: [6, 7, 8, 9, 10]

  References:
  - Pattern:  `.claude/quality-gate.json:1-32` - 활성 phase 및 원본 SHA-256 기준
  - Pattern:  `.claude/hooks/verify-phase.ps1:1-90` - gate/원본 hash 검증 방식
  - Pattern:  `_workspace/05_web/웹앱_WP02_검증결과.md:105-118` - 보호 서식 허용/금지 field 확정
  - Pattern:  `webapp/src/lib/hwpx.ts:1-43` - 허용/금지 token 및 download 경로
  - Test:     `웹앱_test.md:29-38` - WQ 개인정보·서버 미전송 기준

  Acceptance criteria:
  - [ ] F-029/F-047/F-050 허용/금지 matrix가 master, UI, preview, HWPX, PDF, print, XLSX에서 동일하다.
  - [ ] renderer→main IPC 채널 0개 또는 업무값을 받는 채널 0개, external runtime request 0개다.
  - [ ] dist/installer/evidence/log에서 실제 개인정보·비밀·절대경로 0건이다.
  - [ ] 원본 5개 SHA-256이 gate 값과 모두 일치한다.

  QA scenarios:
  ```
  Scenario: 두 산출물 보호 matrix 감사
    Tool:     bash
    Steps:    `powershell -ExecutionPolicy Bypass -File .claude/hooks/verify-phase.ps1 -Mode static`; `cd webapp && npm run audit:privacy -- ../artifacts/desktop dist-web`
    Expected: 보호 field/token/네트워크/로그 위반 0, 원본 hash 5/5 일치
    Evidence: <attemptDir>/task-13-privacy-audit.json

  Scenario: 금지 marker 자기시험
    Tool:     bash
    Steps:    감사기 fixture 디렉터리에 가짜 전화번호·서명 token·POST endpoint·IPC payload를 만들고 `npm run test:unit -- privacy-audit` 실행
    Expected: 네 유형을 모두 탐지하고 실제 산출물에는 동일 탐지 0
    Evidence: <attemptDir>/task-13-privacy-audit-error.txt
  ```

  Commit: YES | Message: `test(security): audit privacy across web and desktop artifacts` | Files: [`webapp/scripts/audit-privacy.mjs`, `webapp/scripts/audit-privacy.test.mjs`]

- [ ] 14. 승인 기반 Cloudflare preview 배포·회귀 확인·프로젝트 상태 갱신

  What to do: 먼저 사용자에게 `uniform-purchase-guide` Pages preview 배포 승인과 공개 URL 노출 범위를 명시해 확인받는다. 승인된 경우 `dist-web`을 Direct Upload preview branch로 배포하고 HTTPS root/refresh/asset/입력/다운로드/print/network를 재검증한다. production 승격은 preview 결과를 제시하고 별도 명시적 승인을 받은 뒤만 한다. 승인 거절 시 배포하지 않고 Task를 BLOCKED로 기록한다. 성공 후 ADR-002의 “외부 호스팅 제외”를 “Electron offline 주 배포 + Cloudflare static 편의 채널”로 supersede하고 DEPLOY/PRD/계획/task/test/checklist/log/quality gate를 실제 증빙 상태로 갱신한다.
  Must NOT do: 승인 없이 login/project create/deploy를 실행하지 말 것. Cloudflare에 실제 업무자료·테스트 XLSX/HWPX/PDF를 업로드하지 말 것. NOT_RUN을 PASS로 바꾸지 말 것. production alias를 preview 승인과 묶어 처리하지 말 것.

  Parallelization: Can parallel: NO | Wave 4 | Blocks: [F1, F2, F3, F4] | Blocked by: [11, 12, 13]

  References:
  - Pattern:  `ADR-002-웹앱아키텍처.md:24-26` - 서버/DB·업무자료 미전송의 유지할 결정
  - Pattern:  `ADR-002-웹앱아키텍처.md:46-50` - 새 사용자 요구로 supersede할 외부 호스팅 제외 결정
  - Pattern:  `웹앱_계획.md:43-63` - WP-03 실행·품질·개인정보 경계
  - Pattern:  `웹앱_task.md:23-43` - WP-03-12/13 상태 및 새 WP-03-14 추가 위치
  - Test:     `웹앱_test.md:16-65` - PASS 갱신은 증빙이 있는 항목만 허용
  - Pattern:  `웹앱_체크리스트.md:19-35` - 구현/배포/공통 gate
  - Pattern:  `webapp/DEPLOY.md:1-20` - 기존 로컬 서버 배포 안내를 이중 채널로 개정할 파일
  - External: `https://developers.cloudflare.com/pages/get-started/direct-upload/` - preview Direct Upload
  - External: `https://developers.cloudflare.com/pages/how-to/use-direct-upload-with-continuous-integration/` - token/CI를 도입할 경우의 공식 경계; 이번에는 수동 승인형 배포만 사용

  Acceptance criteria:
  - [ ] 배포 호출 직전 사용자 승인 메시지/시각이 evidence에 있으며, 승인 없이는 외부 state change가 0건이다.
  - [ ] 승인 시 `npx wrangler pages deploy dist-web --project-name=uniform-purchase-guide --branch=release-candidate`가 HTTPS preview URL을 반환한다.
  - [ ] preview에서 same-origin static GET 외 요청, request body, API/Functions 호출이 0건이고 입력/출력 flow가 PASS한다.
  - [ ] production은 별도 승인 없이는 실행하지 않았고, 실행했다면 preview와 동일 artifact SHA-256임이 기록된다.
  - [ ] `웹앱_task.md`에 WP-03-14, `웹앱_test.md`에 B-02/B-03/C-05·direct print·Electron·Pages test ID가 추가되고 실제 결과만 PASS다.
  - [ ] `.claude/quality-gate.json`과 `.claude/hooks/verify-phase.ps1 -Mode static`이 새 task/test ID를 강제하며 gate가 0 exit다.

  QA scenarios:
  ```
  Scenario: 승인된 Pages preview 검증
    Tool:     browser:control-in-app-browser
    Steps:    사용자 승인 후 `cd webapp && npx wrangler pages deploy dist-web --project-name=uniform-purchase-guide --branch=release-candidate`; 반환 URL을 열어 root→입력→F-008 preview→PDF/HWPX/바로 인쇄 spy→새로고침을 실행
    Expected: HTTPS 200, asset 404/CSP 오류 0, 업무값 outbound request 0, 출력 정상
    Evidence: <attemptDir>/task-14-pages-preview.zip

  Scenario: 무승인/production 보호
    Tool:     bash
    Steps:    배포 승인 기록이 없는 dry-run fixture에서 release script 실행; preview 승인만 있는 상태에서 production branch 실행 시도
    Expected: 둘 다 배포 전 nonzero로 중단하고 `명시적 배포 승인이 필요합니다.` 출력, Cloudflare state change 0
    Evidence: <attemptDir>/task-14-pages-preview-error.txt
  ```

  Commit: YES | Message: `docs(delivery): record verified desktop and pages release gates` | Files: [`ADR-002-웹앱아키텍처.md`, `웹앱_prd.md`, `웹앱_계획.md`, `웹앱_task.md`, `웹앱_test.md`, `웹앱_체크리스트.md`, `웹앱_로그.md`, `webapp/DEPLOY.md`, `.claude/quality-gate.json`]

## Final verification wave (MANDATORY - after all implementation tasks)
> Runs in PARALLEL. ALL must APPROVE. Surface results to the caller and wait for an explicit "okay" before declaring complete.
- [ ] F1. Plan compliance audit - every task done, every acceptance criterion met
- [ ] F2. Code quality review - diagnostics clean, idioms match, no dead code
- [ ] F3. Real manual QA - every QA scenario executed with evidence captured
- [ ] F4. Scope fidelity - nothing extra shipped beyond Must-Have, nothing Must-NOT-Have introduced

## Commit strategy
- One logical change per commit. Conventional Commits (`<type>(<scope>): <subject>` body + footer).
- Atomic: every commit builds and passes tests on its own.
- No "WIP" / "fix typo squash later" commits on the final branch - clean up before merge.
- Reference the plan file path in the final commit footer: `Plan: .omo/plans/webapp-local-and-static-delivery.md`.
- Cloudflare preview/production deploy는 commit과 별도 승인 이벤트로 기록하고, deploy URL·artifact SHA-256·승인 시각을 `<attemptDir>`에 남긴다.
- NSIS가 unsigned이면 commit/release note에 `INTERNAL_QA_UNSIGNED`를 넣고 외부 배포하지 않는다.

## Success criteria
- All Must-Have shipped; all QA scenarios pass with captured evidence; F1-F4 approved; commit history clean.
- B-02/B-03/C-05의 RED가 재현된 뒤 GREEN으로 전환되고, 잘못된 XLSX import 값이 어떤 출력 경로도 통과하지 않는다.
- `PDF 저장`과 `바로 인쇄`가 독립 동작하며 브라우저와 설치형 모두 `.paper` 원문 전체만 인쇄한다.
- F-050은 미리보기/PDF/브라우저 인쇄/Electron 인쇄 모두 빈 설문 양식을 유지한다.
- Electron installer는 x64 NSIS로 offline 설치·실행·인쇄·제거가 검증되고, 사용자 output을 삭제하지 않는다.
- Cloudflare Pages는 승인된 정적 artifact만 배포하며 Functions/DB/API/analytics/업무자료 전송이 0건이다.
- 원본 XLSM/HWPX SHA-256 5개가 모두 유지되고 개인정보·시크릿·절대경로가 dist/installer/log/evidence에 없다.
