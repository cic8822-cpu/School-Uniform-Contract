# WP-03 코드 품질 독립 검토

- 검토 대상 커밋: `8f8fffeb4300b96853e760da87a8a51b9fdb6404` (`fix: HWPX 원문 미리보기 및 작성 모달 구현`)
- 범위: `webapp/src/components/modal/DocumentPreviewModal.tsx`, `webapp/src/lib/hwpxPreview.ts` 및 읽기 전용 연관 코드 `hwpxExport.ts`, `hwpx.ts`, 타입·개인정보 메타데이터
- 범위 제외: 사용자 변경인 `logs/xvba_debug.log`와 문서 변경
- 독립 판정: **BLOCK / REQUEST_CHANGES**

## 검토 방법과 근거

기존 실행 보고서는 신뢰하지 않고 코드와 원본 정적 자산을 다시 대조했다.

- 커밋 diff는 대상 두 파일뿐이다(추가 217줄, 삭제 34줄). `git diff --check 8f8fffeb4300b96853e760da87a8a51b9fdb6404^ 8f8fffeb4300b96853e760da87a8a51b9fdb6404`는 PASS했다.
- `webapp`에서 `npm run lint`와 `npm run build`를 직접 실행했고 모두 종료 코드 0으로 PASS했다.
- `.claude/hooks/verify-phase.ps1`은 프로세스 종료 코드만 0이었지만, 출력상 WE/WM/WA/WF 검사 다수 `NOT_RUN`, `WP-03-12`·`WP-03-13` `IN_PROGRESS`, 체크리스트 `WP-03` 미완료를 명시했다. 따라서 활성 단계 품질 게이트를 통과했다는 증거는 아니다.
- 테스트 파일 및 테스트 스크립트는 없다. `webapp/package.json`의 스크립트는 `dev`, `build`, `lint`, `preview`뿐이고 Vitest/Bun 실행 파일도 없다.
- 실제 배포 템플릿 48개를 ZIP/XML로 다시 읽었다. F-008, F-013, F-047, F-050의 `cellSpan`/`rowSpan`을 HTML 격자로 재구성했을 때 모든 행 폭이 일관되었다(F-008: 3/4/2열 표, F-013: 11개 표, F-047: 1/3열 표, F-050: 1/7/1열 표). 현재 코드는 이 범위의 병합값을 읽는다.
- F-029/F-047/F-050의 실제 템플릿 토큰과 `forms.json`을 대조했다. 각각 `학교명, 학년도` / `학교명` / `학교명`만 있고, 모든 토큰은 각 `privacyAllowedFieldIds` 안의 공통 필드에 매핑된다. 금지 패턴(성명·연락처·서명·설문 등) 토큰은 없었다. 따라서 **현행 정적 메타데이터 기준으로** 이번 미리보기 치환이 실제 개인정보 값을 화면이나 HWPX 생성물에 넣는 경로는 발견하지 못했다.
- 48개 현재 템플릿의 section XML에는 `pic`, `ole`, `video` 시각 컨트롤이 없었다(각 파일의 `BinData` 46개는 존재하지만 section에서 참조되는 해당 컨트롤은 없음). 그러므로 이미지 플레이스홀더 경로는 현재 자산으로는 검증되지 않았다.

## 발견 사항

### CRITICAL

없음.

### HIGH

#### H-01 — 원문 로드 완료 전 PDF가 로딩 문구만 성공적으로 내려받아진다

- 위치: `webapp/src/components/modal/DocumentPreviewModal.tsx:147-149`, `:222-225`
- 재현: HWPX 템플릿 서식(F-008 등)을 열고 `loadHwpxPreview()`의 fetch/ZIP/XML 비동기 작업이 끝나기 전에 `PDF 출력`을 누른다.
- 근거: 초기 상태에서 `sourceBlocks === null`, `sourceError === null`, `busy === false`다. 화면은 `원문 HWPX를 불러오는 중입니다.`만 `.paper`에 넣지만, PDF 버튼은 `disabled={busy}`뿐이라 활성 상태다. `handlePdfExport()`는 즉시 그 DOM을 `exportElementToPdf()`에 넘긴 뒤 성공 메시지를 표시한다.
- 영향: 사용자는 원문 미리보기/PDF가 준비됐다고 믿고 로딩 문구만 포함한 잘못된 공문 PDF를 내려받는다. 이번 변경 전에는 동기적 일반 미리보기가 즉시 존재했으므로 이번 비동기 전환이 만든 회귀다.
- 수정 조건: HWPX 원문 블록이 준비될 때까지 PDF 출력을 비활성화하거나, 완료 상태를 `handlePdfExport()`에서도 검사해 준비 전에는 생성하지 않아야 한다. 실패 상태의 PDF 정책도 명시해야 한다.

### MEDIUM

#### M-01 — 첫 일시적 로드 실패가 탭 전체에서 영구 실패로 캐시된다

- 위치: `webapp/src/lib/hwpxPreview.ts:21`, `:85-103`, `:115-116`
- 재현: 첫 fetch를 일시적으로 실패(정적 서버 재시작, 일시 네트워크 오류, 503 등)시킨 뒤 같은 탭에서 서버를 복구하고 모달을 다시 열거나 입력값을 바꾼다.
- 근거: `sourceCache`가 `loadSource()`의 Promise를 즉시 저장한다. reject된 Promise는 catch에서 제거되지 않으며 이후 호출도 같은 reject Promise를 받아 `loadSource()`를 재시도하지 않는다.
- 영향: 정상으로 복구된 템플릿도 새로고침 전까지 계속 “HWPX 원문을 표시하지 못했습니다.”로만 남는다. 사용자 복구 수단이 없다.
- 수정 조건: 캐시 Promise가 reject되면 해당 key를 제거하고, 재시도 가능한 UI 또는 적어도 재-open 시 재요청을 제공해야 한다.

#### M-02 — 새 ZIP/XML 파서·병합·보호 서식 경계에 자동 회귀 검사가 없다

- 위치: `webapp/src/lib/hwpxPreview.ts:23-127`, `webapp/src/components/modal/DocumentPreviewModal.tsx:60-84,147-168`
- 근거: 이 커밋은 새 파서 127줄과 모달 렌더링 변경 90줄을 추가했지만 테스트 파일·테스트 스크립트가 전혀 없다. 수동 자산 대조로 현재 표 병합·보호 토큰은 확인했으나, 그 결과를 지속적으로 고정하지 않는다.
- 영향: `cellSpan`/`rowSpan`, 다중 section 순서, XML 파싱 실패·재시도, F-029/F-047/F-050의 허용 토큰만 치환하는 경계가 이후 무음 회귀할 수 있다.
- 수정 조건: DOMParser/JSZip을 대상으로 최소한 (1) 병합 표, (2) 다중 section 자연 정렬, (3) 실패 후 재시도, (4) 보호 서식에 금지 토큰값이 전달되지 않음을 관찰 결과로 검증하는 테스트를 추가해야 한다. 구현 상수만 복제하는 테스트는 인정하지 않는다.

#### M-03 — HWPX 템플릿이 없는 서식은 `aria-busy`가 영구 true다

- 위치: `webapp/src/components/modal/DocumentPreviewModal.tsx:147`
- 재현: `hwpxTemplate: null` 서식을 연다.
- 근거: effect는 `:67`에서 바로 return하여 `sourceBlocks`와 `sourceError`는 모두 null로 남는다. 그런데 `aria-busy={sourceBlocks === null && !sourceError}`는 true다.
- 영향: 화면은 일반 미리보기를 표시해도 보조기술에는 문서 영역이 로딩 중으로 남는다. WP-03 접근성 목표와 맞지 않는다.
- 수정 조건: `aria-busy`를 `Boolean(form.hwpxTemplate) && sourceBlocks === null && !sourceError`처럼 HWPX 경로에만 적용한다.

### LOW

#### L-01 — 원문 강제 줄바꿈을 보존하지 않는다

- 위치: `webapp/src/lib/hwpxPreview.ts:36-45`
- 재현: F-008 템플릿의 `-시보리: 카라/소매끝/밑단` 문단은 `<hp:t>` 안에 `<hp:lineBreak/>`를 갖는다. 현재 `textContent`를 이어 붙이면 줄바꿈이 제거되어 한 문단으로 표시된다.
- 영향: 텍스트 자체는 남지만 원문 단락/줄 배치 충실도가 떨어진다. F-005·F-013에도 같은 컨트롤이 있다.
- 수정 조건: `lineBreak`을 `\n` 또는 `<br>`로 명시적으로 변환하고, 렌더링 CSS의 줄바꿈 정책을 검증한다.

## remove-ai-slops / programming 관점 확인

두 skill을 실제로 읽고 적용했다.

- `remove-ai-slops`: **위반 있음(M-02)**. 새 동작에 회귀 검사가 없어 behavior coverage 범주를 충족하지 못한다. 반면 삭제 전용·요청사항 문구만 확인하는 tautological 테스트는 없었고(테스트 자체가 없음), 원문 HWPX 표시라는 목표에 필요한 ZIP/XML 파싱·토큰 치환은 불필요한 production 데이터 추출/정규화로 보지 않았다.
- `programming`: **부분 위반/주의 있음**. `loadSource()`가 timeout·재시도 정책 없는 bare `fetch()`에 의존하고 M-01의 실패 캐시 문제를 만든다. 새 `HwpxPreview*` export는 변하지 않아야 하는 전달 데이터인데 mutable 인터페이스/배열로 선언돼 있으며, `DocumentPreviewModal.tsx`는 pure LOC 215로 200~250 경고 구간이다. 다만 `any`, `as`/non-null assertion, `@ts-ignore`, 빈 catch, 불필요한 추상화, 또는 목표 밖의 검증·파싱은 발견하지 못했다. XML은 정적 HWPX라는 기능 경계에서 실제로 필요하다.

## 판단

- `codeQualityStatus`: **BLOCK**
- `recommendation`: **REQUEST_CHANGES**
- `blockers`:
  1. H-01을 수정해 원문 준비 전/실패 상태에서 잘못된 PDF가 생성되지 않게 할 것.
  2. M-01의 reject Promise 캐시를 제거 또는 재시도 가능하게 할 것.
  3. M-02의 관찰 가능한 회귀 테스트를 추가하고 통과 증적을 남길 것.

H-01이 남아 있는 한 lint/build 통과만으로는 승인이 불가능하다.
