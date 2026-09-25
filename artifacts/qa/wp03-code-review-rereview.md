# WP-03 HWPX 미리보기 수정 재검토

- 검토 커밋: `3800c6e95583e98a67f7a6d0f9e7e1bfd83d93f7` (`fix: HWPX 원문 미리보기 출력 차단 보완`)
- 목표: 이전 독립 리뷰의 H-01, M-01, M-02, M-03, L-01이 현 소스에서 해소됐는지와 새 회귀 여부를 독립 확인함.
- 범위: `webapp/src/components/modal/DocumentPreviewModal.tsx`, `webapp/src/lib/hwpxPreview.ts`, `webapp/scripts/verify-hwpx-preview-cache.mjs`, 관련 정적 템플릿·마스터 데이터·패키지 스크립트.
- 개인정보 경계: F-029/F-047/F-050의 실제 `privacyAllowedFieldIds`, `protectedForms.allowedFieldIds`, HWPX 토큰과 템플릿 존재 여부를 읽기 전용으로 대조함. 개인식별·연락처·서명·설문 원자료는 조회·기록하지 않음.

## 독립 검증 근거

- `git show --find-renames 3800c6e -- <대상 3파일>`로 수정 범위(3파일, +65/-9)를 확인했고, `git diff --check 3800c6e^ 3800c6e`는 PASS함.
- `webapp`에서 `npm run lint`, `npm run build`, `node scripts/verify-hwpx-preview-cache.mjs`를 직접 실행해 모두 종료 코드 0을 확인함. 마지막 명령의 관찰값은 `{"status":"PASS","fetchCalls":2}`임.
- 정적 데이터와 자산을 재대조함. HWPX 템플릿 선언 48개와 `public/templates` 자산 48개가 일치하고 누락은 없었음. 보호 서식 F-029/F-047/F-050의 폼 허용 필드와 `protectedForms` 허용 필드도 각각 일치했으며, HWPX 토큰은 각각 `학교명, 학년도` / `학교명` / `학교명`뿐임.
- `Contents/section*.xml`을 읽어 실제 원문 강제 줄바꿈이 F-005, F-008, F-013에 남아 있음을 재확인함. 이는 개인정보 값 조회 없이 구조 표식만 센 결과임.
- 활성 단계 훅 `.claude/hooks/verify-phase.ps1`도 종료 코드 0이었음. 다만 이 훅은 새 `verify-hwpx-preview-cache.mjs`를 실행하는 근거가 아니므로, 그 스크립트는 위의 직접 실행 결과만 증적으로 삼음.

## 이전 지적 재판정

| ID | 판정 | 근거 |
|---|---|---|
| H-01 | 해소 | `DocumentPreviewModal.tsx:73-79`가 원문 블록 부재 시 PDF 차단 상태와 안내를 만들고, `:238-245`가 버튼을 `disabled` 처리함. 따라서 로딩 중·로드 실패 중의 일반 사용자 입력으로는 로딩/오류 문구만 PDF로 출력할 수 없음. |
| M-01 | 해소 | `hwpxPreview.ts:115-121`은 새 원문 Promise가 reject되면 동일 Promise인지 확인 후 캐시에서 제거함. `verify-hwpx-preview-cache.mjs:21-24`는 두 번의 실패 호출이 실제 fetch 두 번으로 이어지는 행동을 검증하며 직접 PASS함. |
| M-02 | 부분 해소, 아래 MEDIUM으로 유지 | 새 검사는 M-01의 실패 캐시만 관찰한다. ZIP/XML 파싱·병합 셀·다중 section 순서·PDF 차단 UI·보호서식 토큰 경계는 계속 자동 회귀 검사 대상이 아님. |
| M-03 | 해소 | `DocumentPreviewModal.tsx:73`, `:163`에서 `aria-busy`는 HWPX 템플릿이 있고 성공/오류 어느 쪽도 아닌 로드 중일 때만 true임. 템플릿 없는 일반 서식은 false임. |
| L-01 | 미해소, 아래 LOW로 유지 | `hwpxPreview.ts:36-45`는 `hp:t`의 `textContent`만 연결하므로 `<hp:lineBreak/>`를 줄바꿈으로 변환하지 않음. 실제 F-005/F-008/F-013에 해당 표식이 존재함. |

## 발견 사항

### CRITICAL

없음.

### HIGH

없음.

### MEDIUM

#### M-02 — HWPX 파서와 출력 차단의 핵심 관찰 회귀 검사가 아직 부족함

- 위치: `webapp/src/lib/hwpxPreview.ts:36-103`, `webapp/src/components/modal/DocumentPreviewModal.tsx:73-79,163-165,238-245`, `webapp/scripts/verify-hwpx-preview-cache.mjs:15-24`
- 근거: 새 스크립트는 실패한 fetch가 캐시에 남지 않는지만 검사한다. `package.json`에도 이를 실행하는 `test` 스크립트가 없어서 `npm run lint`·`npm run build`만으로는 실행되지 않는다. 원본 HWPX를 실제로 읽어 병합 셀/다중 section 순서/파싱 오류를 확인하거나, 로드 중·실패 중 PDF 버튼이 비활성화되는 DOM 행태, F-029/F-047/F-050에서 허용 토큰 외의 값을 출력하지 않는 경계를 검증하지 않는다.
- 영향: 이후 파서·렌더·개인정보 경계 변경이 현재 자산에서 무음 회귀할 가능성이 남는다. 이번 커밋의 H-01/M-01/M-03 수정을 즉시 깨뜨리는 결함은 아니므로 비차단으로 분류함.
- 권고 검증: `npm run test` 등 표준 실행 경로에 (1) 실제 병합표·다중 section fixture의 파싱 결과, (2) 지연/실패 로드 중 PDF 비활성화, (3) 보호 서식의 허용 토큰만 치환됨을 관찰 기반으로 등록할 것. 단순 상수 복제 테스트는 인정하지 않음.

### LOW

#### L-01 — 원문 내부 강제 줄바꿈이 미리보기에서 사라짐

- 위치: `webapp/src/lib/hwpxPreview.ts:36-45`
- 근거: `getElementsByTagNameNS('*', 't')`의 `textContent`만 이어 붙인다. F-005/F-008/F-013은 실제 section XML에 `<hp:lineBreak/>`를 포함하지만, 이 제어 문자는 `\n` 또는 `<br>`로 매핑되지 않는다.
- 영향: 텍스트는 보존되나 원문 줄 배치 충실도가 낮아진다. 이번 커밋에서 새로 생긴 회귀는 아님.
- 권고: `lineBreak`를 명시적으로 `\n`으로 변환하고, 해당 세 템플릿의 줄바꿈 보존을 관찰하는 검사를 추가할 것.

## remove-ai-slops / programming 관점 확인

- **실제 skill 로딩 여부: 불가.** 제공된 skill 목록에는 `remove-ai-slops`와 `programming`이 없었고, 로컬 skill 루트(`C:\Users\최익창\.codex\skills`, `C:\Users\최익창\.agents\skills`)에서도 해당 `SKILL.md`를 찾지 못했음. 따라서 요청문에 명시된 두 관점의 기준(관찰 가능한 행동 테스트, 삭제 전용·tautological·구현 상수 복제 검사 배제, 불필요한 파싱/정규화·무형식 우회·취약한 프롬프트 테스트·불필요한 추상화 배제)을 직접 적용함.
- **remove-ai-slops:** 새 캐시 검사는 실패 후 실제 fetch가 두 번 발생하는 행동을 검증하므로 삭제 전용·요청 문구 확인·구현 상수 복제 테스트가 아님. 다만 M-02처럼 핵심 새 동작의 관찰 범위가 불완전한 점은 이 관점에 위반됨. 이번 diff의 캐시 삭제와 PDF 상태 계산 자체는 목표 밖 데이터 추출·정규화가 아님.
- **programming:** `any`, `@ts-ignore`, 빈 catch, 무형식 escape hatch, 불필요한 추상화, 목표 밖 입력 검증/파싱은 발견하지 못했음. 캐시 Promise 삭제도 최소 변경임. 그러나 검증 스크립트가 표준 package script에 연결되지 않아 유지보수 시 실행 누락 위험이 있으며, 이는 M-02의 일부임.

## 최종 판단

- `codeQualityStatus`: **WATCH**
- `recommendation`: **APPROVE**
- `blockers`: 없음. H-01을 포함한 기존 HIGH는 해소됐고 CRITICAL/HIGH 신규 결함도 발견하지 못함. M-02와 L-01은 후속 보완이 필요한 비차단 위험임.
