# WP-03 최종 코드 품질 검토

- 검토 대상: `8f8fffeb4300b96853e760da87a8a51b9fdb6404`, `3800c6e95583e98a67f7a6d0f9e7e1bfd83d93f7`
- 목표: webapp의 HWPX 원문 미리보기·HWPX 작성 모달이 원문 문단/표를 표시하고, PDF 로딩·오류를 차단하며, 누락 토큰·Esc 포커스·캐시 재시도·개인정보 경계를 지키는지 독립 검토한다.
- 범위: `webapp/src/components/modal/DocumentPreviewModal.tsx`, `webapp/src/lib/hwpxPreview.ts`, `webapp/scripts/verify-hwpx-preview-cache.mjs` 및 읽기 전용 연관 코드.
- 산출물 위치: 호출자가 지정한 본 경로. `omo ulw-loop status --json`은 이 환경에서 `The syntax of the command is incorrect.`로 실패하여 attempt 경로를 신뢰성 있게 얻지 못했다.

## 독립 검증 결과

기존 보고서와 증빙은 신뢰하지 않고 코드·48개 정적 HWPX·실행 결과를 다시 확인했다.

- `npm run lint`, `npm run build`, `node scripts/verify-hwpx-preview-cache.mjs`는 모두 종료 코드 0으로 통과했다. 마지막 스크립트는 실제 실패 fetch를 두 번 호출하여 reject Promise가 캐시에서 제거되는지를 검증했고 `fetchCalls: 2`를 출력했다.
- 두 커밋 각각의 `git diff --check` 및 현재 워크트리의 `git diff --check`는 통과했다. 현재 워크트리의 다른 문서·로그·`.omo`·QA 변경은 본 검토 범위에서 제외했다.
- 브라우저에서 F-020을 열어 누락 토큰 시 HWPX 작성이 `학교명` 누락 안내만 표시하고 다운로드를 시도하지 않음을 확인했다. F-029/F-047/F-050은 개인정보 안내가 있고 모달 내 `input`/`textarea`/`select`가 0개였으며, 원문 표만 표시했다. F-050에서 Esc를 누르면 모달을 연 F-050 버튼으로 포커스가 복귀했다.
- 코드상 `sourceBlocks === null`이면 HWPX 서식의 PDF 버튼을 비활성화하므로 로딩·오류 상태 모두 PDF 생성 경로가 차단된다. 단, 아래 HIGH 결함 때문에 로드가 완료된 뒤의 PDF 내용 충실성은 승인할 수 없다.
- 활성 Phase 훅은 종료 코드는 0이지만 WE/WM/WA/WF 다수가 `NOT_RUN`, WP-03-12/13이 `IN_PROGRESS`라고 보고했다. 이는 이 커밋의 코드 컴파일 결과가 아니라 Phase 완료 게이트가 아직 닫히지 않았다는 뜻이다.

## 발견 사항

### CRITICAL

없음.

### HIGH

#### H-01 — 표 안의 원문 문단과 중첩 표 구조를 평탄화하여, 미리보기와 그 PDF가 원문 구조를 재현하지 못함

- 위치: `webapp/src/lib/hwpxPreview.ts:48-52, 76-80`, `webapp/src/components/modal/DocumentPreviewModal.tsx:170-180`, `webapp/src/components/modal/modal.css:32-50`
- 재현:
  1. 서식함에서 F-020을 연다.
  2. 원문은 깊이 0의 바깥 표 안에 깊이 1의 주의사항 표가 있는 구조다.
  3. 미리보기 DOM의 `.paper` 직계 자식은 두 개의 형제 `TABLE`이며, 중첩 표는 바깥 `TD` 안에 남아 있지 않다. 바깥 표의 한 `TD` 안에 있던 20개 원문 `hp:p`도 하나의 문자열로 표시된다.
- 근거: `tableCellText()`는 모든 셀 문단을 `join('\n')`한 문자열 하나로 반환하고, 렌더러는 그 문자열을 단일 `<td>`에 그대로 넣는다. 기본 계산 스타일은 `white-space: normal`이므로 줄바꿈은 문단 블록으로 렌더링되지 않는다. 또한 `parseSection()`은 모든 하위 `tbl`을 수집해 형제 블록으로 밀어낸다.
- 범위 측정: 48개 배포 템플릿을 ZIP/XML로 전수 확인한 결과 31개 템플릿에 2개 이상 `hp:p`를 담은 셀이 201개 있으며, 한 셀 최대 문단 수는 32개다. F-020/F-021/F-022/F-023/F-024/F-027/F-031/F-047/F-048에는 중첩 표도 있다.
- 영향: 이번 변경의 핵심인 “원문 전체 문단/표 렌더링”이 표 내부에서 성립하지 않는다. 특히 F-020, F-029, F-047 등은 사실상 한 셀에 공문 본문을 담으므로 문단 경계·중첩 주의 표의 위치가 사라진다. PDF 출력은 이 평탄화된 DOM을 캡처하므로 원문과 다른 문서가 생성된다.
- 수정 조건: 표 셀을 문자열이 아니라 문단 블록 목록으로 모델링·렌더링하고, 중첩 표는 해당 셀 안에서 재귀 렌더링해야 한다. 같은 수정에서 강제 줄바꿈도 보존해야 한다.

### MEDIUM

#### M-01 — 새 HWPX 파서의 실제 문서 구조를 지키는 회귀 검사가 없음

- 위치: `webapp/scripts/verify-hwpx-preview-cache.mjs:14-24`
- 근거: 새 스크립트는 캐시 실패 재시도만 실제로 검증하며, 실제 템플릿의 다중 문단 셀·중첩 표·`cellSpan`/`rowSpan`·`lineBreak`·보호 서식 토큰 경계를 렌더 결과로 검증하지 않는다. 이 때문에 H-01이 lint/build/cache 검사 모두 통과한 상태로 남았다.
- 영향: 템플릿·파서 변경 때 원문 미리보기/PDF의 구조 손실이 다시 유입돼도 감지할 수 없다.
- 수정 조건: 구현 상수를 복제하지 말고 실제 배포 템플릿을 대상으로 F-020(중첩 표), F-029 또는 F-047(다중 문단 보호 서식), F-008(`lineBreak`)의 DOM 구조·텍스트 경계·PDF 버튼 상태를 관찰하는 회귀 검사를 추가한다.

### LOW

#### L-01 — HWPX의 `hp:lineBreak` 세 곳을 텍스트 공백으로 잃음

- 위치: `webapp/src/lib/hwpxPreview.ts:36-45`
- 근거: `textContent` 수집은 자식 `hp:lineBreak`를 줄바꿈으로 변환하지 않는다. F-005, F-008, F-013에 각각 1개가 있으며, F-008의 `-시보리: 카라/소매끝/밑단<hp:lineBreak/>진남색/흰색 2줄`은 미리보기에서 한 줄로 합쳐졌다.
- 영향: 내용 자체는 남지만 원문의 강제 줄바꿈이 사라진다. H-01 수정 때 함께 해결할 수 있다.

## remove-ai-slops / programming 관점

`remove-ai-slops`와 `programming` 스킬은 제공된 사용 가능 목록과 로컬 skill 경로에서 찾을 수 없어 전문을 로드하지 못했다. 따라서 요청문에 적힌 두 스킬의 판단 기준을 직접 적용했다.

- `remove-ai-slops`: 새 캐시 재시도 검사는 삭제 전용·요청 문구 확인·상수 미러링·tautology가 아니라 실제 reject 캐시 동작을 확인하므로 유용한 테스트다. 목표와 무관한 데이터 추출·정규화도 발견하지 못했다. 다만 M-01처럼 핵심 렌더링 동작의 관찰 기반 검사가 없어 회귀 신뢰성은 부족하다.
- `programming`: `any`, `@ts-ignore`, 무타입 escape hatch, 불필요한 추상화, 목표 밖의 검증/파싱은 발견하지 못했다. ZIP/XML 파싱 자체는 원문 미리보기 경계에 필요한 복잡성이다. 그러나 H-01의 평탄화는 필요한 파싱 결과를 UI에서 잘못 모델링한 정확성·유지보수 위반이다.

## 판정

- `codeQualityStatus`: **BLOCK**
- `recommendation`: **REQUEST_CHANGES**
- `blockers`:
  1. **H-01**을 수정해 표 셀의 각 원문 문단, 강제 줄바꿈, 중첩 표의 부모-자식 관계를 보존하고 PDF에도 같은 구조가 반영되게 할 것.
  2. 수정 결과를 실제 HWPX 자산 기반 회귀 검사로 증명할 것. F-020 또는 F-047처럼 현재 평탄화되는 대표 자산이 반드시 포함되어야 한다.

H-01이 해결되기 전에는 컴파일·캐시 재시도·개인정보 경계가 통과하더라도 원문 미리보기/PDF 기능을 승인할 수 없다.
