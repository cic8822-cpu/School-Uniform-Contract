# WP-02 마스터 데이터 재검토 R2

- **리뷰 일시**: 2026-09-25
- **범위**: 이전 `REVIEW_WP02.md`의 H-1~H-5·M-1 및 현재 `data/*.json`, `template-mapping.md`, `scripts/verify_web_wp02.ps1`, Excel 절차 원천, `서식_매핑표.md`, `입력데이터_사전.md`, `웹앱_prd.md`
- **변경 상태**: WP-02 데이터와 검증기는 아직 Git 미추적 파일이다. 따라서 이전 리뷰가 가리킨 결함뿐 아니라 현재 파일 전체를 직접 읽어 판정했다.
- **codeQualityStatus**: `BLOCK`
- **recommendation**: `REQUEST_CHANGES`
- **reportPath**: `_workspace/05_web/REVIEW_WP02_R2.md`
- **blockers**: H-2(20단계 주의사항의 매뉴얼 조항 인용 부재), H-5(F-042 HWPX `제출기한` Field 선언 불일치)

## 결론

이전 지적 중 **H-1, H-3, H-4, M-1은 해소**됐다. H-5는 Form별 반영범위 모델을 도입한 점은 맞지만 F-042의 실제 토큰 Field가 Form 입력·참조 계약에서 빠져 있어 **완전 해소되지 않았다**. H-2도 체크리스트·주의사항 배열은 추가됐으나, PRD가 요구하는 주의사항의 매뉴얼 조항 번호 인용은 아직 제공하지 않는다. 따라서 WP-02를 승인하거나 WP-03 구현으로 진행할 수 없다.

## 검증 결과

| 확인 | 결과 | 근거 |
|---|---|---|
| 대상 JSON 파싱 및 기존 검증기 | PASS | `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify_web_wp02.ps1` 종료 코드 0 |
| H-1 계산/직접입력 분리 | PASS | 52개 Form에서 `inputFieldIds ∩ computedFieldIds = ∅`; Field Master의 계산값 K-01~K-04가 직접입력 배열에 없음. 검증기에도 회귀 검사 있음 (`scripts/verify_web_wp02.ps1:112-116`) |
| H-3 미확정 계약방법의 표기 | PASS (데이터) | 19개 빈 `contractMethods` Form 모두 `manual-confirmation-required`와 `계약방법 필터에서는 확인 필요로만 표시` 계약을 가짐 |
| H-4 일반경쟁입찰의 오안내 차단 | PASS | CM-04는 `reference-only`, `actionable: false`이고 WF-BID도 CM-04를 참고 전용으로 한정한다 (`data/contract-methods.json:193-200`, `data/workflows.json`의 `contractMethodScope`) |
| H-5 HWPX 반영범위 모델 | PARTIAL | 48개 템플릿에 `hwpxCoverage`가 추가됐으나 F-042의 `제출기한` 계약 불일치가 남음 |
| M-1 F-050 D-05 경계 | PASS | F-050 원천 매핑에서 D-05가 제거됐고 자유서술 금지로 명시됐다 (`서식_매핑표.md:67`). Form·보호정책 허용값도 C-01/C-02/D-02뿐이다 (`data/forms.json:5054-5058`, `data/fields.json:1581-1637`) |

## CRITICAL

없음.

## HIGH

### H-2. 주의사항의 매뉴얼 조항 인용이 아직 없어 WFR-03을 충족하지 못함

`웹앱_prd.md:41`은 P0인 WFR-03에서 주의사항에 매뉴얼 조항 번호까지 인용하도록 요구한다. 그러나 수의계약 9개 단계의 `manualReferences`는 모두 `계약방법별 수의계약 전용 단계 미확인` 및 `manual-confirmation-required`일 뿐이다. 예를 들어 첫 단계도 같은 상태를 명시한다 (`data/workflows.json:50-67`). 체크리스트와 주의사항의 `sourceRefs`도 매뉴얼이 아닌 Excel 팝업 소스만 가리킨다.

2단계 입찰 11개 단계에는 절차 단락 수준의 근거가 추가됐지만, 각 `cautions` 항목 자체는 여전히 `scripts/build_excel_v1_vba.ps1`만 출처로 둔다. Excel 원천은 절차 문구를 보존한 근거일 수는 있어도 PRD가 요구한 매뉴얼 조항 번호의 대체물이 아니다. 현 검증기는 각 배열이 비어 있지 않은지만 검사하므로 (`scripts/verify_web_wp02.ps1:65-72`), 모든 인용이 미확정·비특정 상태여도 통과한다.

**수정 요구**: 20개 단계의 각 주의사항에 실제 매뉴얼 번호·쪽·조항/문단 locator와 `source-confirmed` 상태를 넣고, 검증기에서 P0 단계의 locator가 빈 값·`manual-confirmation-required`가 아닌지 검사할 것. 원천에서 이 정보를 확인할 수 없다면 WFR-03의 P0 수용 기준을 변경하는 것은 제품 범위 결정이므로 구현자가 임의로 완화하지 말고 사용자 결정을 받아야 한다.

### H-5. F-042가 HWPX 토큰 `제출기한`을 사용하지만 B-07을 Form Field 계약에서 누락함

`template-mapping.md:32,92`는 F-042 템플릿의 `제출기한` 토큰이 B-07임을 명시한다. Form 자체도 `hwpxTokens`에 `제출기한`을, `hwpxCoverage.supportedFieldIds`에 B-07을 넣었다 (`data/forms.json:4180-4186`, `4209-4215`). 하지만 F-042의 `inputFieldIds`와 `referencedFieldIds`에는 B-07이 없다 (`data/forms.json:4140-4157`, `4190-4207`).

따라서 Form별 기초자료 완성도·누락값·자동반영 대상을 `inputFieldIds`/`referencedFieldIds`로 계산하는 WP-03은 B-07을 이 서식의 필요값으로 알 수 없다. 반대로 HWPX 토큰 치환은 그 값을 사용한다. 이는 같은 Field를 두 계약이 다르게 표현하는 현재의 실제 불일치이며, WFR-05·WFR-07·WFR-09의 입력 반영과 출력 가능 상태를 신뢰할 수 없게 한다.

**수정 요구**: B-07을 F-042의 적절한 Form Field 계약(직접입력·참조·HWPX 전용 필요 Field 중 설계에 맞는 하나)에 명시하고, 필요성/공란 허용 정책을 정의할 것. 이어서 검증기에 실제 템플릿 토큰→Field ID 매핑이 Form의 입력·참조 또는 명시적 HWPX 전용 필요 Field에 존재하는지 전수 검사할 것. 현재 검증기는 `hwpxCoverage`의 비어 있지 않음만 확인하므로 이 불일치를 잡지 못한다 (`scripts/verify_web_wp02.ps1:118-124`).

## MEDIUM

### M-2. 회귀 검증이 H-3/H-5의 핵심 UI 계약을 값 수준에서 보호하지 않음

현재 검증기는 미확정 계약방법 Form에 `contractMethodScope.status`만 있는지 확인한다 (`scripts/verify_web_wp02.ps1:125-132`). 현재 데이터에는 올바른 `filterBehavior`가 있으나, 그 값·표시문구·사용자 차단 동작은 회귀 검사하지 않는다. HWPX도 coverage 객체의 존재만 본다. 이 공백 때문에 위 F-042 오류가 PASS했다.

**수정 요구**: 미확정 Form은 `manual-confirmation-required`와 “확인 필요로만 표시” 동작을 함께 검증하고, HWPX coverage는 토큰↔Field 매핑·지원/미지원 Field의 상호 배타성·Form Field 계약 정합성을 검증할 것.

## LOW

없음.

## 이전 지적별 판정

| 이전 지적 | 판정 | 요약 |
|---|---|---|
| H-1 | 해소 | 계산 Field와 직접입력 Field가 서로소이고 `referencedFieldIds`로 참조 관계가 분리됨 |
| H-2 | 미해소 | 배열은 생겼지만 주의사항별 매뉴얼 조항 인용이 없음 |
| H-3 | 해소 | 19개 미확정 Form을 빈 배열의 임의 해석 대신 확인 필요 상태와 필터 동작으로 모델링함 |
| H-4 | 해소 | CM-04에 2단계 입찰 절차를 확정 Workflow로 제시하지 않도록 reference-only 계약을 둠 |
| H-5 | 부분 해소 / 차단 | coverage 모델은 도입됐으나 F-042 B-07 불일치가 존재함 |
| M-1 | 해소 | F-050 D-05가 JSON과 원천 매핑 모두에서 허용 대상이 아니며 금지 경계가 명시됨 |

## 스킬 관점 점검

`remove-ai-slops` 및 `programming`은 제공된 스킬 루트에서 해당 `SKILL.md`를 찾지 못해 실행할 수 없었다. 대신 요청문에 명시된 기준으로 점검했다. 대상은 데이터와 검증기이며, 삭제만 검증하는 테스트·구현 상수만 그대로 따라 하는 테스트·불필요한 생산 파싱/정규화·무타입 우회·불필요 추상화는 발견하지 못했다. 다만 H-2/H-5의 검증은 존재 여부만 확인하는 얕은 게이트라서, 실제 업무 계약의 회귀를 막지 못한다. 이는 slop이 아니라 검증 범위 부족으로 M-2에 기록했다.

## 승인 조건

1. H-2의 실제 매뉴얼 조항 인용과 그 회귀 검증을 추가한다.
2. H-5의 F-042 B-07 Form Field 계약을 정합화하고, 토큰→Field 전수 정합성 검증을 추가한다.
3. `scripts/verify_web_wp02.ps1` 및 추가된 정합성 검증이 모두 통과한 증빙을 남긴다.

위 두 HIGH가 해소된 뒤에만 WP-02를 승인할 수 있다.

## 보충 재검토 (F-042 H-5, 2026-09-25)

F-042는 이제 `referencedFieldIds`와 `hwpxRequiredFieldIds`에 B-07을 모두 선언하고, `hwpxCoverage.supportedFieldIds` 및 `제출기한` 토큰과도 일치한다. 검증기는 48개 템플릿 전부에 대해 토큰→Field ID가 `inputFieldIds`·`referencedFieldIds`·`hwpxRequiredFieldIds` 중 하나에 선언됐는지, 그리고 지원 Field 집합이 실제 토큰 집합과 같은지를 검사한다 (`scripts/verify_web_wp02.ps1:133-140`). 대상 검증기 재실행도 PASS했다.

따라서 **H-5는 해소**됐고, 이 보충 판정은 위 본문의 H-5 차단 판정을 대체한다. 현재 남은 HIGH 및 유일한 승인 차단 항목은 **H-2(단계별 주의사항의 실제 매뉴얼 조항 인용)**뿐이다. 최종 상태는 여전히 `BLOCK` / `REQUEST_CHANGES`다.

## 보충 재판정 (H-2 기준 정정, 2026-09-25)

기존 H-2는 생성 문서였던 `웹앱_prd.md`의 “매뉴얼 조항 번호까지 인용”을 최상위 요구로 취급해 차단으로 판정했다. 사용자 제공 원문을 다시 대조한 결과, `0-1.교복계약 웹앱개발-프롬프트.txt:11-15, 285-291`은 단계별 **관련 규정/주의사항** 표시를 요구할 뿐 조항 번호 형식을 강제하지 않는다. `0-2.교복 학교주관구매 웹앱 — 화면 구성 및 디자인 계획.md:18, 73-84`는 불확실한 사항을 “매뉴얼 확인 필요”로 명시하고, 수의계약의 적용 조건도 매뉴얼 확인 대상으로 둔다.

개정 WFR-03은 이에 맞춰 “매뉴얼 장절 **또는** Excel 계약절차안내 원문 위치”와 미확정 시 확인 필요 표기를 수용 기준으로 한다 (`웹앱_prd.md:41`). 현재 데이터는 20개 단계 모두 체크리스트·주의사항의 원문 위치를 `sourceRefs`로 갖고, CM-03/WF-BID는 `3 추진 절차 개요 및 교복 납품 사업자 선정`을 `source-confirmed`로 표시한다. CM-01·CM-02의 9개 단계는 Excel 안내 원문 이식임과 수의계약 전용 매뉴얼 단계 미확정을 `manual-confirmation-required`로 명시한다. 이는 [WP02_절차근거_대조.md](WP02_절차근거_대조.md)의 단계 데이터 규칙·표시 정책과도 일치한다. 직접 전수 점검에서 20개 단계의 checklist/caution `sourceRefs`와 `manualReferences`에 빈 document·locator·status는 0건이었고, `scripts/verify_web_wp02.ps1`도 PASS했다.

따라서 **H-2는 해소**됐고, 이 보충 판정은 위 본문의 H-2 차단 판정을 대체한다. 현재 HIGH/CRITICAL 차단 항목은 없다. 최신 판정은 `codeQualityStatus: WATCH`, `recommendation: APPROVE`, `blockers: 없음`이다.

**비차단 관찰(M-3)**: `verify_web_wp02.ps1:65-72`는 단계별 근거 배열의 존재만 검사한다. 현재 데이터의 sourceRefs·상태값은 직접 전수 대조로 확인됐지만, 이후 회귀를 자동 차단하려면 각 단계에 빈 sourceRefs/locator/status가 없는지와 `manual-confirmation-required`일 때 확인 필요 표시 계약이 있는지를 검증기에 추가하는 것이 바람직하다.
