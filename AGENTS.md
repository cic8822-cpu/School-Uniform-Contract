# 교복구매 길라잡이 — 에이전트 작업 지침

교복 학교주관구매의 Excel XLSM·PDF·HWPX 출력 자동화 프로젝트임. 모든 답변과 작업 기록은 한글로 작성함.

## 시작 전 확인

- 현재 단계와 차단 사유는 [task.md](개발문서/task.md), 상세 계획은 [계획.md](개발문서/계획.md), 최근 재개 정보는 [로그.md](개발문서/로그.md)를 먼저 확인함.
- 원본 `20230808_용역계약갈라잡이(디깅모멘텀)_이행원.xlsm` 및 [교복4차 (2)](교복4차%20(2))의 HWPX는 읽기 전용임. 원본을 덮어쓰거나 원본 VBA·외부 연결·깨진 이름 정의를 이식하지 않음.
- 분석 결과는 `_workspace/`에, 생성·검증 산출물은 `artifacts/`에만 둠.
- 개인정보 가능 서식 `F-029`, `F-047`, `F-050`에는 개인식별·연락처·서명·설문 원자료를 입력·저장·자동반영·로그·Git·웹 전송하지 않음. 상세 경계는 [서식_인벤토리.md](개발문서/서식_인벤토리.md) 및 [입력데이터_사전.md](개발문서/입력데이터_사전.md)를 따름.

## 현재 구현 상태

- 활성 단계는 `P3-01`이며, `P2-01`(F-007·F-024·F-013 독립 HWPX 템플릿의 Kordoc `fill -j` 치환)은 라벨 기반 재구축으로 `fill -j → validate → render → 재파싱` 왕복을 통과해 기술 검증 범위 `DONE`으로 전환됨. 다만 한컴 수동 열기·PDF 비교(`P2-02`)는 아직 `TODO`이므로, HWPX 서식을 실제 반영·배포하는 작업은 P2-02가 통과한 뒤에만 수행함. 최신 상태는 자주 바뀌므로 이 파일에 세부를 하드코딩하지 않고 [task.md](개발문서/task.md)·[로그.md](개발문서/로그.md)를 직접 확인함.
- 클린룸 Excel 배포본은 [artifacts/excel](artifacts/excel)에 `교복구매_길라잡이_YYYYMMDD_vN.xlsm` 형식으로 저장함. 안전 빌드가 통과할 때마다 같은 날짜의 `vN`을 증가시켜 기존 통과본을 보존함. 구현·미구현 Form ID 목록은 자주 바뀌므로 이 파일에 하드코딩하지 않음 — 항상 [test.md](개발문서/test.md)를 직접 확인함. `X-06`이 PASS가 되기 전 P3-01 완료로 보고하지 않음. **다만 X-06의 유일한 잔여 원인(추정금액 입력 시 계약방법 자동판정 마법사 미재현)은 법적 위험으로 2026-09-20 사용자가 영구 미구현을 확정한 항목임 — 향후 세션에서 이를 "남은 작업"으로 다시 구현 시도하지 않음(대체 구현: 4대 유형 고정 선택 버튼 + 명시적 법령확인 + 계약절차안내 팝업, 상세는 `개발문서/test.md` X-06 각주).**
- P5는 아직 구현하지 않음. 웹 UX는 절차 탐색 허브→단계 상세→관련 Form ID→입력·검증·출력 흐름으로 설계하며, 근거·최종 검토 안내와 접근성 기준을 [prd.md](개발문서/prd.md) 및 `_workspace/05_web/학교행정업무길라잡이_웹판_UX_벤치마크.md`에서 관리함.
- 기준 기능, 남은 범위 및 검증 결과는 [test.md](개발문서/test.md), [체크리스트.md](개발문서/체크리스트.md), [P3-01_Excel_MVP_설계검증.md](개발문서/P3-01_Excel_MVP_설계검증.md)를 기준으로 판단함.

## 작업 방식

- 최소 변경 원칙을 적용하고, 수정 전 대상 파일을 읽은 뒤 단계별로 검증함. 실패·불확실성·차단 요인은 [로그.md](개발문서/로그.md)에 사실대로 기록함.
- Excel 기능은 기준 XLSM의 사용자 흐름만 벤치마킹하여 클린룸으로 재구현함. 외부 링크, 연결, `#REF!` 이름 정의는 새 생성본에서 각각 0건이어야 함.
- HWPX 생성본은 Kordoc 구조검증, 재열기, PDF 비교, 한컴 수동 열기를 분리하여 기록함. `validate` 성공만으로 완료 처리하지 않음.
- 골든 테스트에는 정상값, 필수값 누락, 최대 길이·특수문자, 복수 반복행을 포함함. 출력 전 사실관계·계약조건·서식 상태는 행정실 담당자가 최종 확인함.
- 배포, 외부 전송, 강제 푸시, 대량 삭제, 시크릿 변경은 반드시 사용자 확인 후 수행함.

## 실행 및 검증

- Excel 안전 빌드(구조 생성·VBA 주입·사본 검증·통과본 저장): `powershell -ExecutionPolicy Bypass -File scripts/run_build_excel_v1_structure_utf8.ps1`. 통과본 파일명은 `교복구매_길라잡이_YYYYMMDD_vN.xlsm`임.
- `build_excel_v1_structure.ps1`, `build_excel_v1_vba.ps1`, `verify_excel_v1.ps1`은 임시 `artifacts/excel/build.xlsm`을 대상으로 하는 보조 단계이며, 배포 후보를 직접 갱신하지 않음.
- 활성 Phase 품질 게이트: `powershell -ExecutionPolicy Bypass -File .claude/hooks/verify-phase.ps1`
- Markdown 공백 검사: `git diff --check`

품질 게이트 설정과 Hook은 [.claude/quality-gate.json](.claude/quality-gate.json), [.claude/settings.json](.claude/settings.json), [.claude/hooks/verify-phase.ps1](.claude/hooks/verify-phase.ps1)에 있음. 활성 Phase의 검사·작업·체크리스트가 모두 통과할 때만 완료 처리함.

## 역할과 참고 문서

- `uniform-workflow-analyst`: 업무 절차·서식 인벤토리·입력 모델 분석
- `document-automation-engineer`: HWPX 템플릿화·치환·조판 검증
- `excel-automation-engineer`: XLSM 워크플로우·자동반영·선택 출력 구현
- `quality-compliance-reviewer`: 개인정보·원본 무결성·출력·회귀 독립 검토

역할별 워크플로우는 [.claude/skills/uniform-purchase-orchestrator/SKILL.md](.claude/skills/uniform-purchase-orchestrator/SKILL.md), 목표 기반 재개·검증은 [.claude/skills/uniform-purchase-goal/SKILL.md](.claude/skills/uniform-purchase-goal/SKILL.md)를 따름.

세부 기준은 중복 작성하지 않고 다음 문서를 참조함: [서식_매핑표.md](개발문서/서식_매핑표.md), [결정사항.md](개발문서/결정사항.md), [엑셀_벤치마크_분석.md](개발문서/엑셀_벤치마크_분석.md), [kordoc_기술검증.md](개발문서/kordoc_기술검증.md).
