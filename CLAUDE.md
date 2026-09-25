# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

교복 학교주관구매 업무를 안내하고 Excel·HWPX·PDF 서식 출력을 자동화하는 프로젝트임. `AGENTS.md`에도 같은 취지의 지침이 있으니 두 파일을 함께 갱신함(중복 서술은 피하고 최신 상태만 반영).

## 프로젝트 정보

- **우선 목표**: 원본 HWPX 서식을 근거로 Excel XLSM에서 기초자료를 한 번 입력해 서식별 반영·선택 출력·PDF 출력을 구현함. Excel에서 HWPX 출력이 곤란하면 대표 POC와 ADR 검증 후 웹앱에서 HWPX 생성·다운로드를 구현함.
- **주요 자료**: 교복 매뉴얼 HWPX, 기존 용역계약 길라잡이 XLSM, [계획.md](개발문서/계획.md)
- **지원 환경**: Windows, Microsoft 365, 한컴오피스
- **원본 보호**: 제공된 HWPX·XLSM은 절대 덮어쓰지 않음. 사본·생성본은 `artifacts/`와 `_workspace/`에 둠.
- **데이터 경계(2026-09-23 개정)**: 사용자 결정으로 업체명·담당자 성명(신설)·교장 성명(결재권자)·위원 실명(기존 마스킹 폐지)·업체 대표자 성명·연락처(텍스트)를 기초자료입력에서 한 번 입력하면 Excel·HWPX 전 서식에 자동 반영하도록 개인정보 자동화 경계를 확대함(상세 근거는 [입력데이터_사전.md](개발문서/입력데이터_사전.md) 20행과 `개발문서/로그.md` 2026-09-23 항목). **학생·학부모 개인정보(F-029·047·050)는 계속 MVP 제외 범위** — 애초에 기초자료입력에 학생·학부모 개별 데이터를 담는 반복 입력 구조 자체가 없고, 이 세 서식은 빈 양식으로 인쇄해 손으로 작성·제출받는 정상적인 업무 방식이므로 그대로 유지함(2026-09-23 재확인). **직인(도장) 이미지 자동 날인과 실제 필적 서명(그래픽), 신체 치수는 여전히 제외**함 — 성명·날짜 텍스트 자동반영과는 별개이며 요청받지 않았음. 테스트·골든 검증은 계속 마스킹 값만 사용하고, 실제 개인정보가 채워진 산출물은 git에 커밋하지 않음. 문서 출력에 필요한 업무·계약 정보는 중앙 서버에 저장하지 않으며 로그·Git·대화에 실제 개인정보 원문을 노출하지 않음.

## 자주 사용하는 명령

- **Excel 안전 빌드(현재 활성 배포 경로)**: `powershell -ExecutionPolicy Bypass -File scripts/run_build_excel_v1_structure_utf8.ps1` — 내부적으로 구조 생성(`scripts/build_excel_v1_structure.ps1`) → VBA 주입(`scripts/build_excel_v1_vba.ps1`) → 사본 검증(`scripts/verify_excel_v1.ps1`)을 순서대로 UTF-8로 재해석해 실행함. Excel COM 자동화이므로 Microsoft 365가 설치된 Windows에서만 동작함. 검증을 통과하면 `artifacts/excel/교복구매_길라잡이_YYYYMMDD_vN.xlsm`으로 저장하며, 같은 날짜에 다시 통과할 때마다 `vN`을 증가시켜 기존 통과본을 보존함(기존 파일을 덮어쓰지 않음).
  - `build_excel_v1_structure.ps1`·`build_excel_v1_vba.ps1`·`verify_excel_v1.ps1`을 직접 실행하면 임시 `artifacts/excel/build.xlsm`을 대상으로 한 개별 단계만 수행하며, 배포 후보(`v N.xlsm`)를 갱신하지 않음.
  - `scripts/create_excel_mvp.ps1`/`run_excel_mvp_utf8.ps1`은 초기 8시트 구조 초안용 구버전 스크립트로, P3-01 진행 중 "구조 초안 수준"이라는 이유로 완료 상태가 철회됨([task.md](개발문서/task.md) 차단 사유 참고). 더 이상 활성 경로가 아니므로 새 작업에는 사용하지 않음.
- **HWPX 명시적 토큰 치환(P2-01, Excel VBA와 미연계된 독립 CLI)**: `scripts/hwpx_token_fill.ps1 -TemplatePath <hwpx> -JsonPath <BOM 없는 UTF-8 JSON> -OutputPath <출력 hwpx>`로 `{{학교명}}` 등 허용 토큰만 원본을 훼손하지 않고(XML 재직렬화 없이 `hp:t` 텍스트만 치환) 채움. `scripts/export_hwpx_from_excel.ps1 -ExcelPath <xlsm> -TokenTemplatePath <hwpx> -OutputPath <출력>`은 기초자료입력 시트의 허용 셀(C4·C5·C8·C9·C21·C22)만 읽어 같은 경로로 연결함. 회귀 검증은 `scripts/verify_hwpx_token_fill.ps1`·`scripts/verify_excel_hwpx_export.ps1`. 허용 토큰 목록(교장명 포함 여부 등)은 개인정보 경계와 직결되므로 임의로 확장하지 않고 [입력데이터_사전.md](개발문서/입력데이터_사전.md) C-07 및 해당 스크립트의 `$AllowedTokens`/`$ForbiddenTokenPattern`을 함께 확인함.
- **Phase 검증 Hook 수동 실행**: `powershell -ExecutionPolicy Bypass -File .claude/hooks/verify-phase.ps1` (Write/Edit·TaskCompleted·Stop 시 `.claude/settings.json` Hook으로 자동 실행되며, `.claude/quality-gate.json`의 `activePhase`·`phaseTestIds`·`sourceHashes` 기준으로 체크리스트·테스트 상태·원본 해시를 검사해 미완료 시 완료 처리를 차단함.)
- **원본 무결성 확인**: 원본 `.xlsm`·`.hwpx`의 SHA-256을 계산해 `.claude/quality-gate.json`의 `sourceHashes`와 대조함.
- **Markdown 공백 오류 검사(Q-02 게이트)**: `git diff --check`
- **단일 테스트 선택 실행 불가**: `verify_excel_v1.ps1`은 [test.md](개발문서/test.md)의 골든 시나리오(E-01~E-06)와 회귀 검사를 하나의 Excel COM 세션·검증 사본 안에서 순서대로 모두 실행하는 단일 스크립트이며, 개별 시나리오만 골라 실행하는 옵션이나 파라미터는 없음. 특정 시나리오(예: E-03) 결과만 확인하려면 전체 스크립트를 실행한 뒤 [test.md](개발문서/test.md)에서 해당 ID의 PASS/PARTIAL/NOT_RUN 행을 확인함.

## 아키텍처 개요

- **Phase 파이프라인**: P0(원본 무결성·Kordoc/XLSM 정적분석) → P1(Form ID 인벤토리·데이터 모델·POC 범위 확정) → P2(HWPX POC) → P3(Excel MVP) → P4(현장 검증) → P5(웹앱). 각 Phase의 세부 작업 ID·담당·선행·상태는 [task.md](개발문서/task.md)에, 전체 로드맵과 근거는 [계획.md](개발문서/계획.md)에 있음. 현재 활성 Phase는 `.claude/quality-gate.json`의 `activePhase` 값이 기준임(자주 바뀌므로 이 파일에 상태를 하드코딩하지 않음 — 항상 [task.md](개발문서/task.md)·[test.md](개발문서/test.md)·[로그.md](개발문서/로그.md)를 직접 확인함).
- **에이전트 역할 분리**(`.claude/agents/`): `uniform-workflow-analyst`(업무·서식 분석 → `개발문서/서식_인벤토리.md`/`개발문서/입력데이터_사전.md`/`개발문서/서식_매핑표.md` 작성) → `document-automation-engineer`(Kordoc 기반 HWPX 템플릿화·치환·검증) / `excel-automation-engineer`(XLSM 클린룸 재구현, 외부링크·매크로 원본 이식 금지) → `quality-compliance-reviewer`(독립 검토, Critical/High 결함 시 반려하며 완료 처리하지 않음). 각 담당자의 작업기록은 `_workspace/0N_*/`, 최종 산출물은 `artifacts/`에 분리해 둠.
- **스킬 오케스트레이션**(`.claude/skills/`): `uniform-purchase-orchestrator`가 분석(병렬)→문서자동화/Excel구현→검증(파이프라인) 순서로 에이전트 실행을 조율함. `uniform-purchase-goal`은 `/goal` 기반 목표 실행·상태 확인·재개·해제를 담당함.
- **품질 게이트**(`.claude/quality-gate.json` + `.claude/hooks/verify-phase.ps1`): Write/Edit(`static` 모드)·TaskCompleted·Stop(`full` 모드) 시점마다 활성 Phase의 [체크리스트.md](개발문서/체크리스트.md)·[test.md](개발문서/test.md)·원본 SHA-256을 검사함. FAIL·PENDING·미해결 BLOCKED·체크리스트 미완료가 있으면 완료를 차단함(Q-05). P3-01에서는 `static` 모드에서도 `_workspace/03_excel/verify_v1_log.txt`의 PASS 기록이 `build_excel_v1_structure.ps1`/`build_excel_v1_vba.ps1`/`verify_excel_v1.ps1`보다 최신인지 검사함 — 세 스크립트 중 하나라도 로그보다 나중에 수정되면 안전 빌드를 재실행하지 않은 것으로 보고 완료를 차단함(2026-09-19 code-reviewer 검토로 추가; `개발문서/test.md`를 수동으로 PASS 표기해도 코드 변경 후 재검증 없이는 통과하지 못하게 막는 회귀 방지 장치임). 미완료 사유가 직전 실행과 완전히 동일하면(`.claude/.verify-phase-stop-state.json`에 시그니처 저장) 두 번째부터는 정보성 안내만 남기고 차단하지 않음 — 실패 사유가 바뀌면(또는 통과 후 다시 실패하면) 다시 한 번 차단함(2026-09-19 사용자 요청으로 추가; 백그라운드 작업 대기 중 Stop 훅이 동일 경고를 무한 반복해 세션을 끝낼 수 없던 문제를 해결함).
- **핵심 문서 지도**: `개발문서/서식_인벤토리.md`(Form ID F-001~F-057 목록, 개인정보 가능 문서 F-029/F-047/F-050 포함) · `개발문서/입력데이터_사전.md`(공통/문서별/반복/계산/보호 필드 정의) · `개발문서/서식_매핑표.md`(F-001~F-052 출력 매핑, F-053~F-057은 읽기 전용) · `개발문서/결정사항.md`(범위·대표 POC 3종 확정 근거) · `개발문서/엑셀_벤치마크_분석.md`(기존 XLSM 정적분석) · `개발문서/kordoc_기술검증.md`(HWPX 파싱/치환 기술검증) · `개발문서/로그.md`(세션 재개 기록, `/clear` 전 필수 갱신).
- **Excel 생성 경로(v1, 활성)**: `scripts/build_excel_v1_structure.ps1`이 Excel COM으로 시트를 순서대로 생성함 — `사용설명서`·`기초자료입력`·`DB`(및 `DB_품목`/`DB_업체`/`DB_위원`/`DB_평가` 반복행 정규화 시트)·`서식선택_출력`(F-001~F-052 선택 목록)·Form ID별 서식 시트(`F-007_구매요청기안문`/`F-024_단가비율표`/`F-014_평가항목배점기준` 등, 계속 추가 중)·`학교정보`·`학교검색`·`절차안내`·`계약방법안내`. 이어서 `build_excel_v1_vba.ps1`이 매크로(신규/초기화/저장/수정/불러오기, 저장 직전 전체 DB 재검증 등)를 주입하고 `verify_excel_v1.ps1`이 외부 링크·연결·`#REF!` 이름 정의 0건과 골든 시나리오를 사본에서 검사함. 미구현 Form ID·현재 통과/실패 상세는 [test.md](개발문서/test.md)를 확인함(하드코딩하지 않음).
- **원본 자산(수정 금지)**: 루트의 `20230808_용역계약갈라잡이(디깅모멘텀)_이행원.xlsm`, `교복4차 (2)/*.hwpx`. 모든 파생 작업은 `_workspace/`와 `artifacts/`에서만 수행함.
- **HWPX 명시적 토큰 치환 경로(P2-01, Excel v1과 별개)**: `scripts/hwpx_token_fill.ps1`(HWPX 템플릿 + JSON → 토큰 치환본)과 `scripts/export_hwpx_from_excel.ps1`(XLSM 허용 셀 → 같은 치환 경로)이 원본을 재직렬화하지 않고 `hp:t` 텍스트만 바이트 단위로 바꾸는 방식으로 동작함. 아직 `build_excel_v1_vba.ps1`의 VBA와 연계되지 않은 독립 CLI이며(CR-02, 미해결), 한컴 수동 열기·PDF 비교(P2-02)도 TODO이므로 이 경로로 실제 서식을 배포하지 않음. 허용 토큰은 화이트리스트 방식이며, 2026-09-19 결정으로 결재권자(교장) 성명은 공문서 관행상 허용 목록에 포함했으나 위원·업체 대표자 등 그 외 개인 성명·서명·연락처는 계속 거부함 — 경계를 바꿀 때는 반드시 [입력데이터_사전.md](개발문서/입력데이터_사전.md) C-07을 함께 갱신함.
- **HWPX P2-04 일괄 생성 경로(Excel VBA와 실제 연계된 활성 경로, P2-01과는 별개 스크립트)**: `scripts/export_hwpx_p204_from_excel.ps1 -ExcelPath <xlsm> [-TemplateDir artifacts/hwpx/P2-04] [-OutputDir artifacts/hwpx/P2-04_생성] [-FormIds F-XXX,...]`가 `build_excel_v1_vba.ps1`의 `HWPX작성_실행`(HWPX 작성 버튼)이 실제로 호출하는 스크립트임. 기초자료입력의 6개 허용 셀(C4·C5·C8·C9·C18·C23)과 담당자·교장 성명(C10·F10, 2026-09-23 결정으로 허용)을 읽어 `TemplateDir`의 템플릿마다 필요한 `{{토큰}}`만 골라 `hwpx_token_fill.ps1`로 채운 뒤, `scripts/hwpx_school_footer_fill.ps1`로 하단 결재란의 학교주소·전화(학교정보!F/G열 조회)와 담당자·교장 성명까지 채움(2026-09-24 신규 구현 — 담당자·교장 성명 텍스트는 채우되 실제 필적 서명·직인 이미지는 여전히 건드리지 않음). **주의**: (1) `hwpx_school_footer_fill.ps1`에 UTF-8 BOM이 없으면 Windows PowerShell 5.1이 소스 안의 "주소"·"전화"·"담당자"·"교장" 한글 리터럴을 시스템 코드페이지로 잘못 읽어 대상 셀을 조용히 0개 찾는다(2026-09-24 실측, 수정 시 BOM 유지 필수). (2) 이 스크립트들이 원본 xlsm의 해시를 재검증할 때 `Get-FileHash`(FileShare.Read)를 그대로 쓰면, 사용자가 그 통합문서를 Excel에서 열어둔 채(FileAccess.ReadWrite로 잠김) 버튼을 눌렀을 때 "다른 프로세스에서 사용 중" 오류로 항상 실패한다 — 반드시 `FileShare.ReadWrite`로 직접 스트림을 여는 `Get-SharedFileHash` 패턴을 씀(2026-09-24 근본 원인 수정, `export_hwpx_p204_from_excel.ps1`·`export_hwpx_from_excel.ps1` 동일 적용).

## 하네스: 교복구매 길라잡이

- **실행 모드**: 분석은 병렬, 문서 자동화→Excel 구현→검증은 파이프라인으로 진행함.
- **오케스트레이터(Excel/HWPX, P0~P4)**: 교복구매 길라잡이의 분석·구현·검증·재개 요청에는 `.claude/skills/uniform-purchase-orchestrator/SKILL.md`를 사용함.
- **오케스트레이터(웹앱, P5)**: 웹앱 설계·디자인·구현·재개 요청에는 `.claude/skills/webapp-purchase-orchestrator/SKILL.md`를 사용함. Excel/HWPX 트랙과 독립적으로 운영하며 `개발문서/웹앱_prd.md`·`개발문서/웹앱_계획.md`·`개발문서/웹앱_trd.md`·`개발문서/웹앱_task.md`·`개발문서/웹앱_test.md`·`개발문서/웹앱_체크리스트.md`·`개발문서/웹앱_로드맵.md`·`개발문서/웹앱_로그.md`·`ADR-001`·`ADR-002`를 근거로 함. 이 문서들의 최상위 근거는 사용자가 2026-09-25 직접 제공한 `0-1.교복계약 웹앱개발-프롬프트.txt`·`개발문서/0-2.교복 학교주관구매 웹앱 — 화면 구성 및 디자인 계획.md`·`첫화면\교복 학교주관구매 웹앱 화면 디자인.pdf`임(업무·서식 원문은 교복 매뉴얼, 데이터 구조는 Excel 자동화 파일을 따르는 이원 원칙). 사용자 요청("천천히 진행")에 따라 WP-01→WP-02(마스터 데이터 구조화)→WP-03(구현) 각 단계 전환 전 사용자 확인을 거침.
- **Goal 운영**: `/goal`, 목표 기반 자동 실행, 상태 확인·재개·해제 요청에는 `.claude/skills/uniform-purchase-goal/SKILL.md`를 사용함.
- **에이전트(Excel/HWPX)**: `uniform-workflow-analyst`, `document-automation-engineer`, `excel-automation-engineer`, `quality-compliance-reviewer` 역할을 사용함.
- **에이전트(웹앱)**: `web-ux-designer`, `web-frontend-engineer`를 사용하며, `quality-compliance-reviewer`는 웹앱 검토에도 공용으로 사용함.
- **상태 기준**: [계획.md](개발문서/계획.md)(Excel/HWPX)와 [웹앱_계획.md](개발문서/웹앱_계획.md)(웹앱)의 가장 높은 우선순위 `TODO` 또는 `IN_PROGRESS` 작업부터 수행하며, 품질 검토에서 Critical/High 결함이 남으면 완료로 보고하지 않음.

## 작업 원칙

- 항상 한글로 답변함.
- 작업 전 실제 대상 파일과 프로젝트 지침을 읽고, 요청 범위·가정·검증 가능한 완료 조건을 구분함. 결론을 바꾸는 불확실성만 한 가지 질문으로 확인함.
- 구현 전 사실·가정·선택지를 구분함. 불확실한 법령·계약 문구·도구 동작은 원문 또는 실제 실행으로 확인함.
- 한 작업에는 사용자 요청을 충족하는 최소 변경만 적용함. 요청하지 않은 리팩터링·추상화·기능 확장은 하지 않음.
- 기존 스타일과 공식 서식의 구조를 유지함. 원본보다 보기 좋다는 이유만으로 공문·계약서의 필수 구조를 변경하지 않음.
- 변경한 모든 줄은 사용자 요청과 완료조건에 연결되어야 함. 변경으로 생긴 미사용 요소만 제거함.
- 여러 단계 작업은 간단한 계획, 단계별 검증, 결과 기록의 순서를 지킴.
- 실패·미확정·차단 요인은 숨기지 않고 원인, 시도 결과, 가능한 다음 조치를 기록함.
- 보고서와 시각 자료에서는 핵심 메시지의 우선순위, 확인된 수치·근거, 남은 병목·최종 확인 항목을 구분하여 제시함. 막연한 형용사로 근거를 대체하지 않음.

## 검증과 완료 기준

- 수정 전 대상 파일을 실제로 읽음.
- Excel 산출물은 외부 링크·외부 연결·`#REF!` 이름 정의가 없는지 확인함.
- HWPX 산출물은 Kordoc 구조검증, 재열기, PDF 비교 및 한컴 수동 확인 항목을 분리해 기록함.
- 골든 테스트는 정상값, 필수값 누락, 최대 길이/특수문자, 복수 반복행을 포함함.
- 생성 문서는 행정실 직원이 사실관계·계약조건·서식 상태를 최종 확인한 뒤 출력함.
- 배포, 외부 전송, 강제 푸시, 대량 삭제, 시크릿 수정은 반드시 사용자 확인 후 진행함.

## /goal 사용 원칙

- `/goal`은 Claude Code의 세션 범위 내장 기능이며 하나의 측정 가능한 완료 조건에만 사용함.
- 목표 문장에는 완료조건, 검사 방법, 원본 보호와 같은 제약을 함께 적음.
- `/goal` 평가자는 대화에 나타난 정보로만 판정하므로 명령·테스트·검증 결과를 작업 기록에 남김.
- `/goal`은 권한을 바꾸지 않음. 자동 모드를 쓰더라도 배포·외부 전송 등 파급 작업은 별도 확인이 필요함.

## 변경 이력

| 날짜 | 변경 | 대상 | 사유 |
|---|---|---|---|
| 2026-09-15 | 교복구매 길라잡이 하네스 초기 구성 | `.claude/agents`, `.claude/skills`, `CLAUDE.md` | 서식 분석, HWPX 자동화, XLSM 구현, 품질 검증 역할을 분리하고 /goal 운영 기준을 등록함 |
| 2026-09-16 | `/init` 재점검: 표준 헤더 추가, Excel 명령·생성 경로를 v1 활성 스크립트(`build_excel_v1_structure.ps1` 등)로 갱신, Phase/기능 상태 하드코딩 제거 | `CLAUDE.md` | `create_excel_mvp.ps1` 기반 구버전 명령이 실제 활성 경로와 달라져 있었고, 상태 서술이 `개발문서/task.md`/`개발문서/test.md`와 어긋날 위험이 있어 최신화함 |
| 2026-09-17 | `/init` 재점검: 내용은 실제 저장소 상태(`개발문서/task.md`/`개발문서/test.md`/`quality-gate.json`, 스크립트 경로, 에이전트·스킬 구성)와 일치함을 확인. `verify_excel_v1.ps1`이 골든 시나리오를 개별 선택 실행할 수 없는 단일 스크립트라는 점만 보완 | `CLAUDE.md` | 향후 인스턴스가 "단일 테스트 실행" 방법을 찾다가 존재하지 않는 옵션을 가정하지 않도록 명시함 |
| 2026-09-18 | `/init` 재점검: `CLAUDE.md`는 실제 스크립트·quality-gate 구성과 일치해 변경 없음. `AGENTS.md`의 "구현 Form ID: F-007, F-024만" 서술이 `개발문서/test.md`(F-007·F-014~F-020·F-024 구현·검증 완료, 2026-09-17 기준)와 어긋나 있어 하드코딩을 제거하고 `개발문서/test.md` 참조로 교체함 | `AGENTS.md` | Form ID 구현 상태가 자주 바뀌어 두 안내 파일 중 하나에만 최신화가 반영되면 어긋남이 재발하므로, `CLAUDE.md`가 이미 채택한 "상태 하드코딩 금지" 원칙을 `AGENTS.md`에도 동일 적용함 |
| 2026-09-19 | `/init` 재점검: `CLAUDE.md`는 변경 없음(activePhase·스크립트·에이전트·스킬 구성이 실제 저장소와 일치). `AGENTS.md`의 "`P2-01`은 Kordoc `fill -j` 치환이 실패해 차단 상태" 서술이 같은 날짜 `개발문서/로그.md`(F-007·F-024·F-013 라벨 기반 재구축으로 `fill -j → validate → render → 재파싱` 왕복 통과, P2-01 기술 검증 범위 `DONE` 전환, P2-02는 `TODO`)와 어긋나 있어 최신 상태로 교체함 | `AGENTS.md` | `P2-01` 차단 원인이라는 구체적 인과 서술이 하드코딩된 채 남아 있었고 해소 이후에도 갱신되지 않아, 향후 인스턴스가 이미 풀린 차단을 여전히 유효한 제약으로 오인할 위험이 있었음 |
| 2026-09-19 | `/init` 재점검: 최근 커밋(`5d4769f`)에서 추가된 `scripts/hwpx_token_fill.ps1`·`scripts/export_hwpx_from_excel.ps1`(HWPX 명시적 토큰 치환, P2-01 독립 CLI 경로)과 대응 회귀 검증 스크립트가 아키텍처 개요·명령 목록에 빠져 있어 추가함 | `CLAUDE.md` | 이 경로는 Excel v1 VBA와 아직 연계되지 않았고(CR-02 미해결) 교장명 허용/기타 개인 성명 거부라는 개인정보 경계 결정을 담고 있어, 문서화하지 않으면 향후 인스턴스가 이 CLI의 존재를 놓치거나 화이트리스트 경계를 임의로 넓힐 위험이 있었음 |
| 2026-09-19 | `/init` 재점검: 작업 트리의 미커밋 변경(`verify-phase.ps1`)에서 P3-01 품질 게이트가 `static` 모드에서도 `verify_v1_log.txt`와 Excel 빌드 스크립트 3종의 수정 시각을 비교해 재검증 여부를 강제하도록 바뀐 것을 확인해 품질 게이트 서술에 반영함 | `CLAUDE.md` | 이 동작이 문서화되지 않으면 향후 인스턴스가 스크립트만 고치고 안전 빌드 재실행 없이 완료 처리를 시도하다 Write/Edit 훅에서 이유를 모른 채 막히는 상황을 겪을 위험이 있었음 |
| 2026-09-19 | 사용자가 "동일알림 반복되지 않게 해"라고 명시적으로 요청해, `verify-phase.ps1`에 실패 시그니처 기반 세션 상태 파일(`.claude/.verify-phase-stop-state.json`)을 추가함. 같은 미완료 사유가 반복되면 두 번째부터는 차단(`exit 2`) 대신 정보성 안내만 남기고 통과(`exit 0`)시킴 | `.claude/hooks/verify-phase.ps1`, `.gitignore`, `CLAUDE.md` | P3-01처럼 여러 세션에 걸쳐 완료되는 Phase에서는 미완료 상태가 정상이고, 이를 Stop 훅이 세션이 끝날 때마다 완전히 동일한 내용으로 무한 재차단하면 백그라운드 작업 대기 중 세션을 정상 종료할 수 없는 문제가 있었음. 실패 사유가 실제로 바뀌면 여전히 다시 차단하므로 게이트의 회귀 방지 목적은 유지됨 |
| 2026-09-24 | HWPX 작성 버튼 종료 코드 1의 근본 원인(열려 있는 xlsm에 `Get-FileHash`가 `FileShare.Read`로 접근해 충돌)을 `Get-SharedFileHash`(`FileShare.ReadWrite`)로 수정, 실제 사용 중인 P2-04 HWPX 생성 경로(`export_hwpx_p204_from_excel.ps1`·`hwpx_school_footer_fill.ps1`)를 아키텍처 개요에 신규 문서화, 하단 결재란 학교주소·전화·담당자·교장 성명 자동반영을 새로 구현(주소·전화는 BOM 누락으로 0셀만 반영되던 버그 수정, 담당자·교장 성명은 로직 자체가 없어 라벨 기반 매칭으로 신규 구현), 계약방법안내 시트 버튼 3개를 하단에서 상단 배너 제목 옆으로 재배치, 계약절차 팝업 버튼 하단 여백 재확대(420→452), 지난 세션 미완료였던 "표 단일 입력 병합 전수조사"를 47개 서식 전체 재검토로 완료(F-007만 결함, 나머지는 이상 없음 확인) | `scripts/export_hwpx_p204_from_excel.ps1`, `scripts/export_hwpx_from_excel.ps1`, `scripts/hwpx_school_footer_fill.ps1`, `scripts/build_excel_v1_vba.ps1`, `CLAUDE.md`, `개발문서/로그.md` | 사용자가 HWPX 작성 버튼 오류 재현과 결재란 자동반영 미반영을 반복 보고해, 코드만 고치고 재검증하지 않았던 이전 회차의 실수를 피하기 위해 48개 서식 전체 재생성·kordoc 구조검증·SVG 렌더 육안 확인까지 끝낸 뒤에만 완료로 기록함 |
| 2026-09-25 | 웹앱(P5) 전용 하네스 신규 구성: 에이전트 2종(`web-ux-designer`, `web-frontend-engineer`) 신설, `quality-compliance-reviewer` 범위를 웹앱까지 확장, 신규 오케스트레이터 스킬 `webapp-purchase-orchestrator` 작성(WP-01→WP-02→WP-03 파이프라인, 단계 전환마다 사용자 확인), `uniform-purchase-orchestrator` description에서 "웹앱 전환" 문구를 제거해 트리거 중복 방지 | `.claude/agents/web-ux-designer.md`, `.claude/agents/web-frontend-engineer.md`, `.claude/agents/quality-compliance-reviewer.md`, `.claude/skills/webapp-purchase-orchestrator/SKILL.md`, `.claude/skills/uniform-purchase-orchestrator/SKILL.md`, `CLAUDE.md` | 사용자가 "harness 를 구성해줘 웹앱 개발관련 문서 참고해서"로 명시 요청. Excel/HWPX 트랙과 도메인이 달라(프론트엔드 구현 vs Excel COM/VBA) 별도 에이전트·오케스트레이터가 필요했고, 사용자가 "웹앱 디자인을 먼저 만들게 천천히 진행"을 요구해 오케스트레이터에 단계별 확인 체크포인트를 내장함 |
| 2026-09-25 | 웹앱 문서·하네스 전면 재정정: 사용자가 대화 밖 `개발문서/0-0. 최초 지시사항.txt`에 직접 추가한 새 지시(웹 HWPX 생성 필요, 계약방법 우선 IA)를 Stop hook 경고 조사 중 발견해 Plan Mode로 1차 정정하던 도중, 사용자가 더 권위 있는 상세 스펙 3종(`0-1.교복계약 웹앱개발-프롬프트.txt`, `개발문서/0-2.교복 학교주관구매 웹앱 — 화면 구성 및 디자인 계획.md`, 화면 디자인 PDF 6장)을 직접 제공함. 이를 최종 근거로 `ADR-002`·`개발문서/웹앱_prd.md`·`개발문서/웹앱_계획.md`·`개발문서/웹앱_trd.md`를 전면 재작성하고 `개발문서/웹앱_task.md`·`개발문서/웹앱_test.md`·`개발문서/웹앱_체크리스트.md`·`개발문서/웹앱_로드맵.md`·`개발문서/웹앱_로그.md`를 신규 작성함. 핵심 변경: 첫 화면을 9단계 절차 허브에서 계약방법 4종 카드로, HWPX를 "Excel 안내만"에서 "웹 브라우저 클라이언트 사이드 직접 생성"(서버 없이 `hwpx_token_fill.ps1` 로직을 JSZip+DOMParser로 포팅)으로 확정. 하네스 4개 파일(`web-ux-designer.md`·`web-frontend-engineer.md`·`quality-compliance-reviewer.md`·`webapp-purchase-orchestrator/SKILL.md`)도 이 결정에 맞춰 갱신함 | `개발문서/ADR-002-웹앱아키텍처.md`, `개발문서/웹앱_prd.md`, `개발문서/웹앱_계획.md`, `개발문서/웹앱_trd.md`, `개발문서/웹앱_task.md`, `개발문서/웹앱_test.md`, `개발문서/웹앱_체크리스트.md`, `개발문서/웹앱_로드맵.md`, `개발문서/웹앱_로그.md`, `.claude/agents/web-ux-designer.md`, `.claude/agents/web-frontend-engineer.md`, `.claude/agents/quality-compliance-reviewer.md`, `.claude/skills/webapp-purchase-orchestrator/SKILL.md`, `CLAUDE.md` | 사용자가 "웹앱제작용으로 필요한 문서를 작성해줘"(계획.md/prd.md/task.md/로드맵.md/test.md/체크리스트.md/로그.md)와 하네스 재구성을 명시 요청. 이전 결정을 조용히 덮어쓰지 않고 왜 뒤집혔는지 기록하는 것이 이 프로젝트의 확립된 관행이라 상세 경위를 남김(`개발문서/웹앱_로그.md` 2026-09-25 항목 참고) |
