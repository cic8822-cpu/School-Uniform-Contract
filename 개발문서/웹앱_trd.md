# 웹앱 TRD — 교복 학교주관구매 길라잡이 웹

- **상태**: `DRAFT — 전면 개정(2026-09-25)`
- **근거**: `웹앱_prd.md`, `웹앱_계획.md`, `ADR-002`, `0-1.txt` §20~23, `0-2.md`, 디자인 PDF, `입력데이터_사전.md`, `서식_매핑표.md`

## 1. 기술 스택

| 영역 | 선택 | 근거 |
|---|---|---|
| 아키텍처 | Static-first — 가능하면 순수 HTML/CSS/JS, 필요 시 React/Vite로 빌드하되 **최종 산출물은 정적 파일** | `0-1.txt` §20 명시 요구사항 |
| 프레임워크 | React + Vite(빌드 도구), 정적 산출물로 export | 52개 Form ID 카드·7개 IA 화면을 컴포넌트로 재사용, 빌드 후 런타임 서버 없음 |
| 스타일 | CSS(디자인 시스템 토큰 직접 구현) — Tailwind 등 유틸리티 프레임워크는 선택사항 | `0-2.md` §8 디자인 시스템을 그대로 CSS 변수로 구현 |
| 상태 관리 | 컴포넌트 로컬 상태 + 브라우저 세션 한정(새로고침 시 소실 가능, 명시적 저장은 Excel 다운로드로만) | 학교 PC 로컬 설치·서버·DB 미사용 원칙과 일치 |
| ZIP/XML 처리 | `JSZip`(HWPX·업무자료 zip 처리), `DOMParser`/`XMLSerializer`(HWPX section XML 조작), 클라이언트 xlsx 파서(Excel Import/Export) | 2026-09-25 기술조사로 서버 없이도 충분함을 확인 |
| 배포 | 학교 PC 로컬 설치 패키지 | 인터넷 연결 없이 정적 파일만으로 동작, GitHub Pages·Vercel 등 외부 호스팅 제외 |

## 2. 데이터 구조 (`0-1.txt` §21 — 고정)

```
/data
  contract-methods.json   # 계약방법 Master (4종: 적용조건/단계수/공통·전용서식)
  workflows.json          # Workflow Master (계약방법별 Step 배열: 단계명/업무설명/관련서식/담당구분)
  forms.json              # Form Master (52종: System ID/매뉴얼번호/서식명/업무단계/작성주체/사용계약방법/필요Field/출력방식)
  fields.json             # Field Master (Field ID/한글명/Type/Required/사용서식/Source)
/templates                # 원본 HWPX 서식 템플릿(원본 바이너리 보존, `{{필드}}` 자리만 치환 대상)
                           # → 이미 `artifacts/hwpx/P2-04/`에 48개 Form ID의 `F-XXX_..._템플릿.hwpx`가 검증 완료 상태로 존재함(P2-04, 2026-09-23~24, 48/48 성공).
                           #   새로 토큰화할 필요 없이 그대로 복사해 사용함. 각 템플릿과 짝을 이루는 `_골든결과.hwpx`(마스킹 데이터 치환 예시)는
                           #   웹 HWPX 엔진(tokenFill.ts)의 회귀 테스트 기준본으로 재사용함.
/assets                   # 로고·아이콘·일러스트(전북교육청 브랜드 자산)
/js
  data-store              # Master 데이터 로딩·Canonical Field 상태 관리
  form-engine              # 서식별 필수값 검증·완성도 판정(● / ▲ / ○)
  excel-engine              # 업무자료.xlsx 파싱(업로드)·생성(다운로드)
  document-engine             # A4 미리보기 렌더링, PDF 생성
  workflow-engine               # 계약방법→Stepper→단계상세 라우팅
  hwpx-engine                    # HWPX 클라이언트 생성(아래 §4)
```

## 3. 데이터 모델 상세

### 3.1 계약방법 Master(4종) — Excel `계약방법안내`·`frmContractGuide` 이식(2026-09-25 Explore 조사 결과)

| 계약방법 | 단계 수 | 공통 필수 서식 | 테마 색상(HEX) |
|---|---|---|---|
| 1인견적 수의계약 | 4 | F-008, F-012 | Teal `#14B8A6` |
| 2인견적 수의계약 | 5 | F-008, F-012 | Coral `#F0653E` |
| 2단계 계약(규격·가격 동시) | 11 | F-008~F-012 + 제안·평가 서식 일체 | Purple `#6D5BD0` |
| 일반입찰계약 | 11 | 2단계계약과 동일 서식 체계 | Blue `#2F6FED` |

11단계 공통(2단계·일반입찰): 입찰공고문작성→지정정보처리장치공고→입찰(제안서)제출→규격(적격)심사→가격개찰→낙찰자결정→계약체결→계약이행→검사·검수→대가지급→변경계약. 각 단계의 관련 Form ID는 Excel `frmContractGuide.초기화`의 `stepFormIDs` 배열 값을 그대로 이식한다(예: STEP04 규격(적격)심사 = F-014·F-015·F-016·F-032~F-040).

### 3.2 Field Master(요약) — `입력데이터_사전.md` 인용

| 그룹 | 필드 ID 예 | Canonical Field 예 |
|---|---|---|
| 공통(C-01~C-08) | 학교명, 학년도, 담당부서, 담당자, 교장성명, 문서발행일, 문서번호 | `{{SCHOOL_NAME}}`, `{{SCHOOL_YEAR}}` |
| 사업(B-01~B-07) | 구매명, 계약방식, 기초금액, 예정수량, 납품기한, 제출기한 | `{{PURCHASE_NAME}}`, `{{BASIC_PRICE}}`, `{{DELIVERY_DATE}}` |
| 문서별(D-01~D-05) | 수신, 제목, 관련문서, 붙임목록, 안내문본문 | — |
| 품목 반복(R-04~R-06, K-01) | 품목명, 수량, 단가, 금액(자동계산) | — |
| 업체 반복(R-03, R-08, R-09) | 업체명, 대표자성명, 대표자연락처 | `{{COMPANY_NAME}}`, `{{COMPANY_CEO}}` |
| 위원 반복(R-01, R-02) | 역할·직위, 위원성명 | — |
| 평가 반복(R-07, K-03) | 평가항목, 배점, 점수, 총점(자동계산) | — |

개인정보 보호 문서(F-029·F-047·F-050)의 금지 필드는 입력 UI 자체를 만들지 않는다(기존 Excel 경계와 동일).

### 3.3 Form Master

전체 52종을 System ID/매뉴얼번호(`[9-7]` 등)/서식명/업무단계/작성주체(학교작성/업체작성/학교·업체공동/참고자료)/사용 계약방법/필요 Field/출력방식(PDF/HWPX/둘다) 스키마로 관리한다. 원본 `서식_매핑표.md`·`서식_인벤토리.md`를 소스로 변환하며, 상태값은 3종(● 작성가능 / ▲ 기초자료 필요 / ○ 미구현)에 더해 F-013처럼 표·이미지 복합조판인 경우 "HWPX 우선(복합조판)" 상태를 별도로 둔다.

## 4. HWPX 클라이언트 생성 엔진 (신규 — 서버 없이 구현)

`ADR-002` §4 결정에 따라 다음 두 PowerShell 스크립트의 로직을 TypeScript로 포팅한다(2026-09-25 기술조사로 이식 가능성 확인 완료):

| 원본(PowerShell) | 포팅 대상(TypeScript) | 핵심 로직 |
|---|---|---|
| `scripts/hwpx_token_fill.ps1` | `js/hwpx-engine/tokenFill.ts` | HWPX(zip)를 `JSZip`으로 열기 → `Contents/section*.xml`을 문자열로 읽기 → `hp:t` 텍스트 런 사이에 걸친 `{{토큰}}`도 인식하는 regex 치환 → 허용 토큰 목록(`$AllowedTokens`)·금지 패턴(`$ForbiddenTokenPattern`) 검사 → 영향받은 문단의 `hp:linesegarray` 캐시만 제거 → `DOMParser`로 재파싱해 XML 유효성 확인 → `JSZip`으로 재압축 → `Blob` 다운로드 |
| `scripts/hwpx_school_footer_fill.ps1` | `js/hwpx-engine/footerFill.ts` | `DOMParser`+`XMLSerializer`로 `hp:tbl`/`hp:tc` 표 순회 → `name` 속성 또는 라벨 텍스트("담당자"/"교장" 등)로 셀 탐색 → 값 채움 |

**정확성 요구사항(Critical)**: `$AllowedTokens`/`$ForbiddenTokenPattern`을 TS로 옮길 때 원본과 완전히 동일해야 한다. 어긋나면 개인정보 경계가 깨진다 — `quality-compliance-reviewer`가 이 일치 여부를 코드 리뷰에서 직접 대조한다.

**구조 검증**: 원본은 `npx kordoc validate`(Node CLI, 브라우저 실행 불가)를 QA 게이트로 쓰지만, 웹에서는 재압축 직후 재파싱해 `{{...}}` 패턴이 남아있지 않은지 자체 검사하는 것으로 대체한다(원본 스크립트도 결국 같은 수준의 자체 검증을 이미 하고 있었음).

## 5. 컴포넌트 구조 (제안 — 7개 IA에 대응)

```
pages/
  Home.tsx                    # 계약방법 4종 카드
  ContractMethodGuide.tsx     # 계약방법 비교(4열 표)
  ContractProcedure.tsx       # Stepper + 단계상세
  DataEntry.tsx                # 기초자료입력(7탭)
  FormLibrary.tsx               # 서식함(검색·필터·카드)
  Export.tsx                     # 내보내기
  Guide.tsx                       # 사용안내(온보딩·FAQ)
components/
  ContractMethodCard.tsx      # 홈 카드(테마색·단계수·필요서식수 배지)
  WorkflowStepper.tsx         # 좌측 수직 Stepper(완료✓/현재/예정)
  StepDetailCard.tsx          # 우측 단계 상세(해야할일/주의사항/서식표)
  FormCard.tsx                 # 서식함 카드(상태배지 포함)
  DocumentPreviewModal.tsx     # A4 미리보기 + 필수값/누락값 패널 + 액션바
  DataEntryTabs.tsx             # 기초자료 7탭 폼
  ExcelUpload.tsx / ExcelExport.tsx
  PrivacyBanner.tsx             # 상시 노출 Privacy 안내
lib/
  hwpx/tokenFill.ts, hwpx/footerFill.ts   # §4
  pdf/generate.ts                          # 클라이언트 PDF 생성
  validation/formReadiness.ts               # ●/▲/○ 상태 판정
data/  # §2의 JSON 4종
```

## 6. 접근성 (WFR-13)

- WCAG 2.1 AA(텍스트 대비 4.5:1 이상), 시맨틱 HTML, 키보드 전용 조작 가능.
- 디자인 시스템의 색상 대비를 직접 검증(`0-2.md` §8 팔레트 기준).

## 7. 로컬 운영·개인정보 경계

- 기초자료 입력값·엑셀 파싱 결과·HWPX 생성 과정 전부 네트워크 요청에 포함되지 않음(devtools 검증 대상). 결과물과 명시적 Excel 백업은 설치된 해당 학교 PC에서만 유지함.
- Privacy 안내 문구를 기초자료입력·내보내기 화면에 고정 노출(`0-1.txt` §13 인용 문구 그대로 사용).

## 8. 미결정(WP-03에서 확정)

1. PDF 생성 라이브러리(브라우저 인쇄 CSS vs jsPDF류) 프로토타입 비교.
2. 학교 PC 로컬 설치 패키지의 실행 경로와 업데이트 절차를 WP-03에서 검증.
3. React+Vite 채택 여부 대 순수 HTML/CSS/JS(성능·유지보수 트레이드오프는 WP-02 완료 후 재검토).
