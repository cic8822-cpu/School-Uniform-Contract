# 웹앱 작업목록 — 교복 학교주관구매 길라잡이 웹

- **최신 상태(2026-09-25)**: `WP-01 DONE(설계 확정, 사용자 제공 스펙 기준) · WP-02 DONE(마스터 데이터·독립 재검토 통과) · WP-03 12/13 DONE, 1건(WP-03-13) TODO`. 상세 배경은 `웹앱_로그.md` 2026-09-25 항목(codex CLI 중단분 이어받아 완료) 참고.

## WP-01 설계 확정

| ID | 작업 | 선행 | 상태 | 완료 증빙 |
|---|---|---|---|---|
| WP-01-01 | 요구사항 정의 | - | DONE | `웹앱_prd.md`(WFR-01~WFR-14) |
| WP-01-02 | 아키텍처 결정(서버 유무, HWPX 생성 위치, 첫 화면 IA) | WP-01-01 | DONE | `ADR-002`(2026-09-25 개정판) |
| WP-01-03 | 화면 설계(7개 IA, 디자인 시스템) | WP-01-01 | DONE | 사용자 제공: `0-2.md`, 디자인 PDF 6장 |

## WP-02 마스터 데이터 구조화 (코딩 전 필수 선행)

| ID | 작업 | 선행 | 상태 | 완료 증빙 |
|---|---|---|---|---|
| WP-02-01 | 계약방법 Master 작성(`contract-methods.json`) | WP-01 | DONE | 4종·자동판정 금지·Workflow 연결 검증 통과 |
| WP-02-02 | Workflow Master 작성(`workflows.json`, 단계별 Form ID 포함) | WP-02-01 | DONE | 4/5/11단계·근거상태·확인 필요 표시 계약 검증 통과 |
| WP-02-03 | Form Master 작성(`forms.json`, 52종) | WP-01 | DONE | 52종·계산 Field 분리·HWPX 토큰 Field 계약 검증 통과 |
| WP-02-04 | Field Master 작성(`fields.json`) | WP-01 | DONE | 33개 Field·보호 정책·F-050 D-05 회귀 검증 통과 |
| WP-02-05 | Template Mapping 작성 | WP-02-03, WP-02-04 | DONE | 48개 검증 템플릿·허용 토큰·토큰↔Field 전수 대조 통과 |

## WP-03 웹앱 구현

| ID | 작업 | 선행 | 상태 | 완료 증빙 |
|---|---|---|---|---|
| WP-03-01 | 프로젝트 스캐폴딩(`웹앱_trd.md` 기술스택) | WP-02 전체 | DONE | 이전 세션(codex CLI)이 Vite+React19+TS로 스캐폴딩. 이번 세션에서 `App.tsx` 압축 1줄 코드를 `웹앱_trd.md` §5 구조(`src/pages`~`components/`·`lib/`·`data/`)로 리팩터링 |
| WP-03-02 | 홈(계약방법 4종 카드) | WP-03-01 | DONE | `webapp/src/components/home/HomePage.tsx`. 크롬 자동화로 4종 카드 실렌더링 확인 |
| WP-03-03 | 계약방법 안내(4열 비교) | WP-03-01 | DONE | `webapp/src/components/contractGuide/ContractGuidePage.tsx`(신규). 크롬 자동화로 4열 비교표 실렌더링 확인 |
| WP-03-04 | 계약절차(Stepper+단계상세) | WP-02-02, WP-03-01 | DONE | `webapp/src/components/procedure/ProcedurePage.tsx`. 크롬 자동화로 1인견적 4단계 Stepper·STEP01 상세 확인 |
| WP-03-05 | 기초자료입력(7탭, Excel 업로드/다운로드 포함) | WP-02-04, WP-03-01 | DONE | `webapp/src/components/input/InputPage.tsx`(공통/사업/문서별/품목/업체/위원/평가 7탭). fields.json 33개 Field + 반복그룹 4종 전부 렌더링, K-01/K-02 자동계산을 크롬 자동화로 실측 확인(수량10×단가50000=500,000원). `workbook.ts`를 반복그룹 지원하도록 확장해 업로드/다운로드 버튼에 연결 |
| WP-03-06 | 서식함(검색·필터·카드 52종) | WP-02-03, WP-03-01 | DONE | `webapp/src/components/forms/FormsPage.tsx`. 개인정보 보호 서식 배지 추가 |
| WP-03-07 | 문서 미리보기 모달(A4) | WP-03-05, WP-03-06 | DONE | `webapp/src/components/modal/DocumentPreviewModal.tsx`(신규). 크롬 자동화로 F-008 모달 실렌더링(라벨·누락값·버튼) 확인 |
| WP-03-08 | PDF 클라이언트 생성 | WP-03-07 | DONE | `webapp/src/lib/pdf.ts`(신규, html2canvas+jsPDF 동적 import). 모달 "PDF 출력" 버튼에 연결 |
| WP-03-09 | HWPX 클라이언트 생성(`tokenFill.ts`/`footerFill.ts` 포팅) | WP-03-07 | DONE | `webapp/src/lib/hwpx.ts`(허용/금지 토큰 목록 원본과 문자 단위 대조 확인, 유지) + `webapp/src/lib/hwpxFooterFill.ts`(신규, `hwpx_school_footer_fill.ps1`의 `Set-NameNearLabel` 포팅, 담당자·교장만) + `webapp/src/lib/hwpxExport.ts`(신규 오케스트레이션). 크롬 브라우저에서 실제 템플릿(F-008, F-007)에 토큰 치환·라벨 채움을 실행해 결과 zip을 재검증(치환된 학교명 포함, 미치환 토큰 0건, 담당자·교장 라벨 채움 확인) |
| WP-03-10 | 내보내기(업무자료.xlsx, 일괄 ZIP) | WP-03-05, WP-03-08, WP-03-09 | DONE | `webapp/src/components/exportPage/ExportPage.tsx`(신규) + `webapp/src/lib/hwpxBulkExport.ts`(신규, 완성된 서식 전체를 ZIP 하나로) |
| WP-03-11 | 사용안내(온보딩·FAQ·용어집) | WP-03-01 | DONE | `webapp/src/components/guide/GuidePage.tsx`(이용 순서 5단계, FAQ 5문항, 용어집 6항목로 확장) |
| WP-03-12 | 접근성·반응형·오프라인 QA | WP-03 전체 | IN_PROGRESS | 기본 접근성만 적용: 스킵 링크(`webapp/src/styles/accessibility.css`), `:focus-visible` 표시, 모달 `role="dialog"`+`aria-modal`+`aria-labelledby`+Esc 닫기+포커스 이동, 시맨틱 랜드마크(`nav aria-label`), 새 컴포넌트 CSS에 `@media (max-width: 600px/900px)` 반응형 규칙 포함. **미수행**: WCAG 대비 실측(자동 도구), 320~1920px 전수 스크린샷, 오프라인(서비스워커/캐싱) 동작 검증 — `웹앱_test.md` WA-01~04에 NOT_RUN으로 남김 |
| WP-03-13 | 학교 PC 로컬 설치 패키지 배포 | WP-03-12 | IN_PROGRESS | `vite.config.ts`에 `base: './'` 적용 + `useAppData.ts`/`hwpxExport.ts`의 자산 경로를 `assetUrl()`(상대경로) 기준으로 수정. 크롬 자동화로 실측: `file://`로 `dist/index.html` 직접 열기는 **불가능**(Vite 설정과 무관하게 모든 브라우저가 `file://` 출처의 `<script type="module">`을 CORS로 차단 — 원인 확정), 반면 `npx serve dist`로 로컬 정적 서버를 띄우면 홈 화면이 콘솔 오류 0건으로 정상 렌더링됨을 확인. `webapp/DEPLOY.md`(신규)에 검증된 배포 방법(로컬 정적 서버 3가지 옵션) 문서화. **남음**: 학교 현장에서 실제 사용할 서버 방식 확정, PC 부팅 시 자동 실행(작업 스케줄러 등) 설치 스크립트는 미작성 |

## 차단 사유

- WP-03-13만 남음(TODO) — 배포 인프라 방식(로컬 정적 서버 선택 등) 결정이 필요해 사용자 확인 후 진행함. 다른 항목은 차단 없음.
