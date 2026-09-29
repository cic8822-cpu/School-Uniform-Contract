# 배포 안내 (WP-03-13)

## 확인된 사실

- `npm run build`로 생성한 `dist/`는 정적 파일(HTML/CSS/JS/JSON/HWPX 템플릿)만으로 구성되며, 서버·DB를 요구하지 않는다.
- `dist/index.html`을 `file://` 경로로 **직접 여는 것은 동작하지 않는다.** Vite 설정(`base: './'`, 상대경로)과 무관하게, 브라우저가 `<script type="module">`을 `file://` 출처에서 CORS 정책으로 차단하기 때문이다(크롬 개발자도구로 실제 확인함 — `Access to script ... blocked by CORS policy`, origin이 `null`로 취급됨). 이는 Vite/이 프로젝트만의 문제가 아니라 모든 최신 브라우저의 ES 모듈 보안 제약이다.
- 반면 로컬 정적 서버로 `dist/`를 서빙하면 정상 동작한다 — `npx serve dist`(포트 4173)로 띄운 뒤 접속했을 때 홈 화면(계약방법 4종 카드)이 콘솔 오류 없이 완전히 렌더링됨을 크롬 자동화로 확인했다.

## 학교 PC 배포 방법

1. `webapp/dist/` 폴더 전체를 학교 PC의 원하는 위치(예: `C:\교복길라잡이\`)로 복사한다.
2. 다음 중 하나로 그 폴더를 정적 서빙한다(인터넷 연결 불필요, localhost로만 접속):
   - Node.js가 설치된 PC: `npx serve -l 5000 .` (해당 폴더에서 실행) 후 브라우저로 `http://localhost:5000` 접속
   - Python이 설치된 PC: `python -m http.server 5000` 후 `http://localhost:5000` 접속
   - IIS가 있는 Windows 서버/PC: `dist/` 내용을 정적 사이트 루트로 지정
3. 브라우저 즐겨찾기에 위 주소를 등록해두면 매번 서버를 다시 켠 뒤 같은 주소로 접속하면 된다.

## 릴리스 파일 이름 규칙 (이름_날짜_버전)

바뀐 파일을 구분할 수 있도록 릴리스 파일 이름에 날짜(`YYYYMMDD`)와 버전을 넣는다.

| 대상 | 형식 | 예 |
|---|---|---|
| 릴리스 zip (GitHub Assets) | `UniformContractGuide_날짜_v버전.zip` | `UniformContractGuide_20260930_v1.1.zip` |
| zip 안의 exe | `교복계약길라잡이_날짜_v버전.exe` | `교복계약길라잡이_20260930_v1.1.exe` |

- GitHub는 한글 파일명을 `.`으로 바꿔 저장하는 경우가 있어 zip은 영문, 그 안의 exe는 한글로 짓는다.
- 릴리스를 올릴 때마다 날짜·버전을 올리고 태그(`v1.1.0`)도 함께 새로 만든다. 이전 태그의 파일은 그대로 둔다.

## exe 만드는 방법 (실측 검증됨)

`package.json`의 `npm run build:exe`는 그대로 실행하면 안 된다(패키저가 로컬에 설치돼 있지 않고 `pkg` 설정 파일도 지정하지 않는다). 아래 순서로 만든다.

```bash
cd webapp
npm run build
npx -y @yao-pkg/pkg deploy/server.cjs --config package.json --targets node24-win-x64 --output deploy/dist-exe/교복계약길라잡이_20260930_v1.1.exe
```

- **`--config package.json`을 빼면 웹앱 파일(`dist`)이 exe에 들어가지 않아** 실행하자마자 "dist 폴더를 찾을 수 없습니다"라며 종료된다(2026-09-30 실측).
- 올리기 전에 exe를 `PORT=9123`으로 실행해 `/`, `/assets/*`, `/templates/*`(48개), `/data/*.json`이 모두 200인지 확인한다.

## 남은 일 (다음 세션)

- 위 3가지 방법 중 학교 현장에서 실제로 어떤 방식을 쓸지(관리자 권한 필요 여부, 자동 시작 설정 등)는 사용자·현장 담당자와 확인이 필요하다.
- 서버를 PC 부팅 시 자동 실행하게 만드는 설치 스크립트(예: 작업 스케줄러 등록)는 아직 만들지 않았다.
