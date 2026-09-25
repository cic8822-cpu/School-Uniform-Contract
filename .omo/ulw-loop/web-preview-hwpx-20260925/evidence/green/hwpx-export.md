# GREEN — F-008 생성 HWPX 무결성

- 시나리오: 비식별 학교명으로 브라우저에서 생성한 F-008 HWPX를 원본 템플릿과 비교했다.
- 실행:

  `Add-Type -AssemblyName System.IO.Compression.FileSystem` 후 각 ZIP의 비본문 파일을 SHA-256으로 비교하고, 생성본 `Contents/section*.xml`의 미치환 토큰 수를 검사했다.

- 이진 관찰값:

  - 생성 파일: `C:\Users\최익창\Downloads\F-008_교복 사양서.hwpx`
  - 파일 크기: 2,090,977 bytes
  - 비본문 파일 SHA-256 일치: `true` (57 files)
  - `{{...}}` 미치환 토큰: `0`
  - 비식별 학교명 포함: `true`

- 판정: PASS. 기존 허용 토큰 치환·다운로드 흐름이 실제 생성본에서 완료됐고, 본문 외 파일은 보존됐다.
