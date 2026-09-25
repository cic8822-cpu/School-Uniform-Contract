# P3-01 클린룸 재구현 — 검증: 외부 링크/연결/#REF! 0건, 매크로 실행, PDF 출력, A4 1쪽 재확인
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$defaultTargetPath = Join-Path $root 'artifacts\excel\build.xlsm'
$targetPath = if ([string]::IsNullOrWhiteSpace($env:UNIFORM_EXCEL_BUILD_PATH)) { $defaultTargetPath } else { $env:UNIFORM_EXCEL_BUILD_PATH }
$logPath = Join-Path $root '_workspace\03_excel\verify_v1_log.txt'
$validationDirectory = Join-Path $root ('artifacts\excel\_validation_' + $PID)
$validationPath = Join-Path $validationDirectory '교복구매_길라잡이_Excel_v1.xlsm'
$legacyValidationPath = Join-Path $root 'artifacts\excel\~검증용_교복구매_길라잡이_Excel_v1.xlsm'
$failures = [System.Collections.Generic.List[string]]::new()

if (-not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
    throw "검증 대상 XLSM이 없습니다: $targetPath"
}

# Phase 훅은 모든 Write/Edit마다 이 스크립트를 다시 실행하므로, 직전 실행이 정리한
# EXCEL.EXE의 파일 잠금이 완전히 풀리기 전에 다음 실행이 곧바로 이어질 수 있음.
# 잠금 해제를 잠깐 기다리는 재시도로 이 경합 때문에 최상단에서 통째로 실패하는 것을 막는다.
function RemovePathWithRetry([string]$path, [int]$maxAttempts = 5, [int]$delayMs = 500) {
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop }
            return
        } catch {
            if ($attempt -eq $maxAttempts) { throw }
            Start-Sleep -Milliseconds $delayMs
        }
    }
}

RemovePathWithRetry -path $logPath
RemovePathWithRetry -path $legacyValidationPath
RemovePathWithRetry -path $validationDirectory
New-Item -ItemType Directory -Path $validationDirectory -Force | Out-Null
function L($s) {
    [System.IO.File]::AppendAllText($logPath, "$s`r`n", [System.Text.UTF8Encoding]::new($false))
}
function Assert-Check([bool]$condition, [string]$message) {
    if ($condition) {
        L "PASS: $message"
    } else {
        L "FAIL: $message"
        $script:failures.Add($message)
    }
}
function Invoke-ValidationMacro([string]$macroName) {
    $excel.Run($macroName)
}

Copy-Item -LiteralPath $targetPath -Destination $validationPath -Force
# Quit()·ReleaseComObject·GC만으로는 남은 스크립트 지역변수가 COM 참조를 계속
# 살려 두어 EXCEL.EXE가 좀비로 남을 수 있음. 생성 전후 프로세스 목록을 비교해
# 새로 뜬 PID를 기록해 두고, finally에서 종료가 확인되지 않으면 이 PID만
# 강제 종료해 고아 프로세스가 다음 실행을 막지 않게 함.
$excelPidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel = New-Object -ComObject Excel.Application
Start-Sleep -Milliseconds 300
$excelProcessId = Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $excelPidsBefore -notcontains $_.Id } |
    Select-Object -First 1 -ExpandProperty Id
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null
try {
    $wb = $excel.Workbooks.Open($validationPath, [Type]::Missing, $false)

    # 1. 외부 링크
    $links = $wb.LinkSources(1)  # xlExcelLinks
    if ($null -eq $links) {
        L "외부 링크(xlExcelLinks): 0건"
    } else {
        L "외부 링크(xlExcelLinks): $($links.Count)건 -- $($links -join '; ')"
    }
    Assert-Check ($null -eq $links -or $links.Count -eq 0) '외부 링크가 0건임'

    # 2. 외부 연결(Connections)
    $connCount = $wb.Connections.Count
    L "외부 연결(Connections): $connCount 건"
    Assert-Check ($connCount -eq 0) '외부 연결이 0건임'

    # 3. #REF! 이름정의
    $refCount = 0
    $refNames = @()
    foreach ($n in $wb.Names) {
        try {
            $r = $n.RefersTo
            if ($r -like "*#REF!*") { $refCount++; $refNames += $n.Name }
        } catch { $refCount++; $refNames += $n.Name }
    }
    L "#REF! 이름정의: $refCount 건 $(if($refNames.Count -gt 0){'-- ' + ($refNames -join ', ')})"
    Assert-Check ($refCount -eq 0) '#REF! 이름 정의가 0건임'

    # 4. 시트 목록
    L "시트 목록: $((@($wb.Worksheets) | ForEach-Object { $_.Name }) -join ', ')"
    $requiredSheets = @('사용설명서', '기초자료입력', 'DB', 'DB_품목', 'DB_업체', 'DB_위원', 'DB_평가', 'DB_정량평가', 'DB_정성평가', 'DB_자기평점', '서식선택_출력', 'F-001_위원회구성기안', 'F-002_위원수락확인서', 'F-003_위원청렴보안서약서', 'F-004_구매추진계획수립기안문', 'F-005_구매추진계획안', 'F-006_운영위원회심의안', 'F-007_구매요청기안문', 'F-014_평가항목배점기준', 'F-015_정량적평가', 'F-016_정성적평가', 'F-017_제출서류자기확인서', 'F-018_정량적평가자기평점표', 'F-019_입찰참가신청서', 'F-020_입찰참가신고서', 'F-021_교복납품제안서', 'F-022_교복납품실적표', 'F-023_교복제조사양서', 'F-024_단가비율표', 'F-025_교복AS계획서', 'F-032_제안서접수결과', 'F-033_제안서접수대장', 'F-034_평가위원회개최', 'F-035_정량평가결과', 'F-036_제안서평가결과', 'F-037_평가위원회참석등록부', 'F-038_업체참가등록부', 'F-039_업체별제안서평가표', 'F-040_위원청렴보안서약서', 'F-041_낙찰자결정', 'F-042_낙찰자결정통보', 'F-043_계약체결', 'F-044_사전안내가정통신문안내', '학교정보', '학교검색', '입찰계약절차안내', '계약방법안내')
    $requiredSheets += 'F-045_사전안내가정통신문'
    $requiredSheets += 'F-046_수요조사가정통신문안내'
    $requiredSheets += 'F-048_교복구매안내신청수량파악'
    $requiredSheets += 'F-047_수요조사가정통신문'
    $requiredSheets += 'F-049_만족도설문조사실시'
    $requiredSheets += 'F-050_만족도조사설문지'
    $requiredSheets += 'F-009_기초금액계약방법결정'
    $requiredSheets += 'F-010_사전규격공개기안문'
    $requiredSheets += 'F-011_입찰공고'
    $requiredSheets += 'F-012_계약특수조건'
    $requiredSheets += 'F-008_교복사양서'
    $requiredSheets += 'F-027_위임장'
    $requiredSheets += 'F-029_개인정보제공동의서'
    $requiredSheets += 'F-030_청렴계약이행서약서'
    $requiredSheets += 'F-031_교복품목별금액표'
    $sheetNames = @($wb.Worksheets | ForEach-Object { $_.Name })
    Assert-Check ($sheetNames.Count -eq 62 -and @($requiredSheets | Where-Object { $_ -notin $sheetNames }).Count -eq 0) 'F-027·F-029·F-030·F-031을 포함한 필수 62개 시트가 모두 존재함'
    Assert-Check ($wb.Worksheets.Item(1).Name -eq '사용설명서' -and $wb.Worksheets.Item(2).Name -eq '기초자료입력' -and $wb.Worksheets.Item(3).Name -eq '계약방법안내') '시트 탭 순서가 1.사용설명서 2.기초자료입력 3.계약방법안내로 배치됨'
    $thisWorkbookComponent = @($wb.VBProject.VBComponents | Where-Object { $_.Type -eq 100 -and $_.Name -notmatch '^Sheet\d+$' }) | Select-Object -First 1
    $thisWorkbookCode = $thisWorkbookComponent.CodeModule.Lines(1, $thisWorkbookComponent.CodeModule.CountOfLines)
    $beforeSaveCount = [regex]::Matches($thisWorkbookCode, '(?im)^\s*(?:Private|Public)\s+Sub\s+Workbook_BeforeSave\b').Count
    $openCount = [regex]::Matches($thisWorkbookCode, '(?im)^\s*(?:Private|Public)\s+Sub\s+Workbook_Open\b').Count
    Assert-Check ($beforeSaveCount -eq 1 -and $openCount -eq 1) 'ThisWorkbook 이벤트 Workbook_Open·Workbook_BeforeSave가 각각 1개만 존재해 컴파일 중복 오류가 없음'
    $wsProtectedDB = $wb.Worksheets.Item('DB')
    $wsProtectedItems = $wb.Worksheets.Item('DB_품목')
    $wsProtectedVendors = $wb.Worksheets.Item('DB_업체')
    $wsProtectedCommittee = $wb.Worksheets.Item('DB_위원')
    $wsProtectedScore = $wb.Worksheets.Item('DB_평가')
    $wsProtectedQuant = $wb.Worksheets.Item('DB_정량평가')
    $wsProtectedQual = $wb.Worksheets.Item('DB_정성평가')
    $wsProtectedSelf = $wb.Worksheets.Item('DB_자기평점')
    Assert-Check ($wsProtectedDB.ProtectContents -and $wsProtectedItems.ProtectContents -and $wsProtectedVendors.ProtectContents -and $wsProtectedCommittee.ProtectContents -and $wsProtectedScore.ProtectContents -and $wsProtectedQuant.ProtectContents -and $wsProtectedQual.ProtectContents -and $wsProtectedSelf.ProtectContents -and $wsProtectedDB.Visible -eq 0 -and $wsProtectedItems.Visible -eq 2 -and $wsProtectedVendors.Visible -eq 2 -and $wsProtectedCommittee.Visible -eq 2 -and $wsProtectedScore.Visible -eq 2 -and $wsProtectedQuant.Visible -eq 2 -and $wsProtectedQual.Visible -eq 2 -and $wsProtectedSelf.Visible -eq 2) 'DB는 숨김·보호되고 반복 DB는 VeryHidden·보호됨'
    L 'DB 숨김·무암호 보호는 우발적 편집 방지 상태로만 확인함. 직접 DB 편집의 영구 저장 차단은 아래 Workbook_BeforeSave 저장 경계 시나리오에서 별도로 검증함.'
    # ProtectionMode=True여야 매크로가 UserInterfaceOnly로 쓸 수 있음(=False면 매크로도 막혀 저장 매크로가
    # 보호된 시트 오류로 멈춤). UserInterfaceOnly는 파일을 다시 열면 유지되지 않아 Workbook_Open이 재적용하며,
    # 새 DB 시트를 추가할 때 그 재적용 목록에 빠뜨리면 이 검사가 잡아낸다(F-015 DB_정량평가 무한 대기 재발 방지).
    Assert-Check ($wsProtectedDB.ProtectionMode -and $wsProtectedItems.ProtectionMode -and $wsProtectedVendors.ProtectionMode -and $wsProtectedCommittee.ProtectionMode -and $wsProtectedScore.ProtectionMode -and $wsProtectedQuant.ProtectionMode -and $wsProtectedQual.ProtectionMode -and $wsProtectedSelf.ProtectionMode) 'Workbook_Open이 재열기 후 모든 DB 시트의 UserInterfaceOnly 보호를 재적용함'

    # 4a. 학교검색 및 절차안내 정적 구조
    $wsSchool = $wb.Worksheets.Item('학교정보')
    Assert-Check ($wsSchool.Range('A793').Value2 -eq '학교정보 행 수' -and $wsSchool.Range('B793').Text -eq '790') '학교정보 공개 기관정보가 전북지역 전체 790개교임(2026-09-24부터 표본이 아닌 전체)'
    $wsSearch = $wb.Worksheets.Item('학교검색')
    Assert-Check ($wsSearch.Range('A1').Value2 -match '학교정보 검색' -and $wsSearch.Range('B3,D3,F3').Interior.Color -eq 16777164) '학교명·지역·급별 검색 조건 UI가 존재함'
    $wsFlow = $wb.Worksheets.Item('입찰계약절차안내')
    Assert-Check ($wsFlow.Range('A5:A13').Count -eq 9 -and $wsFlow.Range('D5').Value2 -eq 'F-001' -and $wsFlow.Range('D13').Value2 -eq 'F-044') '교복구매 9단계와 대표 Form ID 이동값이 존재함'
    $wsMethod = $wb.Worksheets.Item('계약방법안내')
    # 2026-09-23 버그 2 조사 중 발견(요청 4·버그 2와는 무관한 기존 결함): 이 검증은 2026-09-16
    # "예/아니오" 확인값 설계 당시 작성된 것으로, B6도 사용자가 채우는 확인값이라 가정해 빈 칸을
    # 요구했다. 그러나 2026-09-21 카드 버튼 UX 개편 이후 B6은 항상 고정 안내문
    # "버튼을 누르면 즉시 반영"을 담는 정적 라벨로 바뀌었고 build_excel_v1_structure.ps1·
    # build_excel_v1_vba.ps1 어디에도 B6을 지우는 코드가 없어, 개편 이후 이 검증은 구조적으로
    # 항상 FAIL할 수밖에 없었다(이번 버그 2 안전 빌드에서 처음 드러남). 실제 사용자 입력 대상인
    # B5(드롭다운 선택)·B7(선택값 echo 수식)만 빈 칸 여부를 검사하도록 정정한다.
    # 2026-09-24 사용자 요청으로 "계약방법 반영" 버튼(btn계약방법반영)을 삭제해(카드 즉시반영과
    # 기능 중복) 그 존재를 검증하던 조건도 제거하고, 대신 남은 이동 버튼 2개의 존재를 검사한다.
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsMethod.Range('B5').Value2) -and $wsMethod.Range('B6').Value2 -eq '버튼을 누르면 즉시 반영' -and [string]::IsNullOrWhiteSpace([string]$wsMethod.Range('B7').Value2) -and $wsMethod.Buttons('btn이동_기초자료입력').OnAction -eq '이동_기초자료입력' -and $wsMethod.Buttons('btn이동_서식선택출력').OnAction -eq '이동_서식선택출력') '계약방법안내가 빈 확인값(B5·B7)으로 시작하고 B6 고정 안내문·이동 버튼 2개가 존재함'

    # 5. A4 1쪽 자연 충족 재확인 (F-007, F-014, F-024)
    foreach ($sn in @("F-001_위원회구성기안","F-002_위원수락확인서","F-003_위원청렴보안서약서","F-004_구매추진계획수립기안문","F-006_운영위원회심의안","F-007_구매요청기안문","F-014_평가항목배점기준","F-015_정량적평가","F-016_정성적평가","F-017_제출서류자기확인서","F-018_정량적평가자기평점표","F-019_입찰참가신청서","F-020_입찰참가신고서","F-021_교복납품제안서","F-022_교복납품실적표","F-023_교복제조사양서","F-024_단가비율표","F-025_교복AS계획서","F-032_제안서접수결과","F-033_제안서접수대장","F-034_평가위원회개최","F-035_정량평가결과","F-036_제안서평가결과","F-039_업체별제안서평가표","F-041_낙찰자결정","F-042_낙찰자결정통보","F-043_계약체결","F-044_사전안내가정통신문안내","F-045_사전안내가정통신문","F-046_수요조사가정통신문안내","F-049_만족도설문조사실시")) {
        $ws = $wb.Worksheets.Item($sn)
        # HPageBreaks.Count가 워크북을 막 열었을 때 실제 값(0)이 아니라 오래된 캐시 값을
        # 반환하는 경우가 있음을 발견함(2026-09-21, 기초자료입력에 셰이프를 대량 추가한 뒤
        # 재현). 행 Height를 한 번씩 읽기만 해도(수정 없이) Excel이 페이지 레이아웃을 다시
        # 계산해 정확한 값으로 갱신됨 — 이 사이드이펙트 없는 방법으로 매 시트마다 갱신시킴.
        $usedRowCount = $ws.UsedRange.Rows.Count
        for ($touchRow = 1; $touchRow -le $usedRowCount; $touchRow++) { $null = $ws.Rows.Item($touchRow).Height }
        $hb = $ws.HPageBreaks.Count
        $vb = $ws.VPageBreaks.Count
        $zoom = $ws.PageSetup.Zoom
        L "$sn : Zoom=$zoom, HPageBreaks=$hb, VPageBreaks=$vb => $(if($hb -eq 0 -and $vb -eq 0){'A4 1쪽 자연 충족 OK'}else{'FAIL - 1쪽 초과'})"
        Assert-Check ($zoom -eq 100 -and $hb -eq 0 -and $vb -eq 0) "$sn A4 1쪽 자연 배율 출력"
    }

    # 5a. F-005는 원문이 6개 물리 페이지 분량이라 A4 1쪽 원칙의 예외로 문서화됨(F-005_구현검증.md).
    # 강제 축소(FitToPages)만 없으면 되고, 가로 폭 초과(VPageBreaks)만 결함으로 간주하며 세로 다중
    # 페이지(HPageBreaks>0)는 정상으로 허용한다.
    $wsF05chk = $wb.Worksheets.Item('F-005_구매추진계획안')
    $vbF05chk = $wsF05chk.VPageBreaks.Count
    $zoomF05chk = $wsF05chk.PageSetup.Zoom
    L "F-005_구매추진계획안 : Zoom=$zoomF05chk, HPageBreaks=$($wsF05chk.HPageBreaks.Count)(다중 페이지 허용, 예외), VPageBreaks=$vbF05chk"
    Assert-Check ($zoomF05chk -eq 100 -and $vbF05chk -eq 0 -and -not $wsF05chk.PageSetup.FitToPagesWide -and -not $wsF05chk.PageSetup.FitToPagesTall) 'F-005 자연 배율(강제 축소 없음)이며 가로 폭 초과가 없음(세로 다중 페이지는 원문이 6쪽이라 허용된 예외)'

    # 5b. F-048은 원문이 안내문+절취선+신청서+신청 수량 표로 구성된 2부 결합 문서라 A4 1쪽 원칙의
    # 예외로 문서화됨(F-048_구현검증.md). F-005와 같은 기준: 강제 축소만 없으면 되고, 가로 폭 초과
    # (VPageBreaks)만 결함으로 간주하며 세로 다중 페이지(HPageBreaks>0)는 정상으로 허용한다.
    $wsF48chk = $wb.Worksheets.Item('F-048_교복구매안내신청수량파악')
    $vbF48chk = $wsF48chk.VPageBreaks.Count
    $zoomF48chk = $wsF48chk.PageSetup.Zoom
    L "F-048_교복구매안내신청수량파악 : Zoom=$zoomF48chk, HPageBreaks=$($wsF48chk.HPageBreaks.Count)(다중 페이지 허용, 예외), VPageBreaks=$vbF48chk"
    Assert-Check ($zoomF48chk -eq 100 -and $vbF48chk -eq 0 -and -not $wsF48chk.PageSetup.FitToPagesWide -and -not $wsF48chk.PageSetup.FitToPagesTall) 'F-048 자연 배율(강제 축소 없음)이며 가로 폭 초과가 없음(세로 다중 페이지는 안내문+신청서+표 결합 구성이라 허용된 예외)'

    # 5c. F-047도 안내문+절취선+신청서 결합 구성이라 F-005·F-048과 같은 기준(VPageBreaks=0, 강제 축소
    # 없음)으로 확인하며, 세로 다중 페이지는 허용된 예외로 둔다.
    $wsF47chk = $wb.Worksheets.Item('F-047_수요조사가정통신문')
    $vbF47chk = $wsF47chk.VPageBreaks.Count
    $zoomF47chk = $wsF47chk.PageSetup.Zoom
    L "F-047_수요조사가정통신문 : Zoom=$zoomF47chk, HPageBreaks=$($wsF47chk.HPageBreaks.Count)(다중 페이지 허용, 예외), VPageBreaks=$vbF47chk"
    Assert-Check ($zoomF47chk -eq 100 -and $vbF47chk -eq 0 -and -not $wsF47chk.PageSetup.FitToPagesWide -and -not $wsF47chk.PageSetup.FitToPagesTall) 'F-047 자연 배율(강제 축소 없음)이며 가로 폭 초과가 없음(세로 다중 페이지는 안내문+신청서 결합 구성이라 허용된 예외)'

    # 6. 매크로 실행 테스트 (마스킹 테스트값 사용 — 실제 업체/개인정보 아님)
    $wsIn = $wb.Worksheets.Item("기초자료입력")
    $wsIn.Range("C4").Value2 = "테스트초등학교"
    $wsIn.Range("C5").Value2 = "2026"
    $wsIn.Range("C10").Value2 = "박운영"  # C-08 담당자 성명 (2026-09-23 신설, 마스킹 테스트값 — 다른 곳의 '홍길동' 기존 블랙리스트 검사와 충돌하지 않도록 구분된 값 사용)
    $wsIn.Range("F10").Value2 = "김교장"  # C-07 교장 성명 (2026-09-23 신설, 마스킹 테스트값)
    $wsIn.Range("C12").Value2 = "2026학년도 동복 구매"
    $wsIn.Range("C14").Value2 = "2단계 입찰(규격, 가격 동시입찰)"
    $wsIn.Range("C22").Value2 = "2026학년도 동복 학교주관구매 요청"
    $wsIn.Range("C21").Value2 = "교육지원청"
    $wsIn.Range("C24").Value2 = "구매계획서 1부"
    $wsIn.Range("B29").Value2 = "동복 상의"
    $wsIn.Range("C29").Value2 = 100
    $wsIn.Range("D29").Value2 = 50000
    $wsIn.Range("B30").Value2 = "동복 하의"
    $wsIn.Range("C30").Value2 = 100
    $wsIn.Range("D30").Value2 = 40000
    $wsIn.Range("C15").Value2 = 5000000
    $wsIn.Range("C16").Value2 = 100
    $wsIn.Range("C17").Formula = "=DATE(2026,3,1)"
    $wsIn.Range("C18").Formula = "=DATE(2026,2,20)+18/24"
    $wsIn.Range("B44").Value2 = "검증업체가"
    $wsIn.Range("S44").Value2 = "박대표"  # R-08 대표자 성명 (2026-09-23 신설, 마스킹 테스트값)
    $wsIn.Range("X44").Value2 = "02-000-0000"  # R-09 대표자 연락처 (2026-09-23 신설, 마스킹 테스트값)
    $wsIn.Range("C44").Value2 = 10
    $wsIn.Range("D44").Value2 = 10
    $wsIn.Range("E44").Value2 = 15
    $wsIn.Range("F44").Value2 = 15
    $wsIn.Range("B45").Value2 = "검증업체나"
    $wsIn.Range("C45").Value2 = 6
    $wsIn.Range("D45").Value2 = 8
    $wsIn.Range("E45").Value2 = 9
    $wsIn.Range("F45").Value2 = 0
    $wsIn.Range("N44").Value2 = 8
    $wsIn.Range("O44").Value2 = 8
    $wsIn.Range("P44").Value2 = 12
    $wsIn.Range("Q44").Value2 = 15
    $wsIn.Range("N45").Value2 = 10
    $wsIn.Range("O45").Value2 = 5
    $wsIn.Range("P45").Value2 = 5
    $wsIn.Range("Q45").Value2 = 0
    $wsIn.Range("H44").Value2 = '탁월'
    $wsIn.Range("I44").Value2 = '탁월'
    $wsIn.Range("J44").Value2 = '탁월'
    $wsIn.Range("K44").Value2 = '탁월'
    $wsIn.Range("L44").Value2 = 5
    $wsIn.Range("H45").Value2 = '보통'
    $wsIn.Range("I45").Value2 = '보통'
    $wsIn.Range("J45").Value2 = '우수'
    $wsIn.Range("K45").Value2 = '우수'
    $wsIn.Range("L45").Value2 = -2
    $wsIn.Range("B58").Value2 = "교원위원"
    $wsIn.Range("C58").Value2 = "위원 ○○"
    $wsIn.Range("B59").Value2 = "학부모위원"
    $wsIn.Range("C59").Value2 = "위원 **"
    $wsIn.Range("B72").Value2 = "제품 품질"
    $wsIn.Range("C72").Value2 = 40
    $wsIn.Range("D72").Value2 = 38.5
    $wsIn.Range("B73").Value2 = "사후 관리"
    $wsIn.Range("C73").Value2 = 60
    $wsIn.Range("D73").Value2 = 60
    $excel.CalculateFullRebuild()

    $wsF14 = $wb.Worksheets.Item('F-014_평가항목배점기준')
    Assert-Check ($wsF14.Range('B3').Value2 -eq '[붙임 3] 제안서 평가항목 및 배점기준' -and $wsF14.Range('B9').Value2 -eq '구분' -and $wsF14.Range('C9').Value2 -eq '평가항목' -and $wsF14.Range('D9').Value2 -eq '배점기준' -and $wsF14.Range('E9').Value2 -eq '배점' -and $wsF14.Range('F9').Value2 -eq '평가점수') 'F-014 원문 대조 제목과 평가기준 표 머리글이 존재함'
    Assert-Check ($wsF14.Range('C10').Value2 -eq '제품 품질' -and $wsF14.Range('E10').Value2 -eq 40 -and $wsF14.Range('F10').Value2 -eq 38.5 -and $wsF14.Range('C11').Value2 -eq '사후 관리' -and $wsF14.Range('F11').Value2 -eq 60 -and $wsF14.Range('F20').Value2 -eq 98.5) 'F-014가 R-07 복수 평가행과 K-03 총점을 자동 반영함'
    Assert-Check ([bool]$excel.Run('검증_F014출력가능')) 'F-014 선택 출력이 필수값과 완전한 복수 평가행을 허용함'
    Assert-Check ([string]::IsNullOrEmpty([string]$wsF14.Range('G20').Value2)) 'F-014 합계 행 비고(G20)에 내부 계산필드ID 문자열("K-03")이 인쇄되지 않고 빈 칸임(2026-09-23 버그 2 수정)'

    $wsF15 = $wb.Worksheets.Item('F-015_정량적평가')
    Assert-Check ($wsF15.Range('B3').Value2 -eq '[9-4] [붙임 3_1] [1단계] 정량적 평가' -and $wsF15.Range('B8').Value2 -eq '구분' -and $wsF15.Range('C8').Value2 -eq '평가항목' -and $wsF15.Range('D8').Value2 -eq '배점기준' -and $wsF15.Range('E8').Value2 -eq '배점' -and $wsF15.Range('F8').Value2 -eq '평가점수') 'F-015 원문 대조 제목과 고정 배점표 머리글이 존재함'
    $wsF15.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF15.Range('G5').Value2 -eq '검증업체가' -and $wsF15.Range('F9').Value2 -eq 10 -and $wsF15.Range('F13').Value2 -eq 10 -and $wsF15.Range('F16').Value2 -eq 15 -and $wsF15.Range('F20').Value2 -eq 15 -and $wsF15.Range('F22').Value2 -eq 50) 'F-015가 선택 업체 순번(I3)의 4개 항목 점수와 합계 50점을 정확히 반영함'
    $wsF15.Range('I3').Value2 = 2
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF15.Range('G5').Value2 -eq '검증업체나' -and $wsF15.Range('F9').Value2 -eq 6 -and $wsF15.Range('F22').Value2 -eq 23) 'F-015가 선택 업체 순번을 바꾸면 다른 업체의 점수를 반영함'
    $wsF15.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ([bool]$excel.Run('검증_F015출력가능')) 'F-015 선택 출력이 완전한 업체별 정량평가를 허용함'

    $wsF18 = $wb.Worksheets.Item('F-018_정량적평가자기평점표')
    Assert-Check ($wsF18.Range('B3').Value2 -eq '[9-7] [서식 1_1] 정량적 평가 자기 평점표' -and $wsF18.Range('F7').Value2 -eq '자기 평점') 'F-018 원문 대조 제목과 자기평점 머리글이 존재함'
    $wsF18.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF18.Range('C5').Value2 -eq '검증업체가' -and $wsF18.Range('F8').Value2 -eq 8 -and $wsF18.Range('F12').Value2 -eq 8 -and $wsF18.Range('F15').Value2 -eq 12 -and $wsF18.Range('F19').Value2 -eq 15 -and $wsF18.Range('F21').Value2 -eq 43 -and $wsF15.Range('F9').Value2 -eq 10) 'F-018이 F-015 학교평가(10점)와 다른 업체 자기평점(8점)만 반영하여 혼용하지 않음'
    $wsF18.Range('I3').Value2 = 2
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF18.Range('C5').Value2 -eq '검증업체나' -and $wsF18.Range('F21').Value2 -eq 20) 'F-018이 선택 업체별 자기평점 합계를 반영함'
    $wsF18.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ([bool]$excel.Run('검증_F018출력가능')) 'F-018 선택 출력이 완전한 업체별 자기평점을 허용함'

    $wsF19 = $wb.Worksheets.Item('F-019_입찰참가신청서')
    Assert-Check ($wsF19.Range('B3').Value2 -match '입 찰 참 가 신 청 서' -and $wsF19.Range('D6').Value2 -match '상호' -and $wsF19.Range('I6').Value2 -eq '법인등록번호') 'F-019 원문 대조 제목과 신청인 표 머리글이 존재함'
    $wsF19.Range('N3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF19.Range('F6').Value2 -eq '검증업체가' -and $wsF19.Range('F10').Value2 -eq '테스트초등학교 공고 제20○○-00호' -and $wsF19.Range('F11').Value2 -eq '테스트초등학교 2026학년도 교복 학교주관 구매 입찰') 'F-019가 선택 업체명과 학교 공통정보만 반영함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF19.Range('J6').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF19.Range('F7').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF19.Range('J7').Value2) -and $wsF19.Range('F8').Value2 -eq '박대표' -and [string]::IsNullOrWhiteSpace([string]$wsF19.Range('J8').Value2) -and $wsF19.Range('D15').Value2 -match '성 명' -and $wsF19.Range('I15').Value2 -match '사용인감') 'F-019 법인등록번호·연락처·대리인·인감 개인정보는 빈칸 또는 수기 안내로 보존하고, 대표자(R-08)는 2026-09-23부터 자동 반영됨'
    Assert-Check ([bool]$excel.Run('검증_F019출력가능')) 'F-019 선택 출력이 필수 학교정보와 업체명을 허용함'

    $wsF20 = $wb.Worksheets.Item('F-020_입찰참가신고서')
    Assert-Check ($wsF20.Range('B3').Value2 -eq '교복 학교주관구매 입찰 참가 신고서' -and $wsF20.Range('B7').Value2 -eq '사업자 상호' -and $wsF20.Range('B8').Value2 -eq '사업자 번호') 'F-020 원문 대조 제목과 사업자 정보 표 머리글이 존재함'
    $wsF20.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF20.Range('C7').Value2 -eq '검증업체가' -and $wsF20.Range('B11').Value2 -match '테스트초등학교' -and $wsF20.Range('B20').Value2 -eq '테스트초등학교장 귀하') 'F-020이 선택 업체명과 학교 공통정보만 반영함'
    # 2026-09-24 사용자 요청: 대표자 성명(R-08)·제출일자는 자동 반영으로 바뀌고, 사업자
    # 번호·직인만 계속 수기 보존함(F-019와 동일한 R-08 재사용 패턴).
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF20.Range('C8').Value2) -and $wsF20.Range('C9').Value2 -eq '박대표' -and $wsF20.Range('B19').Value2 -match '박대표' -and $wsF20.Range('B19').Value2 -match '사업자 직인') 'F-020 사업자 번호·직인은 빈칸/수기 보존, 대표자 성명은 2026-09-24부터 자동 반영됨'
    Assert-Check ([bool]$excel.Run('검증_F020출력가능')) 'F-020 선택 출력이 필수 학교정보와 업체명을 허용함'

    $wsF21 = $wb.Worksheets.Item('F-021_교복납품제안서')
    $wsF21.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF21.Range('B7').Value2 -eq '1. 일반현황 및 연혁' -and $wsF21.Range('B19').Value2 -eq '2. 제안(제출)서류' -and $wsF21.Range('C8').Value2 -eq '검증업체가' -and $wsF21.Range('B31').Value2 -eq '테스트초등학교장 귀하') 'F-021 원문 대조 표와 업체명·학교 공통정보 반영'
    Assert-Check ($wsF21.Range('C9').Value2 -eq '박대표' -and [string]::IsNullOrWhiteSpace([string]$wsF21.Range('C11').Value2) -and $wsF21.Range('B30').Value2 -match '서명 또는 날인' -and [bool]$excel.Run('검증_F021출력가능')) 'F-021이 R-08 대표자 성명을 반영하고(2026-09-23) 주소는 빈칸으로 보존하며 선택 출력을 허용함'

    $wsF22 = $wb.Worksheets.Item('F-022_교복납품실적표')
    $wsF22.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF22.Range('B3').Value2 -eq '교복 납품 실적표' -and $wsF22.Range('C9').Value2 -eq '사업명' -and $wsF22.Range('D9').Value2 -eq '사업기간' -and $wsF22.Range('E9').Value2 -match '계약금액' -and $wsF22.Range('F9').Value2 -eq '발주처' -and $wsF22.Range('B5').Value2 -match '검증업체가') 'F-022 원문 대조 제목·실적표 머리글과 선택 업체명이 존재함'
    Assert-Check ($wsF22.Range('B7').Value2 -match '최근 3년' -and $wsF22.Range('B10').Value2 -eq 1 -and $wsF22.Range('B18').Value2 -eq 9 -and $wsF22.Range('B20').Value2 -match '실적증명서' -and $wsF22.Range('B24').Value2 -match '서명 또는 날인' -and [bool]$excel.Run('검증_F022출력가능')) 'F-022 실적증명서 9행·수기 서명 경계 및 선택 출력 허용'

    $wsF23 = $wb.Worksheets.Item('F-023_교복제조사양서')
    $wsF23.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF23.Range('B3').Value2 -eq '납품 교복 제조 사양서' -and $wsF23.Range('B5').Value2 -match '검증업체가' -and $wsF23.Range('B7').Value2 -eq '생활형 교복' -and $wsF23.Range('B9').Value2 -eq '구분' -and $wsF23.Range('C9').Value2 -eq '견본 품목' -and $wsF23.Range('D9').Value2 -eq '재질' -and $wsF23.Range('E9').Value2 -eq '설명 내용') 'F-023 원문 대조 제목·표 머리글과 선택 업체명이 존재함'
    Assert-Check ($wsF23.Range('B10').Value2 -eq '동복' -and $wsF23.Range('C10').Value2 -eq '후드 점퍼' -and $wsF23.Range('C11').Value2 -eq '집업티' -and $wsF23.Range('C12').Value2 -eq '맨투맨티' -and $wsF23.Range('C13').Value2 -eq '긴바지' -and $wsF23.Range('B14').Value2 -eq '하복' -and $wsF23.Range('C14').Value2 -eq '반팔티' -and $wsF23.Range('C15').Value2 -eq '반바지') 'F-023 동복 4행·하복 2행 견본 품목이 원문과 일치함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF23.Range('D10').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF23.Range('E10').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF23.Range('D15').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF23.Range('E15').Value2) -and $wsF23.Range('B17').Value2 -match '국산섬유원단' -and $wsF23.Range('B20').Value2 -match '서명 또는 날인' -and [bool]$excel.Run('검증_F023출력가능')) 'F-023 재질·설명 내용 수기 경계 및 선택 출력 허용'

    $wsF25 = $wb.Worksheets.Item('F-025_교복AS계획서')
    $wsF25.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF25.Range('B3').Value2 -eq '교복 A/S 계획서' -and $wsF25.Range('B6').Value2 -eq '1. A/S 기간' -and $wsF25.Range('B9').Value2 -eq '2. A/S 편의성' -and $wsF25.Range('B15').Value2 -match 'A/S지정업체' -and $wsF25.Range('B20').Value2 -eq '4. A/S 내용') 'F-025 원문 대조 제목과 4개 항목(A/S 기간·편의성·지정업체·내용) 머리글이 존재함'
    Assert-Check ($wsF25.Range('B23').Value2 -match '바지/치마 기장수선' -and $wsF25.Range('B30').Value2 -match '서명 또는 날인' -and $wsF25.Range('B31').Value2 -eq '테스트초등학교장 귀하' -and [bool]$excel.Run('검증_F025출력가능')) 'F-025 원문에 업체명 표시 자리가 없어 학교 공통정보만 반영하고 선택 출력을 허용함'
    $bodyText025 = ($wsF25.Range('B1:E31').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyText025 -notmatch '검증업체가') 'F-025 본문 어디에도 업체명이 출력되지 않음(원문에 업체명 표시 자리 없음)'

    $wsF01 = $wb.Worksheets.Item('F-001_위원회구성기안')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF01.Range('B3').Value2 -eq '교복선정위원회 구성(안)' -and $wsF01.Range('B7').Value2 -eq '수  신' -and $wsF01.Range('C7').Value2 -eq '내부결재' -and $wsF01.Range('B9').Value2 -eq '제  목' -and $wsF01.Range('C9').Value2 -match '교복선정위원회 구성\(안\)') 'F-001 원문 대조 제목·수신·제목란이 존재함'
    Assert-Check ($wsF01.Range('B15').Value2 -match '위원회명：교복선정위원회' -and $wsF01.Range('B16').Value2 -match '구성인원' -and $wsF01.Range('B18').Value2 -eq '다. 역 할' -and $wsF01.Range('B24').Value2 -eq '라. 위원 명단' -and $wsF01.Range('B25').Value2 -eq '순' -and $wsF01.Range('C25').Value2 -eq '구분' -and $wsF01.Range('D25').Value2 -eq '직급(직위)' -and $wsF01.Range('E25').Value2 -eq '성명') 'F-001 위원회 구성 안내와 위원 명단 표 머리글이 원문과 일치함'
    Assert-Check ($wsF01.Range('C26').Value2 -eq '위원장' -and $wsF01.Range('D26').Value2 -eq '교원위원' -and $wsF01.Range('E26').Value2 -eq '위원 ○○' -and $wsF01.Range('C27').Value2 -eq '위원') 'F-001 위원 명단이 기초자료입력 위원 반복행(DB_위원 공용)에서 마스킹 식별표시로만 반영됨'
    $bodyText001 = ($wsF01.Range('B3:H43').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyText001 -notmatch '홍길동' -and $bodyText001 -notmatch '위원 ○○○') 'F-001 본문 어디에도 실명 형태 문자열이 출력되지 않음(마스킹 식별표시만 허용)'
    Assert-Check ($wsF01.Range('B37').Value2 -match '위원 수락 및 확인서' -and $wsF01.Range('B38').Value2 -match '청렴 및 보안 서약서' -and [bool]$excel.Run('검증_F001출력가능')) 'F-001 붙임 문구가 원문과 일치하고 선택 출력을 허용함'

    $wsF02 = $wb.Worksheets.Item('F-002_위원수락확인서')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF02.Range('B3').Value2 -eq '교복선정위원회 위원 수락 및 확인서' -and $wsF02.Range('B6').Value2 -match '공정하게 수행할 것이며' -and $wsF02.Range('B7').Value2 -match '전혀 관련이 없음을 확인합니다' -and $wsF02.Range('B8').Value2 -match '즉시 위원직을 사퇴') 'F-002 원문 대조 제목과 3단락 서약 본문이 원문과 일치함'
    Assert-Check ($wsF02.Range('B14').Value2 -match '주\s*소' -and $wsF02.Range('B16').Value2 -match '성\s*명' -and $wsF02.Range('B18').Value2 -match '서\s*명' -and [bool]$excel.Run('검증_F002출력가능')) 'F-002 주소·성명·서명란이 존재하고 선택 출력을 허용함'
    $bodyText002 = ($wsF02.Range('B3:E21').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyText002 -notmatch '홍길동') 'F-002 본문 어디에도 실명 형태 문자열이 출력되지 않음(주소·성명·서명 전부 수기 작성 공란)'

    $wsF03 = $wb.Worksheets.Item('F-003_위원청렴보안서약서')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF03.Range('B3').Value2 -eq '청렴 및 보안 서약서') 'F-003 B3 제목이 원문과 일치함'
    Assert-Check ($wsF03.Range('B6').Value2 -match '부패 없는 투명한 사회') 'F-003 B6 도입 문구에 부패 없는 투명한 사회가 포함됨'
    Assert-Check ($wsF03.Range('B6').Value2 -notmatch '○○학교' -and $wsF03.Range('B6').Value2 -match [regex]::Escape($wsIn.Range('C4').Value2)) 'F-003 B6 도입 문구에 학교 공통정보가 반영됨(기본값 미노출)'
    Assert-Check ($wsF03.Range('B8').Value2 -match '객관적이고 공정하게 심사' -and $wsF03.Range('B9').Value2 -match '금품' -and $wsF03.Range('B9').Value2 -match '향응' -and $wsF03.Range('B10').Value2 -match '보안관계 규정' -and $wsF03.Range('B11').Value2 -match '업무상 비밀을 타인에게 누설') 'F-003 4개 준수사항 항목이 원문과 일치함'
    Assert-Check ($wsF03.Range('B17').Value2 -match '서\s*약\s*자\s*성\s*명' -and [bool]$excel.Run('검증_F003출력가능')) 'F-003 서약자 성명·서명란이 존재하고 선택 출력을 허용함'
    $bodyText003 = ($wsF03.Range('B3:E20').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyText003 -notmatch '홍길동') 'F-003 본문 어디에도 실명 형태 문자열이 출력되지 않음(성명·서명·날짜 전부 수기 작성 공란)'

    $wsF04 = $wb.Worksheets.Item('F-004_구매추진계획수립기안문')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF04.Range('B3').Value2 -eq '교복 학교주관구매 추진 계획 수립(안)' -and $wsF04.Range('B7').Value2 -eq '수  신' -and $wsF04.Range('C7').Value2 -eq '내부결재' -and $wsF04.Range('C9').Value2 -match '교복 학교주관구매 추진 계획\(안\)') 'F-004 원문 대조 제목·수신·제목란이 존재함'
    Assert-Check ($wsF04.Range('B13').Value2 -match '공정하고 투명한 교복 선정' -and $wsF04.Range('B13').Value2 -notmatch '○○학교' -and $wsF04.Range('B15').Value2 -match '교복 학교주관구매 추진 계획\(안\) 1부' -and [bool]$excel.Run('검증_F004출력가능')) 'F-004 본문·붙임 문구가 원문과 일치하고 선택 출력을 허용함'

    $wsF05 = $wb.Worksheets.Item('F-005_구매추진계획안')
    $excel.CalculateFullRebuild()
    $bodyTextF05 = ($wsF05.UsedRange.Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF05 -match '교복 학교주관구매 추진 계획\(안\)' -and $bodyTextF05 -match '기본방향' -and $bodyTextF05 -match '추진 절차 개요' -and $bodyTextF05 -match '세부 추진 절차' -and $bodyTextF05 -match '기대효과') 'F-005 원문 대조 제목과 5개 장(개요/기본방향/추진절차개요/세부추진절차/기대효과) 표제가 모두 존재함'
    Assert-Check ($bodyTextF05 -match '단계 및 시기' -and $bodyTextF05 -match '학교운영위원회 심의' -and $bodyTextF05 -match '2단계 가격 경쟁') 'F-005 3장 추진 절차 개요 표가 원문과 일치함'
    Assert-Check ($bodyTextF05 -notmatch '○○학교' -and $bodyTextF05 -match [regex]::Escape($wsIn.Range('C4').Value2)) 'F-005 본문에 학교 공통정보가 반영되고 기본값이 노출되지 않음'
    Assert-Check ($bodyTextF05 -notmatch '홍길동' -and $bodyTextF05 -notmatch '위원 ○○○') 'F-005 위원 명단에 실명 형태 문자열이 출력되지 않음(마스킹 식별표시만 허용)'
    Assert-Check ([bool]$excel.Run('검증_F005출력가능')) 'F-005 선택 출력을 허용함'

    $wsF06 = $wb.Worksheets.Item('F-006_운영위원회심의안')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF06.Range('B3').Value2 -match '교복 학교주관구매 계획\(안\)' -and $wsF06.Range('B6').Value2 -eq '안건번호' -and $wsF06.Range('B8').Value2 -eq '제 안 자') 'F-006 원문 대조 제목·안건 정보란이 존재함'
    Assert-Check ($wsF06.Range('B12').Value2 -match '초·중등교육법' -and $wsF06.Range('B21').Value2 -eq '다. 교복구매 상한가격' -and $wsF06.Range('B22').Value2 -match '하복' -and $wsF06.Range('B22').Value2 -match '동복') 'F-006 제안 이유·상한가격 문구가 원문과 일치함'
    Assert-Check ($wsF06.Range('B24').Value2 -match '교복선정위원회 구성' -and $wsF06.Range('B25').Value2 -eq '순' -and $wsF06.Range('C25').Value2 -eq '구분' -and $wsF06.Range('D25').Value2 -eq '직위' -and $wsF06.Range('E25').Value2 -eq '성명') 'F-006 협의사항 위원회 구성 표 머리글이 원문과 일치함'
    Assert-Check ($wsF06.Range('C26').Value2 -eq '위원장' -and $wsF06.Range('D26').Value2 -eq '교원위원' -and $wsF06.Range('E26').Value2 -eq '위원 ○○') 'F-006 위원 명단이 기초자료입력 위원 반복행(F-001과 동일 공용)에서 마스킹 식별표시로만 반영됨'
    $bodyTextF06 = ($wsF06.Range('B3:H40').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF06 -notmatch '홍길동' -and $bodyTextF06 -notmatch '위원 ○○○') 'F-006 본문 어디에도 실명 형태 문자열이 출력되지 않음(마스킹 식별표시만 허용)'
    Assert-Check ($wsF06.Range('B40').Value2 -match '참고 서식' -and [bool]$excel.Run('검증_F006출력가능')) 'F-006 참고 서식 안내문이 존재하고 선택 출력을 허용함'

    $wsF32 = $wb.Worksheets.Item('F-032_제안서접수결과')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF32.Range('B3').Value2 -match '테스트초등학교' -and $wsF32.Range('B17').Value2 -eq '접수번호' -and $wsF32.Range('D18').Value2 -eq '업체명' -and $wsF32.Range('D19').Value2 -eq '검증업체가' -and $wsF32.Range('D20').Value2 -eq '검증업체나') 'F-032 원문 대조 표와 업체 반복행 2건이 자동 반영됨'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF32.Range('C19').Value2) -and $wsF32.Range('E19').Value2 -eq '박대표') 'F-032 접수일자는 실제 접수값이므로 수기 공란으로 보존되고, 대표자(R-08)는 2026-09-23부터 자동 반영됨'
    Assert-Check ([bool]$excel.Run('검증_F032출력가능')) 'F-032 선택 출력이 공통 필수값과 유효한 업체 반복행을 요구함'
    $wsIn.Range('B44:B53').ClearContents(); $excel.CalculateFullRebuild()
    Assert-Check (-not [bool]$excel.Run('검증_F032출력가능')) 'F-032 선택 출력이 업체명 누락을 거부함'
    $wsIn.Range('B44').Value2 = '검증업체가'; $wsIn.Range('B45').Value2 = '검증업체나'

    $wsF33 = $wb.Worksheets.Item('F-033_제안서접수대장')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF33.Range('B3').Value2 -match '2026학년도' -and $wsF33.Range('D8').Value2 -eq '검증업체가' -and $wsF33.Range('D9').Value2 -eq '검증업체나' -and $wsF33.Range('C8').Value2 -eq 1 -and $wsF33.Range('C9').Value2 -eq 2) 'F-033 원문 대조 대장 표와 업체 접수순번·업체명 반복행이 자동 반영됨'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF33.Range('B8').Value2) -and $wsF33.Range('E8').Value2 -eq '박대표' -and $wsF33.Range('F8').Value2 -eq '02-000-0000' -and [string]::IsNullOrWhiteSpace([string]$wsF33.Range('H8').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF33.Range('K8').Value2)) 'F-033 접수일·대리인·수령인은 자동 처리하지 않고 수기 공란으로 보존되며, 대표자·전화(R-08·R-09)는 2026-09-23부터 자동 반영됨'
    Assert-Check ([bool]$excel.Run('검증_F033출력가능')) 'F-033 선택 출력이 공통 필수값과 유효한 업체 반복행을 요구함'

    $wsF34 = $wb.Worksheets.Item('F-034_평가위원회개최')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF34.Range('B3').Value2 -match '테스트초등학교' -and $wsF34.Range('B20').Value2 -eq '순' -and $wsF34.Range('C20').Value2 -eq '구분' -and $wsF34.Range('D21').Value2 -eq '위원 ○○' -and $wsF34.Range('C21').Value2 -eq '위원장') 'F-034 원문 대조 위원회 표와 역할·마스킹 식별표시가 자동 반영됨'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF34.Range('C16').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF34.Range('C17').Value2)) 'F-034 개최 일시·장소는 문서별 실제값이므로 수기 공란으로 보존됨'
    Assert-Check ([bool]$excel.Run('검증_F034출력가능')) 'F-034 선택 출력이 공통 필수값과 유효한 위원 반복행을 요구함'
    $wsIn.Range('C58').Value2 = '홍길동'; $excel.CalculateFullRebuild()
    Assert-Check ([bool]$excel.Run('검증_F034출력가능')) 'F-034 선택 출력이 위원 실명 형태 식별표시를 허용함(2026-09-23 마스킹 폐지)'
    $wsIn.Range('C58').Value2 = '위원 ○○'

    $wsF35 = $wb.Worksheets.Item('F-035_정량평가결과')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF35.Range('B3').Value2 -match '테스트초등학교' -and $wsF35.Range('C19').Value2 -eq '검증업체가' -and $wsF35.Range('C20').Value2 -eq '검증업체나' -and $wsF35.Range('H19').Value2 -eq 50) 'F-035 원문 대조 표에 업체 반복행과 학교 정량평가 합계(10+10+15+15=50, 앞서 입력한 골든 데이터 기준)가 F-015와 동일 산식으로 반영됨'
    Assert-Check ([bool]$excel.Run('검증_F035출력가능')) 'F-035 선택 출력이 공통 필수값과 유효한 업체 반복행을 요구함'
    $wsIn.Range('C44:F44').ClearContents(); $excel.CalculateFullRebuild()
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF35.Range('H19').Value2)) 'F-035 합계가 학교 정량평가 점수 미입력 시 0이 아닌 공란으로 표시됨(빈 셀 참조가 0으로 강제되지 않도록 소스 셀 공백을 직접 확인함)'
    $wsIn.Range('C44').Value2 = 10; $wsIn.Range('D44').Value2 = 10; $wsIn.Range('E44').Value2 = 15; $wsIn.Range('F44').Value2 = 15
    $excel.CalculateFullRebuild()

    $wsF36 = $wb.Worksheets.Item('F-036_제안서평가결과')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF36.Range('B19').Value2 -eq '검증업체가' -and $wsF36.Range('C19').Value2 -eq 50 -and $wsF36.Range('D19').Value2 -eq 55 -and $wsF36.Range('E19').Value2 -eq 105 -and $wsF36.Range('F19').Value2 -eq '적격' -and $wsF36.Range('G19').Value2 -eq '가격개찰대상') 'F-036 정량(50)·정성(55) 합계·적부 판정(105점 → 적격)·비고가 골든 데이터 기준으로 정확히 계산됨'
    Assert-Check ([bool]$excel.Run('검증_F036출력가능')) 'F-036 선택 출력이 공통 필수값과 유효한 업체 반복행을 요구함'

    $wsF37 = $wb.Worksheets.Item('F-037_평가위원회참석등록부')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF37.Range('B3').Value2 -match '교복선정 평가위원회 참석 등록부' -and $wsF37.Range('C12').Value2 -eq '교원위원' -and [string]::IsNullOrWhiteSpace([string]$wsF37.Range('D12').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF37.Range('E12').Value2)) 'F-037이 학년도·위원 구분만 반영하고 참석 성명·서명은 빈칸으로 보존함'
    Assert-Check ([bool]$excel.Run('검증_F037출력가능')) 'F-037 선택 출력이 공통 필수값과 완전한 위원 반복행을 요구함'
    $wsIn.Range('B58').ClearContents(); $excel.CalculateFullRebuild()
    Assert-Check (-not [bool]$excel.Run('검증_F037출력가능')) 'F-037 선택 출력이 불완전 위원 반복행을 거부함'
    $wsIn.Range('B58').Value2 = '교원위원'

    $wsF38 = $wb.Worksheets.Item('F-038_업체참가등록부')
    $excel.CalculateFullRebuild()
    $wsIn.Range('B53').Value2 = '검증업체열'; $excel.CalculateFullRebuild()
    Assert-Check ($wsF38.Range('B3').Value2 -eq '제안서 설명 업체 참가 등록부' -and $wsF38.Range('B9').Value2 -eq '검증업체가' -and $wsF38.Range('B18').Value2 -eq '검증업체열' -and $wsF38.Range('C9').Value2 -eq '박대표' -and $wsF38.Range('D9').Value2 -eq '02-000-0000' -and [string]::IsNullOrWhiteSpace([string]$wsF38.Range('F9').Value2)) 'F-038이 업체 1~10행을 모두 반영하고 대리인 개인정보는 빈칸으로 보존하며, 대표자·전화(R-08·R-09)는 2026-09-23부터 자동 반영됨'
    $wsIn.Range('B53').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_F038출력가능')) 'F-038 선택 출력이 공통 필수값과 유효한 업체 반복행을 요구함'
    $wsIn.Range('B44:B53').ClearContents(); $excel.CalculateFullRebuild()
    Assert-Check (-not [bool]$excel.Run('검증_F038출력가능')) 'F-038 선택 출력이 업체명 누락을 거부함'
    $wsIn.Range('B44').Value2 = '검증업체가'; $wsIn.Range('B45').Value2 = '검증업체나'

    $wsF40 = $wb.Worksheets.Item('F-040_위원청렴보안서약서')
    $excel.CalculateFullRebuild()
    $bodyTextF40 = ($wsF40.Range('B3:H25').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF40 -match '청렴 및 보안 서약서' -and $bodyTextF40 -match '테스트초등학교' -and $bodyTextF40 -notmatch '홍길동' -and $wsF40.Range('B23').Value2 -match '성 명') 'F-040이 학교 공통정보만 반영하고 서약자 성명·서명은 자필 공란으로 보존함'
    Assert-Check ([bool]$excel.Run('검증_F040출력가능')) 'F-040 선택 출력이 공통 필수값을 요구함'

    $wsF41 = $wb.Worksheets.Item('F-041_낙찰자결정')
    $wsF41.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF41.Range('B3').Value2 -eq '교복 학교주관구매 낙찰자 결정(안)' -and $wsF41.Range('B7').Value2 -eq '수  신' -and $wsF41.Range('C7').Value2 -eq '내부결재' -and $wsF41.Range('B19').Value2 -eq '라. 낙찰자') 'F-041 원문 대조 제목·수신·낙찰자 항목 라벨이 존재함'
    Assert-Check ($wsF41.Range('B20').Value2 -eq '순위' -and $wsF41.Range('C20').Value2 -eq '업체명' -and $wsF41.Range('D20').Value2 -eq '대표자' -and $wsF41.Range('E20').Value2 -eq '투찰금액(원)' -and $wsF41.Range('F20').Value2 -eq '낙찰률(%)' -and $wsF41.Range('G20').Value2 -eq '비고(주소)') 'F-041 낙찰자 표 머리글(순위/업체명/대표자/투찰금액/낙찰률/비고)이 원문과 일치함'
    Assert-Check ($wsF41.Range('C21').Value2 -eq '검증업체가' -and $wsF41.Range('B23').Value2 -match '지방자치단체를 당사자로 하는 계약에 관한 법률 시행령') 'F-041이 I3 순번 1의 낙찰 업체명을 반영하고 근거 법령 문구가 원문과 일치함'
    $wsF41.Range('I3').Value2 = 2
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF41.Range('C21').Value2 -eq '검증업체나') 'F-041이 선택 순번을 바꾸면 다른 낙찰 업체명을 반영함'
    $wsF41.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF41.Range('D21').Value2 -eq '박대표' -and [string]::IsNullOrWhiteSpace([string]$wsF41.Range('G21').Value2)) 'F-041 비고(주소)는 업체 상세정보이므로 빈칸으로 보존되고, 대표자(R-08)는 2026-09-23부터 자동 반영됨'
    $wsF41.Range('C17').Value2 = 4500000; $wsF41.Range('E21').Value2 = 4400000; $wsF41.Range('F21').Value2 = 98
    Assert-Check ([bool]$excel.Run('검증_F041출력가능')) 'F-041 선택 출력이 낙찰 업체·예정가격·투찰금액·낙찰률이 모두 있을 때만 허용함'
    $wsF41.Range('E21').ClearContents()
    Assert-Check (-not [bool]$excel.Run('검증_F041출력가능')) 'F-041 선택 출력이 업체명만 있고 투찰금액이 비어 있으면 거부함'
    $wsF41.Range('E21').Value2 = 4400000
    $bodyTextF41 = ($wsF41.Range('B3:H29').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF41 -match '박대표') 'F-041 본문에 대표자(R-08) 실명이 반영됨(2026-09-23부터 허용)'

    $wsF42 = $wb.Worksheets.Item('F-042_낙찰자결정통보')
    $wsF42.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF42.Range('B3').Value2 -eq '교복 학교주관구매 낙찰자 결정 통보' -and $wsF42.Range('B7').Value2 -eq '수  신' -and $wsF42.Range('C7').Value2 -eq '검증업체가' -and $wsF42.Range('B15').Value2 -eq '가. 낙찰자') 'F-042 원문 대조 제목이 존재하고 수신이 선택 낙찰 업체명으로 반영됨(통보문 유형)'
    Assert-Check ($wsF42.Range('B16').Value2 -eq '업체명' -and $wsF42.Range('C16').Value2 -eq '대표자' -and $wsF42.Range('D16').Value2 -eq '단가(원)' -and $wsF42.Range('E16').Value2 -eq '예정수량' -and $wsF42.Range('F16').Value2 -eq '계약체결금액(원)' -and $wsF42.Range('G16').Value2 -eq '비고') 'F-042 낙찰자 표 머리글이 원문과 일치함'
    Assert-Check ($wsF42.Range('B17').Value2 -eq '검증업체가' -and $wsF42.Range('E17').Value2 -eq 100) 'F-042가 선택 업체명과 예정수량(B-05)을 반영함'
    Assert-Check ($wsF42.Range('B13').Value2 -match '2026\.02\.20\.') 'F-042 제출기한(B-07) 안내문이 기초자료입력 값을 반영함'
    Assert-Check ($wsF42.Range('B20').Value2 -match '품목별 산출내역서' -and $wsF42.Range('B21').Value2 -match '원단 의류 시험성적서' -and $wsF42.Range('B22').Value2 -match '사용인장계' -and $wsF42.Range('B23').Value2 -match '법인 등기사항증명서') 'F-042 제출서류 4개 항목이 원문과 일치함'
    Assert-Check ($wsF42.Range('C17').Value2 -eq '박대표' -and [string]::IsNullOrWhiteSpace([string]$wsF42.Range('G17').Value2)) 'F-042 비고는 업체 상세정보이므로 빈칸으로 보존되고, 대표자(R-08)는 2026-09-23부터 자동 반영됨'
    $wsF42.Range('D17').Value2 = 44000; $wsF42.Range('F17').Value2 = 4400000
    Assert-Check ([bool]$excel.Run('검증_F042출력가능')) 'F-042 선택 출력이 낙찰 업체·단가·예정수량·반올림 계약금액·제출기한이 모두 일치할 때만 허용함'
    $wsF42.Range('F17').Value2 = 4400001
    Assert-Check (-not [bool]$excel.Run('검증_F042출력가능')) 'F-042 선택 출력이 단가×예정수량과 불일치한 계약체결금액을 거부함'
    $wsF42.Range('F17').Value2 = 4400000
    $wsF42.Range('F17').ClearContents()
    Assert-Check (-not [bool]$excel.Run('검증_F042출력가능')) 'F-042 선택 출력이 업체명만 있고 계약체결금액이 비어 있으면 거부함'
    $wsF42.Range('F17').Value2 = 4400000
    $bodyTextF42 = ($wsF42.Range('B3:H29').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF42 -match '박대표') 'F-042 본문에 대표자(R-08) 실명이 반영됨(2026-09-23부터 허용)'

    $wsF43 = $wb.Worksheets.Item('F-043_계약체결')
    $wsF43.Range('I3').Value2 = 1
    $wsF43.Range('C16').Value2 = 50000
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF43.Range('B3').Value2 -eq '교복 학교주관구매 계약 체결(안)' -and $wsF43.Range('B7').Value2 -eq '수  신' -and $wsF43.Range('C7').Value2 -eq '내부결재' -and $wsF43.Range('B20').Value2 -eq '3. 계약상대자') 'F-043 원문 대조 제목·수신·계약상대자 라벨이 존재함'
    Assert-Check ($wsF43.Range('C17').Value2 -eq '금5,000,000원[50,000원(단가)×100벌(예정수량)]') 'F-043이 담당자 입력 단가×예정수량(B-05)으로 계약금액을 자동 계산함(K-01/K-02)'
    Assert-Check ($wsF43.Range('B21').Value2 -eq '업체명' -and $wsF43.Range('C21').Value2 -eq '소재지' -and $wsF43.Range('D21').Value2 -eq '대표자' -and $wsF43.Range('E21').Value2 -eq '전화번호') 'F-043 계약상대자 표 머리글이 원문과 일치함'
    Assert-Check ($wsF43.Range('B22').Value2 -eq '검증업체가') 'F-043이 I3 순번 1의 계약상대자 업체명을 반영함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF43.Range('C22').Value2) -and $wsF43.Range('D22').Value2 -eq '박대표' -and $wsF43.Range('E22').Value2 -eq '02-000-0000') 'F-043 소재지는 업체 상세정보이므로 빈칸으로 보존되고, 대표자·전화번호(R-08·R-09)는 2026-09-23부터 자동 반영됨'
    Assert-Check ($wsF43.Range('B25').Value2 -match '2026\.03\.01\.') 'F-043 납품기한(B-06)이 기초자료입력 값을 반영함'
    Assert-Check ($wsF43.Range('B28').Value2 -match '계약서\(초안\)' -and $wsF43.Range('B29').Value2 -match '교복 사양서' -and $wsF43.Range('B30').Value2 -match '교복 구매계약 특수조건' -and $wsF43.Range('B31').Value2 -match '입찰가격산출내역서' -and $wsF43.Range('B32').Value2 -match '기타서류') 'F-043 붙임 5개 항목이 원문과 일치함'
    Assert-Check ([bool]$excel.Run('검증_F043출력가능')) 'F-043 선택 출력이 계약상대자·단가·예정수량·납품기한이 모두 있을 때만 허용함'
    $wsF43.Range('C16').Value2 = ''; $excel.CalculateFullRebuild()
    Assert-Check (-not [bool]$excel.Run('검증_F043출력가능')) 'F-043 선택 출력이 업체명만 있고 계약 단가가 비어 있으면 거부함'
    $wsF43.Range('C16').Value2 = 50000; $excel.CalculateFullRebuild()
    $bodyTextF43 = ($wsF43.Range('B3:H37').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF43 -match '박대표') 'F-043 본문에 대표자(R-08) 실명이 반영됨(2026-09-23부터 허용)'

    $wsF44 = $wb.Worksheets.Item('F-044_사전안내가정통신문안내')
    Assert-Check ($wsF44.Range('B3').Value2 -eq '신입생 교복구매 사전 안내 가정통신문 발송' -and $wsF44.Range('B7').Value2 -eq '수  신' -and $wsF44.Range('C7').Value2 -eq '내부결재(사업부서)' -and $wsF44.Range('B15').Value2 -eq '가. 대상:') 'F-044 원문 대조 제목·수신·대상 항목 라벨이 존재함'
    Assert-Check ($wsF44.Range('G2').Value2 -eq '테스트초등학교' -and $wsF44.Range('C9').Value2 -eq '2026학년도 동복 학교주관구매 요청' -and $wsF44.Range('C15').Value2 -eq '2026학년도 예비 중1(예비 고1) 학생 및 학부모' -and $wsF44.Range('C16').Value2 -match '2026학년도 동복 구매 관련 교복 학교주관구매 참여 및 지원 방법' -and $wsF44.Range('B18').Value2 -match '교복구매 사전 안내 가정통신문 1부') 'F-044가 학교명·제목·학년도·구매명과 원문 안내 내용을 반영함'
    Assert-Check ([bool]$excel.Run('검증_F044출력가능')) 'F-044 선택 출력이 공통 필수값이 있을 때 허용함'
    $bodyTextF44 = ($wsF44.Range('B3:H25').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF44 -notmatch '홍길동' -and $bodyTextF44 -notmatch '010-') 'F-044가 학생·학부모 개인정보를 자동 출력하지 않음'

    $wsF45 = $wb.Worksheets.Item('F-045_사전안내가정통신문')
    Assert-Check ($wsF45.Range('B6').Value2 -eq '안녕하십니까? 졸업과 함께 상급학교로 진학함을 축하드립니다.' -and $wsF45.Range('B19').Value2 -eq '감사합니다.') 'F-045 원문 대조 인사말·맺음말 문구가 존재함'
    Assert-Check ($wsF45.Range('B3').Value2 -eq '2026학년도 동복 학교주관구매 요청' -and $wsF45.Range('B8').Value2 -match '2026학년도 중학교, 고등학교 신입생 배정을 앞두고 2026학년도 동복 구매에 대한 사항' -and $wsF45.Range('B22').Value2 -eq '테스트초등학교장') 'F-045가 학교명·학년도·구매명·제목 공통값을 반영함'
    Assert-Check ([bool]$excel.Run('검증_F045출력가능')) 'F-045 선택 출력이 공통 필수값이 있을 때 허용함'
    $bodyTextF45 = ($wsF45.Range('B3:H22').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF45 -notmatch '홍길동' -and $bodyTextF45 -notmatch '010-') 'F-045가 학생·학부모 개인정보를 자동 출력하지 않음'

    $wsF46 = $wb.Worksheets.Item('F-046_수요조사가정통신문안내')
    Assert-Check ($wsF46.Range('B3').Value2 -eq '신입생 예비소집일 교복구매 수요조사 가정통신문 발송' -and $wsF46.Range('B7').Value2 -eq '수  신' -and $wsF46.Range('C7').Value2 -eq '내부결재(사업부서)' -and $wsF46.Range('B15').Value2 -eq '가. 대상:') 'F-046 원문 대조 제목·수신·대상 항목 라벨이 존재함'
    Assert-Check ($wsF46.Range('G2').Value2 -eq '테스트초등학교' -and $wsF46.Range('C9').Value2 -eq '2026학년도 동복 학교주관구매 요청' -and $wsF46.Range('C15').Value2 -eq '2026학년도 신입생 및 학부모' -and $wsF46.Range('C16').Value2 -match '2026학년도 동복 구매 관련 교복 학교주관구매 참여 여부 조사' -and $wsF46.Range('B18').Value2 -match '교복구매 수요조사 가정통신문 1부') 'F-046이 학교명·제목·학년도·구매명과 원문 안내 내용을 반영함'
    Assert-Check ([bool]$excel.Run('검증_F046출력가능')) 'F-046 선택 출력이 공통 필수값이 있을 때 허용함'
    $bodyTextF46 = ($wsF46.Range('B3:H25').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF46 -notmatch '홍길동' -and $bodyTextF46 -notmatch '010-') 'F-046이 학생·학부모 개인정보를 자동 출력하지 않음'

    $wsF48 = $wb.Worksheets.Item('F-048_교복구매안내신청수량파악')
    Assert-Check ($wsF48.Range('B6').Value2 -match '테스트초등학교 입학을 진심으로 축하드립니다' -and $wsF48.Range('B26').Value2 -match '절 취 선' -and $wsF48.Range('B33').Value2 -eq '【예시】교복 신청 수량(품목별)') 'F-048 원문 대조 인사말·절취선·예시 라벨이 존재함'
    Assert-Check ($wsF48.Range('B3').Value2 -eq '테스트초등학교 교복 구매 안내문' -and $wsF48.Range('B17').Value2 -match '2026\.02\.20\.' -and $wsF48.Range('B24').Value2 -eq '테스트초등학교장' -and $wsF48.Range('B28').Value2 -eq '< 2026학년도 신입생 교복 구매 신청서 >') 'F-048이 학교명·제출기한(B-07)·학교장·학년도 공통값을 반영함'
    Assert-Check ($wsF48.Range('B35').Value2 -eq '품목(남학생)' -and $wsF48.Range('F35').Value2 -eq '품목(여학생)' -and $wsF48.Range('B36').Value2 -eq '자켓' -and $wsF48.Range('B42').Value2 -eq '합계' -and [string]::IsNullOrWhiteSpace([string]$wsF48.Range('C36').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF48.Range('D36').Value2)) 'F-048 신청 수량 표 머리글·품목명만 존재하고 단가·수량·구매가격은 전부 공란임(신청 수량 미수집)'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF48.Range('B30').Value2) -eq $false -and $wsF48.Range('B30').Value2 -match '이름' -and [string]::IsNullOrWhiteSpace([string]$wsF48.Range('B31').Value2) -eq $false -and $wsF48.Range('B31').Value2 -match '보호자 성명') 'F-048 이름·보호자 성명·전화번호 응답란은 빈 괄호 안내 문구만 있고 실제 값이 입력·저장되지 않음'
    Assert-Check ([bool]$excel.Run('검증_F048출력가능')) 'F-048 선택 출력이 공통 필수값이 있을 때 허용함'
    $bodyTextF48 = ($wsF48.Range('B3:I45').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF48 -notmatch '홍길동' -and $bodyTextF48 -notmatch '010-') 'F-048이 학생·학부모 개인정보를 자동 출력하지 않음'

    $wsF47 = $wb.Worksheets.Item('F-047_수요조사가정통신문')
    Assert-Check ($wsF47.Range('B3').Value2 -eq '2026학년도 동복 학교주관구매 요청' -and $wsF47.Range('B6').Value2 -match '테스트초등학교 입학을 진심으로 축하드립니다' -and $wsF47.Range('B29').Value2 -match '절 취 선' -and $wsF47.Range('B31').Value2 -eq '< 교복 구매 신청서 >') 'F-047 원문 대조 표제(D-02 우선 반영)·인사말·절취선·신청서 라벨이 존재함'
    Assert-Check ($wsF47.Range('B17').Value2 -match '2026\.02\.20\.' -and $wsF47.Range('B27').Value2 -eq '테스트초등학교장') 'F-047이 제출기한(B-07)·학교장 공통값을 반영함'
    Assert-Check ($wsF47.Range('B34').Value2 -match '학교주관 교복구매.*참여' -and $wsF47.Range('D34').Value2 -match '미참여' -and [string]::IsNullOrWhiteSpace([string]$wsF47.Range('B36').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF47.Range('D36').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF47.Range('F36').Value2)) 'F-047 참여/미참여 응답 표는 머리글만 있고 응답 셀 3칸은 전부 공란임(신청 여부 미수집)'
    Assert-Check ($wsF47.Range('B40').Value2 -match '이름' -and $wsF47.Range('B41').Value2 -match '보호자 성명' -and $wsF47.Range('B39').Value2 -eq '년   월   일') 'F-047 이름·보호자 성명·전화번호 응답란은 빈 괄호 안내 문구만 있고 실제 값이 입력·저장되지 않음'
    Assert-Check ([bool]$excel.Run('검증_F047출력가능')) 'F-047 선택 출력이 공통 필수값이 있을 때 허용함'
    $bodyTextF47 = ($wsF47.Range('B3:H41').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF47 -notmatch '홍길동' -and $bodyTextF47 -notmatch '010-') 'F-047이 학생·학부모 개인정보를 자동 출력하지 않음'

    $wsF49 = $wb.Worksheets.Item('F-049_만족도설문조사실시')
    Assert-Check ($wsF49.Range('B3').Value2 -eq '교복 학교주관구매 만족도 설문조사 실시' -and $wsF49.Range('B7').Value2 -eq '수  신' -and $wsF49.Range('C7').Value2 -eq '내부결재(사업부서)' -and $wsF49.Range('B15').Value2 -eq '가. 대상:') 'F-049 원문 대조 제목·수신·대상 항목 라벨이 존재함'
    Assert-Check ($wsF49.Range('G2').Value2 -eq '테스트초등학교' -and $wsF49.Range('C9').Value2 -eq '2026학년도 동복 학교주관구매 요청' -and $wsF49.Range('C15').Value2 -eq '2026학년도 교복 학교주관구매 참여 학생 및 학부모' -and $wsF49.Range('C16').Value2 -match '2026학년도 동복 구매 관련 가격, 품질, 구매절차 및 사후관리 등 전반에 걸친 만족도 조사' -and $wsF49.Range('B18').Value2 -match '교복 학교주관구매 만족도 설문지 1부') 'F-049가 학교명·제목·학년도·구매명과 원문 안내 내용을 반영함'
    Assert-Check ([bool]$excel.Run('검증_F049출력가능')) 'F-049 선택 출력이 공통 필수값이 있을 때 허용함'
    $bodyTextF49 = ($wsF49.Range('B3:H25').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF49 -notmatch '홍길동' -and $bodyTextF49 -notmatch '010-') 'F-049가 응답자 개인정보를 자동 출력하지 않음'

    $wsF09 = $wb.Worksheets.Item('F-009_기초금액계약방법결정')
    Assert-Check ($wsF09.Range('B3').Value2 -eq '교복 학교주관구매 기초금액 및 계약방법 결정' -and $wsF09.Range('B19').Value2 -eq '다. 기초금액：' -and $wsF09.Range('B20').Value2 -eq '라. 납품기한：' -and $wsF09.Range('B21').Value2 -eq '마. 계약방법：') 'F-009 원문 대조 제목·기초금액·납품기한·계약방법 항목이 존재함'
    Assert-Check ($wsF09.Range('C19').Value2 -eq '금5,000,000원' -and $wsF09.Range('C20').Value2 -match '2026\.\s*3\.\s*1\.' -and $wsF09.Range('C21').Value2 -match '2단계 입찰') 'F-009가 기초금액·납품기한·계약방법 공통값을 반영함'
    Assert-Check ([bool]$excel.Run('검증_F009출력가능')) 'F-009 선택 출력이 기초금액·납품기한·계약방법이 모두 있을 때 허용함'
    $wsIn.Range('C15').ClearContents(); $excel.CalculateFullRebuild()
    Assert-Check (-not [bool]$excel.Run('검증_F009출력가능')) 'F-009 선택 출력이 기초금액 누락 시 거부됨'
    $wsIn.Range('C15').Value2 = 5000000; $excel.CalculateFullRebuild()

    $wsF10 = $wb.Worksheets.Item('F-010_사전규격공개기안문')
    Assert-Check ($wsF10.Range('B17').Value2 -eq '가. 사 업 명：' -and $wsF10.Range('B18').Value2 -eq '나. 기초금액：' -and $wsF10.Range('B20').Value2 -eq '라. 공개방법：' -and $wsF10.Range('C20').Value2 -match 'G2B') 'F-010 원문 대조 사업명·기초금액·공개방법 항목이 존재함'
    Assert-Check ($wsF10.Range('C18').Value2 -match '5,000,000' -and $wsF10.Range('C19').Value2 -match '100벌' -and $wsF10.Range('C23').Value2 -match '2026') 'F-010이 기초금액·예정수량·의견마감 공통값을 반영함'
    Assert-Check ([bool]$excel.Run('검증_F010출력가능')) 'F-010 선택 출력이 필수 공통값이 있을 때 허용함'
    $wsIn.Range('C18').ClearContents(); $excel.CalculateFullRebuild()
    Assert-Check (-not [bool]$excel.Run('검증_F010출력가능')) 'F-010 선택 출력이 의견등록마감일시 누락 시 거부됨'
    $wsIn.Range('C18').Formula = '=DATE(2026,2,20)+18/24'; $excel.CalculateFullRebuild()

    $wsF11 = $wb.Worksheets.Item('F-011_입찰공고')
    # 2026-09-24 사용자 요청: 제목 위 빈 여백행(원 2행)·문서번호(원 3행, 기초자료 중복 표시라
    # 불필요)를 삭제해 아래 모든 행이 2행씩 위로 당겨짐(문서번호는 결재란의 "시행"에 이미 있음).
    Assert-Check ($wsF11.Range('B10').Value2 -eq '1. 입찰에 부치는 사항' -and $wsF11.Range('B18').Value2 -eq '특기사항' -and $wsF11.Range('B20').Value2 -eq '2. 입찰방법' -and $wsF11.Range('B24').Value2 -eq '3. 입찰 참가자격' -and $wsF11.Range('B32').Value2 -eq '5. 제안서 제출' -and $wsF11.Range('B56').Value2 -eq '11. 계약체결 및 이행') 'F-011 원문 대조 입찰표와 1~12항의 핵심 구조가 존재함(2026-09-24 2행 삭제 반영)'
    $f11SingleValueRows = 0
    for ($f11Row = 11; $f11Row -le 17; $f11Row++) {
        if ($wsF11.Range("C${f11Row}:H${f11Row}").MergeCells -and $wsF11.Range("C$f11Row").HorizontalAlignment -eq -4131) { $f11SingleValueRows++ }
    }
    Assert-Check ($f11SingleValueRows -eq 7) 'F-011 입찰건명 등 단일값 행 7개가 값 영역 전체 병합·좌측 정렬로 표시됨'
    Assert-Check ($wsF11.Range('C12').Value2 -match '5,000,000' -and $wsF11.Range('C13').Value2 -match '100벌' -and $wsF11.Range('C16').Value2 -match '2026' -and $wsF11.Range('B33').Value2 -match '2026') 'F-011이 기초금액·예정수량·납품기한·제출기한 공통값을 반영함'
    Assert-Check ($wsF11.Range('A1').MergeCells -and $wsF11.Range('A1').WrapText -and $wsF11.Rows.Item(1).RowHeight -ge 28) 'F-011 상단 안내문이 전체 폭·줄바꿈·행높이로 인쇄 잘림을 방지함(2026-09-23 조판 수정)'
    Assert-Check ($wsF11.PageSetup.Zoom -eq 100 -and -not $wsF11.PageSetup.FitToPagesWide -and -not $wsF11.PageSetup.FitToPagesTall -and $wsF11.VPageBreaks.Count -eq 0 -and $wsF11.HPageBreaks.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$wsF11.PageSetup.PrintArea)) 'F-011이 자연 배율에서 가로 분할 없이 세로 다쪽 인쇄영역을 사용함'
    Assert-Check ([bool]$excel.Run('검증_F011출력가능')) 'F-011 선택 출력이 필수 공통값이 있을 때 허용함'
    $wsIn.Range('C17').ClearContents(); $excel.CalculateFullRebuild()
    Assert-Check (-not [bool]$excel.Run('검증_F011출력가능')) 'F-011 선택 출력이 납품기한 누락 시 거부됨'
    $wsIn.Range('C17').Formula = '=DATE(2026,3,1)'; $excel.CalculateFullRebuild()

    $wsF12 = $wb.Worksheets.Item('F-012_계약특수조건')
    Assert-Check ($wsF12.Range('B3').Value2 -match '교복 학교주관구매 계약 특수조건' -and $wsF12.Range('B3').Value2 -match '2026' -and $wsF12.Range('B3').Value2 -match '테스트초등학교') 'F-012 제목이 학교명·학년도 공통값을 반영함'
    Assert-Check ($wsF12.Range('A1').MergeCells -and $wsF12.Range('A1').WrapText -and $wsF12.Rows.Item(1).RowHeight -ge 28 -and $wsF12.Range('B3').Font.Size -le 12 -and $wsF12.Range('B3').WrapText) 'F-012 안내문·제목이 긴 값에서도 인쇄 폭 안에 표시되도록 설정됨(2026-09-23 조판 수정)'
    Assert-Check ($wsF12.Range('B6').Value2 -match '이하 「학교」라 한다' -and $wsF12.Range('B6').Value2 -match '이하 「공급자」라 한다') 'F-012 도입 문구가 원문과 일치함(계약상대자명은 수기 확인 대상으로 보존)'
    $f12Headers = @($wsF12.UsedRange.Rows | ForEach-Object { $_.Cells.Item(1, 2).Value2 } | Where-Object { $_ -match '^제\d+조' })
    Assert-Check (($f12Headers -join '|') -match '제1조' -and ($f12Headers -join '|') -match '제9조\(하자보증\)' -and ($f12Headers -join '|') -match '제19조\(소송관할\)' -and $f12Headers.Count -eq 19) 'F-012 제1조~제19조 19개 조항 제목이 모두 원문과 일치함'
    $f12Art1Row = @($wsF12.UsedRange.Rows | Where-Object { $_.Cells.Item(1, 2).Value2 -eq '제1조(구매 수량 및 구매 단가)' } | Select-Object -First 1)
    $f12Art1BodyRow = if ($f12Art1Row.Count -gt 0) { $f12Art1Row[0].Row + 1 } else { 0 }
    Assert-Check ($wsF12.Cells.Item($f12Art1BodyRow, 2).Value2 -match '교복 구매 수량은 2026학년도 본교 신입생') 'F-012 제1조 학년도가 기초자료입력!C5를 반영함(나머지 문구는 원문 유지)'
    $f12Art3Row = @($wsF12.UsedRange.Rows | Where-Object { $_.Cells.Item(1, 2).Value2 -eq '제3조(납품 기한 및 장소)' } | Select-Object -First 1)
    $f12Art3BodyRow = if ($f12Art3Row.Count -gt 0) { $f12Art3Row[0].Row + 1 } else { 0 }
    Assert-Check ($wsF12.Cells.Item($f12Art3BodyRow, 2).Value2 -match '동복: 2026\.3\.1\.\(일\)까지, 학교에서 지정하는 장소' -and $wsF12.Cells.Item($f12Art3BodyRow, 2).Value2 -match '하복: 2026\.3\.1\.\(일\)까지, 학교에서 지정하는 장소') 'F-012 제3조 납품기한 날짜가 기초자료입력!C17을 반영함(동복·하복 공통, 업체명·장소 문구는 원문 유지)'
    Assert-Check ($wsF12.PageSetup.Zoom -eq 100 -and -not $wsF12.PageSetup.FitToPagesWide -and -not $wsF12.PageSetup.FitToPagesTall -and $wsF12.VPageBreaks.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$wsF12.PageSetup.PrintArea)) 'F-012가 자연 배율에서 가로 분할 없이 다쪽 인쇄영역을 사용함(참고자료 성격상 세로 다중 페이지 허용)'
    Assert-Check ([bool]$excel.Run('검증_F012출력가능')) 'F-012 선택 출력이 공통 필수값이 있을 때 허용함'
    $wsIn.Range('C17').ClearContents(); $excel.CalculateFullRebuild()
    Assert-Check ($wsF12.Cells.Item($f12Art3BodyRow, 2).Value2 -match '동복: 20○○\.OO\.OO\(요일\)까지, 학교에서 지정하는 장소') 'F-012 제3조가 납품기한 공란 시 원문 플레이스홀더(20○○.OO.OO)를 그대로 표시함'
    $wsIn.Range('C17').Formula = '=DATE(2026,3,1)'; $excel.CalculateFullRebuild()

    $wsF08 = $wb.Worksheets.Item('F-008_교복사양서')
    Assert-Check ($wsF08.Range('B3').Value2 -match '테스트초등학교' -and $wsF08.Range('B3').Value2 -match '교복 사양서') 'F-008 제목이 학교명 공통값을 반영함'
    Assert-Check ($wsF08.Range('B6').Value2 -match '섬유혼용률' -and $wsF08.Range('B22').Value2 -match '교복 디자인 및 규격서' -and $wsF08.Range('C23').Value2 -eq '디자인 이미지(전,후면)') 'F-008 원문 대조 섬유혼용률·디자인 규격서 표 구조가 존재함'
    Assert-Check ($wsF08.Range('B8').Value2 -match '동복 상의' -or $wsF08.Range('B24').Value2 -match '동복 상의') 'F-008이 품목명(R-04)을 기초자료입력 품목 반복행에서 반영함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF08.Range('E8').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF08.Range('E24').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF08.Range('G24').Value2)) 'F-008 섬유혼용률·원단 및 색상·제작시 유의사항 칸이 자동 반영 없이 공란임(담당자 작성 대상)'
    Assert-Check ($wsF08.PageSetup.Zoom -eq 100 -and -not $wsF08.PageSetup.FitToPagesWide -and -not $wsF08.PageSetup.FitToPagesTall -and $wsF08.VPageBreaks.Count -eq 0) 'F-008이 자연 배율에서 가로 분할 없이 A4에 출력됨'
    Assert-Check ([bool]$excel.Run('검증_F008출력가능')) 'F-008 선택 출력이 유효한 품목 반복행이 있을 때 허용함'

    $wsF50 = $wb.Worksheets.Item('F-050_만족도조사설문지')
    Assert-Check ($wsF50.PageSetup.Zoom -eq 100 -and $wsF50.VPageBreaks.Count -eq 0 -and -not $wsF50.PageSetup.FitToPagesWide -and -not $wsF50.PageSetup.FitToPagesTall) 'F-050이 자연 배율이며 가로 폭 초과 없이 A4에 출력됨'
    Assert-Check ($wsF50.Range('B3').Value2 -eq '2026학년도 동복 학교주관구매 요청' -and $wsF50.Range('B6').Value2 -match '교복 품질 향상' -and $wsF50.Range('B9').Value2 -eq '테스트초등학교 교복선정위원회' -and $wsF50.Range('B10').Value2 -match '2026학년도') 'F-050이 설문 제목(D-02)·학교명·학년도의 비식별 공통값과 원문 안내를 반영함'
    Assert-Check ($wsF50.Range('B12').Value2 -eq '영역' -and $wsF50.Range('C12').Value2 -eq '구분' -and $wsF50.Range('D12').Value2 -eq '내용' -and $wsF50.Range('E12').Value2 -eq '매우 만족' -and $wsF50.Range('I12').Value2 -eq '매우 불만' -and $wsF50.Range('D14').Value2 -eq '교복 가격에 만족한다.' -and $wsF50.Range('D23').Value2 -match 'A/S 절차') 'F-050 원문 대조 문항표 머리글·10개 문항의 첫·마지막 문항이 존재함'
    $responseCellsF50 = @($wsF50.Range('E14:I23').Cells) + @($wsF50.Range('E24:I24').Cells) + @($wsF50.Range('B32:I35').Cells)
    $blankResponsesF50 = @($responseCellsF50 | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Value2) -or -not [string]::IsNullOrWhiteSpace([string]$_.Formula) })
    Assert-Check ($blankResponsesF50.Count -eq 0 -and $wsF50.Range('H26').Value2 -match '^합계 \(\s*\)점$' -and $wsF50.Range('H27').Value2 -match '^평균 \(\s*\)점$') 'F-050 문항별 응답·소계·기타 의견은 수식·기본값 없이 완전 공란이고 총점·평균은 수기 기입 안내만 존재함'
    Assert-Check ([bool]$wsF50.ProtectContents) 'F-050이 매크로 비활성 상태에서도 직접 입력을 막는 인쇄 전용 시트 보호 상태임'
    # COM 검증 사본은 저장하지 않으므로, 이후 다른 시트의 계산·저장 회귀가 F-050 UI 보호와
    # 상호작용하지 않도록 보호를 해제한다. 배포 후보는 구조 생성 단계에서 이미 보호된 채 저장된다.
    $wsF50.Unprotect()

    # 계획.md 8.1절 후순위 보류 4종(F-027·F-029·F-030·F-031, 2026-09-21 구현)
    $wsF27 = $wb.Worksheets.Item('F-027_위임장')
    Assert-Check ($wsF27.Range('B3').Value2 -eq '위 임 장' -and $wsF27.Range('B6').Value2 -eq '○ 상 호 :' -and $wsF27.Range('B9').Value2 -eq '○ 주 소 :') 'F-027 원문 대조 표제·업체 정보 라벨이 존재함'
    Assert-Check ($wsF27.Range('B12').Value2 -match '테스트초등학교' -and $wsF27.Range('B12').Value2 -match '2026학년도' -and $wsF27.Range('B12').Value2 -match '일체의 권한을 다음의 사람에게 위임합니다') 'F-027 위임 문구가 학교명·학년도 공통값을 반영함'
    $bodyTextF27 = ($wsF27.Range('B6:F17').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF27.Range('C16').Value2) -and $bodyTextF27 -notmatch '홍길동' -and $bodyTextF27 -notmatch '\d{6}-\d{7}') 'F-027 위임 업체·수임인의 상호·사업자번호·주민등록번호·전화번호·주소가 전부 공란임'
    Assert-Check ([bool]$excel.Run('검증_F027출력가능')) 'F-027 선택 출력이 공통 필수값이 있을 때 허용함'
    Assert-Check ($wsF27.PageSetup.Zoom -eq 100 -and -not $wsF27.PageSetup.FitToPagesWide -and -not $wsF27.PageSetup.FitToPagesTall -and $wsF27.HPageBreaks.Count -eq 0 -and $wsF27.VPageBreaks.Count -eq 0) 'F-027이 자연 배율에서 A4 1쪽으로 출력됨'

    $wsF29 = $wb.Worksheets.Item('F-029_개인정보제공동의서')
    Assert-Check ($wsF29.Range('B3').Value2 -eq '개인정보제공 동의서' -and $wsF29.Range('B6').Value2 -match '개인정보보호법' -and $wsF29.Range('B6').Value2 -match '테스트초등학교') 'F-029 원문 대조 표제·동의 문구가 학교명 공통값을 반영함'
    Assert-Check ($wsF29.Range('B9').Value2 -match '수집 및 이용 목적' -and $wsF29.Range('B16').Value2 -eq '동의 □                    동의하지 않음 □') 'F-029 목적·항목·보유기간 안내와 동의 응답 칸(정적 텍스트)이 원문과 일치함'
    $bodyTextF29 = ($wsF29.Range('B16:F23').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF29 -notmatch '홍길동' -and $bodyTextF29 -notmatch '\d{6}-\d{7}' -and $wsF29.Range('B16').Formula -notmatch '^=') 'F-029 동의 응답·업체명·사업자번호·대표자 등 개인정보가 전부 공란이며 동의 칸은 수식이 아닌 정적 텍스트임'
    Assert-Check ([bool]$excel.Run('검증_F029출력가능')) 'F-029 선택 출력이 공통 필수값이 있을 때 허용함'
    Assert-Check ($wsF29.PageSetup.Zoom -eq 100 -and -not $wsF29.PageSetup.FitToPagesWide -and -not $wsF29.PageSetup.FitToPagesTall -and $wsF29.HPageBreaks.Count -eq 0 -and $wsF29.VPageBreaks.Count -eq 0) 'F-029가 자연 배율에서 A4 1쪽으로 출력됨'

    $wsF30 = $wb.Worksheets.Item('F-030_청렴계약이행서약서')
    Assert-Check ($wsF30.Range('B3').Value2 -eq '청렴계약 이행서약서' -and $wsF30.Range('B6').Value2 -match '지방계약법' -and $wsF30.Range('B6').Value2 -match '제6조의2') 'F-030 원문 대조 표제·도입 문구(지방계약법 제6조의2)가 존재함'
    Assert-Check ($wsF30.Range('B9').Value2 -match '^1\. ' -and $wsF30.Range('B10').Value2 -match '^2\. ' -and $wsF30.Range('B11').Value2 -match '^3\. ' -and $wsF30.Range('B12').Value2 -match '^4\. ') 'F-030 4개 준수사항이 순서대로 원문과 일치함'
    Assert-Check ($wsF30.Range('B17').Value2 -match '서약자 업체명' -and [string]::IsNullOrWhiteSpace([string]$wsF30.Range('C17').Value2)) 'F-030 서약자 업체명·대표자·서명란이 수기 작성 공란임'
    Assert-Check ([bool]$excel.Run('검증_F030출력가능')) 'F-030 선택 출력이 공통 필수값이 있을 때 허용함'
    Assert-Check ($wsF30.PageSetup.Zoom -eq 100 -and -not $wsF30.PageSetup.FitToPagesWide -and -not $wsF30.PageSetup.FitToPagesTall -and $wsF30.HPageBreaks.Count -eq 0 -and $wsF30.VPageBreaks.Count -eq 0) 'F-030이 자연 배율에서 A4 1쪽으로 출력됨'

    $wsF31 = $wb.Worksheets.Item('F-031_교복품목별금액표')
    Assert-Check ($wsF31.Range('B3').Value2 -eq '교복 품목별 금액표(품목별 1인당 계약단가)' -and $wsF31.Range('B5').Value2 -match '건  명' -and $wsF31.Range('B5').Value2 -match '2026학년도 동복 구매' -and $wsF31.Range('B7').Value2 -match '50,000원') 'F-031 표제·건명(B-01 구매명 반영)·낙찰금액(F-043 낙찰단가 50,000원 반영)이 원문과 일치함'
    Assert-Check ($wsF31.Range('B11').Value2 -eq '품목' -and $wsF31.Range('C11').Value2 -eq '수량' -and $wsF31.Range('D11').Value2 -eq '금액비율(%)' -and $wsF31.Range('E11').Value2 -eq '계약 금액(원단위 절사)') 'F-031 표 머리글이 원문과 일치함'
    Assert-Check ($wsF31.Range('B12').Value2 -eq '동복 상의' -and $wsF31.Range('C12').Value2 -eq 1 -and $wsF31.Range('D12').Text -eq '55.6%' -and $wsF31.Range('E12').Value2 -eq 27777 -and $wsF31.Range('B13').Value2 -eq '동복 하의' -and $wsF31.Range('D13').Text -eq '44.4%' -and $wsF31.Range('E13').Value2 -eq 22222) 'F-031 품목명·수량·금액비율(F-024와 동일 산식)·계약금액(F-043 낙찰단가 기준 K-01)이 골든 데이터로 정확히 계산됨'
    Assert-Check ($wsF31.Range('E18').Value2 -eq 49999) 'F-031 합계(K-02)가 개별 계약금액의 합으로 정확히 계산됨(원단위 절사로 낙찰단가 50,000원과 1원 차이는 정상)'
    Assert-Check ([string]::IsNullOrEmpty([string]$wsF31.Range('G18').Value2)) 'F-031 합계 행 비고(G18)에 내부 계산필드ID 문자열("K-02")이 인쇄되지 않고 빈 칸임(2026-09-23 버그 2 수정)'
    Assert-Check ($wsF31.Range('B25').Value2 -match '업체명' -and [string]::IsNullOrWhiteSpace([string]$wsF31.Range('C25').Value2)) 'F-031 업체명·대표자·서명란이 최종 낙찰 업체가 직접 작성하는 수기 공란임'
    Assert-Check ($wsF31.PageSetup.Zoom -eq 100 -and -not $wsF31.PageSetup.FitToPagesWide -and -not $wsF31.PageSetup.FitToPagesTall -and $wsF31.HPageBreaks.Count -eq 0 -and $wsF31.VPageBreaks.Count -eq 0) 'F-031이 자연 배율에서 A4 1쪽으로 출력됨'
    Assert-Check ([bool]$excel.Run('검증_F031출력가능')) 'F-031 선택 출력이 F-043 낙찰단가가 숫자로 있을 때 허용함'
    $wsF43.Range('C16').Value2 = ''; $excel.CalculateFullRebuild()
    Assert-Check (-not [bool]$excel.Run('검증_F031출력가능')) 'F-031 선택 출력이 F-043 낙찰단가 미입력 시 거부됨(0원 오표시 방지)'
    $wsF43.Range('C16').Value2 = 50000; $excel.CalculateFullRebuild()

    # ---- 요청 4(2026-09-23): DB 시트 5종 탭 순서 재배치 + 관리자용 열기/닫기 버튼 왕복 ----
    $dbGroupNames = @('DB', 'DB_품목', 'DB_업체', 'DB_위원', 'DB_평가')
    $idxSchoolInfo = $wb.Worksheets.Item('학교정보').Index
    $idxFirstForm = $wb.Worksheets.Item('F-001_위원회구성기안').Index
    $dbIndexes = @($dbGroupNames | ForEach-Object { $wb.Worksheets.Item($_).Index })
    Assert-Check (($dbIndexes | Measure-Object -Minimum).Minimum -gt $idxSchoolInfo -and ($dbIndexes | Measure-Object -Maximum).Maximum -lt $idxFirstForm) 'DB 시트 5종(DB·DB_품목·DB_업체·DB_위원·DB_평가) 탭 순서가 학교정보 뒤·F-001 앞에 모여 있음'
    Assert-Check (($dbIndexes | Sort-Object) -join ',' -eq ($dbIndexes -join ',')) 'DB 시트 5종이 DB→DB_품목→DB_업체→DB_위원→DB_평가 순서 그대로 인접 배치됨'
    $btnGoDbCheck = $wsIn.Buttons('btn이동_DB열기')
    Assert-Check ($btnGoDbCheck.OnAction -eq 'DB시트열기') '기초자료입력에 관리자용 DB 열기 버튼이 DB시트열기 매크로에 연결됨'
    $excel.Run('검증_DB시트열기') | Out-Null
    $wsDbCheck = $wb.Worksheets.Item('DB'); $wsItemsCheck = $wb.Worksheets.Item('DB_품목'); $wsVendorsCheck = $wb.Worksheets.Item('DB_업체'); $wsCommitteeCheck = $wb.Worksheets.Item('DB_위원'); $wsScoreCheck = $wb.Worksheets.Item('DB_평가')
    Assert-Check ($wsDbCheck.Visible -eq -1 -and $wsItemsCheck.Visible -eq -1 -and $wsVendorsCheck.Visible -eq -1 -and $wsCommitteeCheck.Visible -eq -1 -and $wsScoreCheck.Visible -eq -1) '검증_DB시트열기 실행 후 DB 시트 5종이 전부 보임(xlSheetVisible)'
    $btnDbCloseCheck = $wsDbCheck.Buttons('btnDB닫기')
    Assert-Check ($btnDbCloseCheck.OnAction -eq 'DB시트닫기') 'DB 시트에 닫기(다시 숨김) 버튼이 DB시트닫기 매크로에 연결됨'
    $excel.Run('검증_DB시트닫기') | Out-Null
    Assert-Check ($wsDbCheck.Visible -eq 0 -and $wsItemsCheck.Visible -eq 2 -and $wsVendorsCheck.Visible -eq 2 -and $wsCommitteeCheck.Visible -eq 2 -and $wsScoreCheck.Visible -eq 2) '검증_DB시트닫기 실행 후 DB=Hidden·나머지 4종=VeryHidden으로 원래 보호 상태가 복원됨'

    $wsF16 = $wb.Worksheets.Item('F-016_정성적평가')
    Assert-Check ($wsF16.Range('B3').Value2 -eq '[9-5] [붙임 3_2] [2단계] 정성적 평가' -and $wsF16.Range('B8').Value2 -eq '구분' -and $wsF16.Range('C8').Value2 -eq '평가항목' -and $wsF16.Range('F8').Value2 -eq '평가점수') 'F-016 원문 대조 제목과 고정 배점표 머리글이 존재함'
    $f016GradeValidation = $wsIn.Range('H44').Validation
    Assert-Check ($f016GradeValidation.Type -eq 3 -and $f016GradeValidation.Formula1 -match '탁월' -and $f016GradeValidation.Formula1 -match '우수' -and $f016GradeValidation.Formula1 -match '보통' -and $f016GradeValidation.Formula1 -match '미흡' -and $f016GradeValidation.Formula1 -match '불량') '기초자료입력 F-016 4개 정성평가 항목이 원문 5단계 등급 드롭다운을 제공함'
    $wsF16.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF16.Range('B6').Value2 -eq '검증업체가' -and $wsF16.Range('F9').Value2 -eq 15 -and $wsF16.Range('F13').Value2 -eq 5 -and $wsF16.Range('F14').Value2 -eq 55) 'F-016이 선택 업체 순번의 정성평가 점수와 가감점 합계를 정확히 반영함'
    $wsF16.Range('I3').Value2 = 2
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF16.Range('B6').Value2 -eq '검증업체나' -and $wsF16.Range('F9').Value2 -eq 9 -and $wsF16.Range('F14').Value2 -eq 33) 'F-016이 선택 업체 순번을 바꾸면 다른 업체의 점수를 반영함'
    $wsF16.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ([bool]$excel.Run('검증_F016출력가능')) 'F-016 선택 출력이 완전한 업체별 정성평가를 허용함'

    $wsIn.Range('C14').Value2 = '수동입력값'
    $wsMethod.Range('B5').Value2 = ''
    Invoke-ValidationMacro -macroName '검증_계약방법안내_기초자료반영'
    Assert-Check ($wsIn.Range('C14').Value2 -eq '수동입력값') '계약방법안내가 선택값이 없으면 B-03을 덮어쓰지 않음'
    $excel.Run('검증_계약방식선택', '2단계 입찰(규격·가격 동시)') | Out-Null
    Assert-Check ($wsIn.Range('C14').Value2 -eq '2단계 입찰(규격·가격 동시)' -and $wsMethod.Range('B5').Value2 -eq '2단계 입찰(규격·가격 동시)') '2단계 계약방식 버튼이 확인 질문 없이 B-03에 즉시 반영함'
    Assert-Check ($wsMethod.Range('C6').Value2 -match '규격.*가격') '계약방법안내가 2단계 입찰 절차 안내를 동적으로 반영함'
    $excel.Run('검증_계약방식선택', '일반경쟁입찰') | Out-Null
    Assert-Check ($wsIn.Range('C14').Value2 -eq '일반경쟁입찰' -and $wsMethod.Range('C6').Value2 -match '공고') '일반입찰 버튼 선택이 이전 2단계 계약방식을 덮어씀'
    $excel.Run('검증_계약방식선택', '1인견적 수의계약') | Out-Null
    Assert-Check ($wsIn.Range('C14').Value2 -eq '1인견적 수의계약' -and $wsMethod.Range('C6').Value2 -match '수의계약') '1인견적 버튼 선택이 계약방식과 안내를 즉시 전환함'
    $wsMethod.Range('B5').Value2 = ''; $wsMethod.Range('B6').Value2 = ''
    $excel.CalculateFullRebuild()

    # 계약절차안내 팝업(2026-09-21): 항상 노출되는 계약방법 3종 버튼과 frmContractGuide 폼의
    # 단계 데이터 로직을 검증함. .Show는 모달이라 절대 호출하지 않음(검증_계약절차폼이 우회).
    # 2026-09-24: 카드 4개가 도형 버튼에서 시안 이미지 위 투명 클릭 영역(Shapes)으로 바뀌어
    # Buttons() 컬렉션이 아니라 Shapes() 컬렉션에서 찾아야 함(OnAction은 도형도 동일 지원).
    Assert-Check ($wsMethod.Shapes('btn계약절차_1인견적').OnAction -eq '계약방식선택_1인견적' -and $wsMethod.Shapes('btn계약절차_2인견적').OnAction -eq '계약방식선택_2인견적' -and $wsMethod.Shapes('btn계약절차_2단계입찰').OnAction -eq '계약방식선택_2단계입찰' -and $wsMethod.Shapes('btn계약절차_일반입찰').OnAction -eq '계약방식선택_일반입찰') '계약방법안내 시안 이미지 위에 1인견적·2인견적·2단계입찰·일반입찰 투명 클릭 영역 4개가 정확히 배치됨'
    Assert-Check ($wsMethod.Shapes('imgContractBanner').Type -eq 13) '계약방법안내 1행에 시안 배너 이미지가 삽입됨(13=msoPicture)'
    $guideFormComponent = $wb.VBProject.VBComponents.Item('frmContractGuide')
    $guideFormCode = $guideFormComponent.CodeModule.Lines(1, $guideFormComponent.CodeModule.CountOfLines)
    Assert-Check ($guideFormComponent.Designer.Controls.Item('btnHWPX').Caption -eq 'HWPX 작성' -and $guideFormCode -match 'Private Sub btnHWPX_Click\(\)') '계약절차 팝업의 관련 서식 영역에 HWPX 작성 버튼과 클릭 처리가 존재함'
    # 2026-09-24 재디자인(시안-5-*.png): 계약유형별 테마색 배너·2열 서식 표가 실제로 적용되는지 헤드리스로 확인
    Assert-Check ($guideFormComponent.Designer.Controls.Item('lstForms').ColumnCount -eq 2) '계약절차 팝업 관련 서식 목록이 코드·서식명 2열 표로 구성됨'
    Assert-Check ($guideFormCode -match 'Private Sub ApplyTheme\(ByVal contractType As String\)' -and $guideFormCode -match 'Me\.fraHeader\.BackColor = themeColor') '계약절차 팝업이 계약유형별 테마색을 상단 배너에 적용함'
    $bidResult = [string]$excel.Run('검증_계약절차폼', '입찰계약')
    Assert-Check ($bidResult -match '^11\|' -and $bidResult -match '입찰공고문') '입찰계약 절차 팝업이 11단계로 구성되고 1단계 설명에 입찰공고문 관련 내용이 포함됨'
    $twoQuoteResult = [string]$excel.Run('검증_계약절차폼', '2인견적 수의계약')
    Assert-Check ($twoQuoteResult -match '^5\|') '2인견적 수의계약 절차 팝업이 5단계로 구성됨'
    $oneQuoteResult = [string]$excel.Run('검증_계약절차폼', '1인견적 수의계약')
    Assert-Check ($oneQuoteResult -match '^4\|') '1인견적 수의계약 절차 팝업이 4단계로 구성됨'
    $bidStepsWithForms = [string]$excel.Run('검증_계약절차폼', '입찰계약')
    Assert-Check ($bidStepsWithForms -match '\|(\d+)$' -and [int]$Matches[1] -eq 5) '입찰계약 1단계(입찰공고문 작성)의 관련 서식이 5건(F-009·010·011·008·012) 모두 목록에 채워짐'

    try {
        Invoke-ValidationMacro -macroName '검증_저장하기'
        L "매크로 저장하기(): 정상 실행"
    } catch {
        $failures.Add("매크로 저장하기() 실패: $($_.Exception.Message)")
        L "FAIL: 매크로 저장하기() 실패: $($_.Exception.Message)"
    }
    # 2026-09-23 버그 수정 검증: 저장하기()가 DB만 쓰고 파일은 저장하지 않아(기존 버그로 추정) 사용자가
    # "저장되었습니다" 메시지를 믿고 파일을 닫으면 데이터가 사라지던 문제를 고쳤는지 확인함.
    Assert-Check ($wb.Saved) '저장하기()가 DB 기록뿐 아니라 실제 파일 저장까지 수행함(2026-09-23 버그 수정)'

    $wsDB = $wb.Worksheets.Item("DB")
    $dbRow2A = $wsDB.Cells.Item(2,1).Value2
    $dbRow2B = $wsDB.Cells.Item(2,2).Value2
    L "DB 시트 2행 기록 확인: 순번=$dbRow2A, 학교명=$dbRow2B"
    Assert-Check ($dbRow2A -eq 1 -and $dbRow2B -eq '테스트초등학교') '저장하기()가 DB에 마스킹 테스트값을 기록함'
    $wsItems = $wb.Worksheets.Item('DB_품목')
    L "DB_품목 기록 확인: $($wsItems.Cells.Item(2, 1).Value2)/$($wsItems.Cells.Item(2, 3).Value2), $($wsItems.Cells.Item(3, 1).Value2)/$($wsItems.Cells.Item(3, 3).Value2)"
    Assert-Check ($wsItems.Cells.Item(2, 1).Value2 -eq 1 -and $wsItems.Cells.Item(2, 3).Value2 -eq '동복 상의' -and $wsItems.Cells.Item(3, 1).Value2 -eq 1 -and $wsItems.Cells.Item(3, 3).Value2 -eq '동복 하의') '저장하기()가 복수 품목 반복행을 DB_품목에 기록함'
    $wsVendors = $wb.Worksheets.Item('DB_업체')
    L "DB_업체 기록 확인: $($wsVendors.Cells.Item(2, 1).Value2)/$($wsVendors.Cells.Item(2, 3).Value2), $($wsVendors.Cells.Item(3, 1).Value2)/$($wsVendors.Cells.Item(3, 3).Value2)"
    Assert-Check ($wsVendors.Cells.Item(2, 1).Value2 -eq 1 -and $wsVendors.Cells.Item(2, 2).Value2 -eq 1 -and $wsVendors.Cells.Item(2, 3).Value2 -eq '검증업체가' -and $wsVendors.Cells.Item(3, 1).Value2 -eq 1 -and $wsVendors.Cells.Item(3, 2).Value2 -eq 2 -and $wsVendors.Cells.Item(3, 3).Value2 -eq '검증업체나') '저장하기()가 복수 업체 반복행을 DB_업체에 기록함'
    $wsQuant = $wb.Worksheets.Item('DB_정량평가')
    Assert-Check ($wsQuant.Cells.Item(2, 1).Value2 -eq 1 -and $wsQuant.Cells.Item(2, 2).Value2 -eq 1 -and $wsQuant.Cells.Item(2, 3).Value2 -eq 10 -and $wsQuant.Cells.Item(2, 4).Value2 -eq 10 -and $wsQuant.Cells.Item(2, 5).Value2 -eq 15 -and $wsQuant.Cells.Item(2, 6).Value2 -eq 15 -and $wsQuant.Cells.Item(3, 2).Value2 -eq 2 -and $wsQuant.Cells.Item(3, 3).Value2 -eq 6) '저장하기()가 업체별 F-015 정량평가 점수를 DB_정량평가에 정규화 저장함(Workbook_Open 재보호 목록에 DB_정량평가 포함되어 무한 대기 재발 없음)'
    $wsQual = $wb.Worksheets.Item('DB_정성평가')
    Assert-Check ($wsQual.Cells.Item(2, 1).Value2 -eq 1 -and $wsQual.Cells.Item(2, 2).Value2 -eq 1 -and $wsQual.Cells.Item(2, 3).Value2 -eq 15 -and $wsQual.Cells.Item(2, 4).Value2 -eq 10 -and $wsQual.Cells.Item(2, 5).Value2 -eq 15 -and $wsQual.Cells.Item(2, 6).Value2 -eq 10 -and $wsQual.Cells.Item(2, 7).Value2 -eq 5 -and $wsQual.Cells.Item(3, 2).Value2 -eq 2 -and $wsQual.Cells.Item(3, 7).Value2 -eq -2) '저장하기()가 업체별 F-016 정성평가 점수와 가감점을 DB_정성평가에 정규화 저장함'
    $wsSelf = $wb.Worksheets.Item('DB_자기평점')
    Assert-Check ($wsSelf.Cells.Item(2, 1).Value2 -eq 1 -and $wsSelf.Cells.Item(2, 2).Value2 -eq 1 -and $wsSelf.Cells.Item(2, 3).Value2 -eq 8 -and $wsSelf.Cells.Item(2, 4).Value2 -eq 8 -and $wsSelf.Cells.Item(2, 5).Value2 -eq 12 -and $wsSelf.Cells.Item(2, 6).Value2 -eq 15 -and $wsSelf.Cells.Item(3, 2).Value2 -eq 2 -and $wsSelf.Cells.Item(3, 3).Value2 -eq 10) '저장하기()가 F-018 업체 자기평점을 DB_자기평점에 학교평가와 분리하여 정규화 저장함'
    $wsCommittee = $wb.Worksheets.Item('DB_위원')
    Assert-Check ($wsCommittee.Cells.Item(2, 1).Value2 -eq 1 -and $wsCommittee.Cells.Item(2, 2).Value2 -eq 1 -and $wsCommittee.Cells.Item(2, 3).Value2 -eq '교원위원' -and $wsCommittee.Cells.Item(2, 4).Value2 -eq '위원 ○○' -and $wsCommittee.Cells.Item(3, 3).Value2 -eq '학부모위원' -and $wsCommittee.Cells.Item(3, 4).Value2 -eq '위원 **') '저장하기()가 복수 위원 역할·마스킹 식별표시를 DB_위원에 정규화 저장함'
    Assert-Check ($wsDB.Cells.Item(2, 22).Value2 -eq '박운영' -and $wsDB.Cells.Item(2, 23).Value2 -eq '김교장') '저장하기()가 C-08 담당자 성명·C-07 교장 성명을 DB V·W열(22·23)에 저장함(2026-09-23 신설, U열 불러오기표시라벨과 충돌 없음)'

    # 2026-09-24 사용자 요청: [불러오기] 버튼이 드롭다운 사전 선택 없이 바로 팝업(frmLoadPicker)으로
    # 저장된 레코드 목록을 보여주는지 헤드리스로 검증함(.Show는 호출하지 않음).
    $loadPickerCount = [string]$excel.Run('검증_불러오기팝업')
    Assert-Check ($loadPickerCount -eq '1') '불러오기 팝업이 현재까지 저장된 레코드 1건을 목록에 채움'
    Assert-Check ($wsDB.Cells.Item(2, 21).Value2 -match '^순번 1: 테스트초등학교') 'DB U열(21) 불러오기표시라벨이 신규 담당자성명·교장성명 열에 의해 덮어써지지 않음'
    Assert-Check ($wsVendors.Cells.Item(2, 4).Value2 -eq '박대표' -and $wsVendors.Cells.Item(2, 5).Value2 -eq '02-000-0000') '저장하기()가 R-08 대표자 성명·R-09 대표자 연락처를 DB_업체 4·5열에 저장함(2026-09-23 신설)'
    $wsScore = $wb.Worksheets.Item('DB_평가')
    Assert-Check ($wsScore.Cells.Item(2, 1).Value2 -eq 1 -and $wsScore.Cells.Item(2, 2).Value2 -eq 1 -and $wsScore.Cells.Item(2, 3).Value2 -eq '제품 품질' -and $wsScore.Cells.Item(2, 4).Value2 -eq 40 -and $wsScore.Cells.Item(2, 5).Value2 -eq 38.5 -and $wsScore.Cells.Item(3, 5).Value2 -eq 60 -and $wsIn.Range('D82').Value2 -eq 98.5) '저장하기()가 복수 평가점수와 K-03 총점을 DB_평가에 정규화 저장함'
    Assert-Check ([bool]$excel.Run('검증_위원행검증') -and [bool]$excel.Run('검증_평가행검증')) '위원 마스킹·평가점수 정상값과 경계값(점수=배점)을 허용함'

    # 6.0 저장 경계: 무암호 보호를 해제한 직접 DB 편집도 Workbook_BeforeSave에서 재검증해 영구 저장을 취소한다.
    foreach ($internalSheet in @($wsDB, $wsItems, $wsVendors, $wsCommittee, $wsScore, $wsQuant, $wsQual, $wsSelf)) { $internalSheet.Unprotect("") }
    Assert-Check ([bool]$excel.Run('검증_저장경계')) '정상 DB·반복행은 저장 경계 재검증을 통과함'
    $wb.Save()
    Assert-Check ($wb.Saved) '정상 DB·반복행은 Workbook_BeforeSave 경계를 거쳐 저장됨'

    # 2026-09-24 사용자 결정: 완성도(개인정보 보호와 무관한) 문제는 더 이상 저장을 막지
    # 않는다. 검증_저장경계()는 여전히 문제를 감지해 False를 반환하지만(경고 표시용),
    # Workbook_BeforeSave는 이제 이를 이유로 Cancel하지 않으므로 $wb.Saved는 True가 된다.
    $wsCommittee.Cells.Item(2, 4).Value2 = '홍'
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_위원 R-02 성명이 2자 미만인 직접 입력을 저장 경계가 감지함(2026-09-23 마스킹 폐지, 길이 규칙만 적용)'
    $wb.Saved = $false; $wb.Save()
    Assert-Check ($wb.Saved) '완성도 문제(개인정보 무관)는 경고만 하고 저장은 계속 진행됨(2026-09-24)'
    $wsCommittee.Cells.Item(2, 4).Value2 = '위원 ○○'

    $wsCommittee.Cells.Item(2, 3).ClearContents()
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_위원 역할·식별표시 불완전 직접 입력을 저장 경계가 감지함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check ($wb.Saved) 'DB_위원 역할 누락 직접 입력도 경고만 하고 저장은 계속 진행됨(2026-09-24)'
    $wsCommittee.Cells.Item(2, 3).Value2 = '교원위원'

    $wsScore.Cells.Item(2, 5).Value2 = 40.01
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_평가 배점 초과 점수 직접 입력을 저장 경계가 감지함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check ($wb.Saved) 'DB_평가 배점 초과 직접 입력도 경고만 하고 저장은 계속 진행됨(2026-09-24)'
    $wsScore.Cells.Item(2, 5).Value2 = 38.5

    $wsQual.Cells.Item(2, 7).Value2 = 6
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_정성평가 가감점 범위 초과 직접 입력을 저장 경계가 감지함'
    $wsQual.Cells.Item(2, 7).Value2 = 5

    $wsVendors.Cells.Item(2, 3).Value2 = '010-1234-5678'
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_업체 연락처성 직접 입력을 저장 경계가 감지함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check ($wb.Saved) 'DB_업체 연락처성 직접 입력도 경고만 하고 저장은 계속 진행됨(2026-09-24)'
    $wsVendors.Cells.Item(2, 3).Value2 = '검증업체가'

    $wsDB.Cells.Item(3, 1).Value2 = 2
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB 불완전 주 레코드 직접 입력을 저장 경계가 감지함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check ($wb.Saved) 'DB 불완전 주 레코드 직접 입력도 경고만 하고 저장은 계속 진행됨(2026-09-24)'
    $wsDB.Range('A3:T3').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_저장경계')) '직접 입력 오류를 원복하면 저장 경계가 정상 데이터만 허용함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check ($wb.Saved) '직접 입력 오류를 원복한 정상 데이터는 Workbook_BeforeSave를 거쳐 저장됨'

    $wsIn.Range('B60').Value2 = '외부위원'
    $wsIn.Range('C60').Value2 = '홍'
    Assert-Check (-not [bool]$excel.Run('검증_위원행검증')) '위원 성명이 2자 미만이면 저장을 거부함(2026-09-23 마스킹 폐지, 길이 규칙만 적용)'
    $wsIn.Range('C60').ClearContents()
    $wsIn.Range('B60').Value2 = '외부위원'
    Assert-Check (-not [bool]$excel.Run('검증_위원행검증')) '위원 역할만 있는 불완전 반복행을 거부함'
    $wsIn.Range('B60:C60').ClearContents()
    $wsIn.Range('C58').Value2 = '홍길동'
    Assert-Check ([bool]$excel.Run('검증_위원행검증')) '위원 성명 실명 형태(2026-09-23 마스킹 폐지, 정상값)를 허용함'
    $wsIn.Range('C58').Value2 = '홍길동' + [char]10 + '위원'
    Assert-Check (-not [bool]$excel.Run('검증_위원행검증')) '위원 성명에 줄바꿈이 있으면 검증 함수는 여전히 문제로 감지함(2026-09-24부터 경고 용도로만 사용)'
    Invoke-ValidationMacro -macroName '검증_수정하기'
    Assert-Check ($wsCommittee.Cells.Item(2, 4).Value2 -match [char]10) '2026-09-24 사용자 결정: 형식 문제도 경고만 하고 수정 저장이 그대로 진행됨(줄바꿈 포함 값도 DB에 저장됨)'
    $wsIn.Range('C58').Value2 = '위원 ○○'
    Invoke-ValidationMacro -macroName '검증_수정하기'
    Assert-Check ($wsCommittee.Cells.Item(2, 4).Value2 -eq '위원 ○○') '오염된 DB 값을 정상값으로 재저장해 이후 검증에 영향이 없도록 복구함'

    # R-08/R-09(대표자 성명·연락처) 라이브 검증(2026-09-23 신설)
    $wsIn.Range('S46').Value2 = '최대표'
    Assert-Check (-not [bool]$excel.Run('검증_업체대표자행검증')) '업체명 없이 대표자 성명만 입력한 행을 거부함'
    $wsIn.Range('B46').Value2 = '검증업체다'
    Assert-Check ([bool]$excel.Run('검증_업체대표자행검증')) '업체명과 대표자 성명이 함께 있으면 허용함'
    $wsIn.Range('X46').Value2 = '전화번호아님'
    Assert-Check (-not [bool]$excel.Run('검증_업체대표자행검증')) '대표자 연락처가 숫자·하이픈 형식이 아니면 거부함'
    $wsIn.Range('X46').Value2 = '02-111-2222'
    Assert-Check ([bool]$excel.Run('검증_업체대표자행검증')) '대표자 연락처가 숫자·하이픈 형식이면 허용함'
    $wsIn.Range('B46:F46,S46,X46').ClearContents()

    $wsIn.Range('D74').Value2 = 1
    Assert-Check (-not [bool]$excel.Run('검증_평가행검증')) '평가항목·배점 없는 점수 고아 입력행을 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F014출력가능')) 'F-014 선택 출력이 불완전 평가행을 차단함'
    $wsIn.Range('B74:D74').ClearContents()
    $wsIn.Range('B74').Value2 = '가격 경쟁력'
    $wsIn.Range('C74').Value2 = 20
    $wsIn.Range('D74').Value2 = 20.01
    Assert-Check (-not [bool]$excel.Run('검증_평가행검증')) '배점을 초과한 평가점수를 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F014출력가능')) 'F-014 선택 출력이 배점 초과 평가행을 차단함'
    $wsIn.Range('B74:D74').ClearContents()

    $wsIn.Range('C44').Value2 = 11
    Assert-Check (-not [bool]$excel.Run('검증_정량평가행검증')) 'F-015 배점 상한(수행경험 10점)을 초과한 점수를 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F015출력가능')) 'F-015 선택 출력이 배점 초과 정량평가를 차단함'
    $wsIn.Range('C44').Value2 = 10
    $wsIn.Range('D46').Value2 = 5
    Assert-Check (-not [bool]$excel.Run('검증_정량평가행검증')) '업체명 없이 정량평가 점수만 있는 F-015 고아 입력을 거부함'
    $wsIn.Range('D46').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_정량평가행검증')) 'F-015 정량평가 정상 복귀 후 저장 검증을 통과함'
    $wsIn.Range('N44').Value2 = 11
    Assert-Check (-not [bool]$excel.Run('검증_자기평점행검증')) 'F-018 자기 수행경험 배점 상한(10점)을 초과한 점수를 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F018출력가능')) 'F-018 선택 출력이 배점 초과 자기평점을 차단함'
    $wsIn.Range('N44').Value2 = 8
    $wsIn.Range('O46').Value2 = 5
    Assert-Check (-not [bool]$excel.Run('검증_자기평점행검증')) '업체명 없이 자기평점만 있는 F-018 고아 입력을 거부함'
    $wsIn.Range('O46').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_자기평점행검증')) 'F-018 자기평점 정상 복귀 후 저장 검증을 통과함'
    $wsIn.Range('H44').Value2 = '나일론'
    Assert-Check (-not [bool]$excel.Run('검증_정성평가행검증')) 'F-016 재질에 원문 등급 외 값(나일론)을 입력하면 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F016출력가능')) 'F-016 선택 출력이 등급 외 재질 입력을 차단함'
    $wsIn.Range('H44').Value2 = '탁월'
    $wsIn.Range('L46').Value2 = 0
    Assert-Check (-not [bool]$excel.Run('검증_정성평가행검증')) '업체명 없이 정성평가 점수만 있는 F-016 고아 입력을 거부함'
    $wsIn.Range('L46').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_정성평가행검증')) 'F-016 정성평가 정상 복귀 후 저장 검증을 통과함'
    $wsIn.Range('B58:C67').ClearContents()
    $wsIn.Range('B72:D81').ClearContents()
    $wsIn.Range('B44:B53').ClearContents()
    $wsIn.Range('H44:L53').ClearContents()
    $wsIn.Range('N44:Q53').ClearContents()
    Invoke-ValidationMacro -macroName '검증_첫레코드불러오기'
    Assert-Check ($wsIn.Range('B44').Value2 -eq '검증업체가' -and $wsIn.Range('B45').Value2 -eq '검증업체나') '불러오기()가 업체 반복행을 원래 행에 복원함'
    Assert-Check ($wsIn.Range('C44').Value2 -eq 10 -and $wsIn.Range('D44').Value2 -eq 10 -and $wsIn.Range('E44').Value2 -eq 15 -and $wsIn.Range('F44').Value2 -eq 15 -and $wsIn.Range('C45').Value2 -eq 6) '불러오기()가 F-015 업체별 정량평가 점수를 원래 행에 복원함'
    Assert-Check ($wsIn.Range('H44').Value2 -eq '탁월' -and $wsIn.Range('I44').Value2 -eq '탁월' -and $wsIn.Range('J44').Value2 -eq '탁월' -and $wsIn.Range('K44').Value2 -eq '탁월' -and $wsIn.Range('H45').Value2 -eq '보통' -and $wsIn.Range('J45').Value2 -eq '우수' -and $wsIn.Range('L44').Value2 -eq 5 -and $wsIn.Range('L45').Value2 -eq -2) '불러오기()가 F-016 정성평가 등급과 가감점을 원래 행에 복원함'
    Assert-Check ($wsIn.Range('N44').Value2 -eq 8 -and $wsIn.Range('O44').Value2 -eq 8 -and $wsIn.Range('P44').Value2 -eq 12 -and $wsIn.Range('Q44').Value2 -eq 15 -and $wsIn.Range('N45').Value2 -eq 10 -and $wsIn.Range('Q45').Value2 -eq 0) '불러오기()가 F-018 업체별 자기평점을 학교평가와 분리하여 원래 행에 복원함'

    # 불러오기 드롭다운(2026-09-21): DB!U열 라벨 생성과 드롭다운 선택값 파싱을 분리 검증함.
    # 실제 [불러오기] Sub는 형식 오류 시 MsgBox를 띄우므로 헤드리스로 직접 호출하지 않고,
    # MsgBox 없는 파싱 함수(검증_불러오기_드롭다운파싱)만 호출함.
    Assert-Check ($wsDB.Cells.Item(2, 21).Value2 -eq '순번 1: 테스트초등학교 / 2026학년도 / 2026학년도 동복 구매') 'DB!U2 불러오기 표시 라벨이 순번·학교명·학년도·구매명으로 정확히 구성됨'
    Assert-Check ([long]$excel.Run('검증_불러오기_드롭다운파싱', $wsDB.Cells.Item(2, 21).Value2) -eq 1) '불러오기 드롭다운 파싱이 U2 라벨에서 순번 1을 정확히 추출함'
    Assert-Check ([long]$excel.Run('검증_불러오기_드롭다운파싱', '이상한 형식') -eq 0) '불러오기 드롭다운 파싱이 형식이 다른 문자열을 0으로 거부함'
    Assert-Check ([long]$excel.Run('검증_불러오기_드롭다운파싱', '') -eq 0) '불러오기 드롭다운 파싱이 빈 문자열을 0으로 거부함'
    $wsInI2Validation = $wsIn.Range('I2').Validation
    Assert-Check ($wsInI2Validation.Type -eq 3 -and $wsInI2Validation.Formula1 -match 'U\$2') '기초자료입력!I2가 DB!U열 범위를 참조하는 목록 형식 데이터 유효성 검사임'

    $wsF39 = $wb.Worksheets.Item('F-039_업체별제안서평가표')
    $wsF39.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF39.Range('C8').Value2 -eq '검증업체가' -and -not [bool]$excel.Run('검증_F039출력가능')) 'F-039가 I3 선택 업체명을 반영하고, 위원 8명 점수 미입력 시 선택 출력을 차단함'
    $scoresF39 = @([double]15,[double]10,[double]15,[double]10, [double]12,[double]8,[double]12,[double]8, [double]15,[double]10,[double]15,[double]10, [double]3,[double]2,[double]3,[double]2, [double]9,[double]6,[double]9,[double]6, [double]12,[double]8,[double]12,[double]8, [double]15,[double]10,[double]15,[double]10, [double]6,[double]4,[double]6,[double]4)
    # 2026-09-19 발견: 이 시트에서 .Range('E8').Value2 = <숫자>처럼 "주소 문자열로 얻은
    # Range에 숫자를 대입"하면 PowerShell-COM 상호운용 계층에서 간헐적으로
    # "InvalidCastException: Int32/Double을 String으로 캐스팅할 수 없습니다" 오류가 재현됨
    # (원인 불명, 격리된 재현 스크립트로는 재현되지 않고 이 검증 스크립트의 누적 상태에서만
    # 발생함). .Cells.Item(행,열).Value2 = <숫자> 형태는 매번 안정적으로 동작해 이 패턴으로
    # 통일함(이 파일의 다른 대부분 코드가 이미 이 방식을 씀).
    for ($rowIdx = 0; $rowIdx -lt 8; $rowIdx++) {
        $rr = 8 + $rowIdx
        $wsF39.Cells.Item($rr, 5).Value2 = [double]$scoresF39[$rowIdx * 4]
        $wsF39.Cells.Item($rr, 6).Value2 = [double]$scoresF39[$rowIdx * 4 + 1]
        $wsF39.Cells.Item($rr, 7).Value2 = [double]$scoresF39[$rowIdx * 4 + 2]
        $wsF39.Cells.Item($rr, 8).Value2 = [double]$scoresF39[$rowIdx * 4 + 3]
    }
    $excel.CalculateFullRebuild()
    Assert-Check ([bool]$excel.Run('검증_F039출력가능')) 'F-039가 위원 8명의 4개 항목 점수를 모두 입력하면 선택 출력을 허용함'
    Assert-Check ($wsF39.Range('I8').Value2 -eq 50) 'F-039 행별 점수합계(옷감15+완성도10+A/S15+하자10=50)가 정확히 계산됨'
    Assert-Check ($wsF39.Range('E16').Value2 -eq 69 -and $wsF39.Range('E17').Value2 -eq 11.5) 'F-039 옷감 항목 점수(15,12,15,3,9,12,15,6)의 총계가 최고점(15)·최저점(3) 제외 6명 합(69)이고 평균은 총계/6(11.5)임'
    $wsF39.Range('D9').Value2 = '홍길동'
    $bodyTextF39 = ($wsF39.Range('B7:J17').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF39 -notmatch '자동') 'F-039 위원명(D열)은 이 시트 어떤 수식으로도 자동 채워지지 않고 담당자가 직접 입력한 값만 남음(청탁방지·익명성 원칙)'
    $wsF39.Range('D9').ClearContents()
    $wsF39.Range('E8').ClearContents()
    Assert-Check (-not [bool]$excel.Run('검증_F039출력가능')) 'F-039가 위원 점수 1칸만 비어도 선택 출력을 차단함'
    $wsF39.Cells.Item(8, 5).Value2 = [double]$scoresF39[0]

    Assert-Check ($wsIn.Range('B58').Value2 -eq '교원위원' -and $wsIn.Range('C58').Value2 -eq '위원 ○○' -and $wsIn.Range('B59').Value2 -eq '학부모위원' -and $wsIn.Range('D82').Value2 -eq 98.5) '불러오기()가 위원·평가 반복행과 K-03 총점을 복원함'
    Assert-Check ($wsIn.Range('C10').Value2 -eq '박운영' -and $wsIn.Range('F10').Value2 -eq '김교장') '불러오기()가 C-08 담당자 성명·C-07 교장 성명을 복원함(2026-09-23 신설)'
    Assert-Check ($wsIn.Range('S44').Value2 -eq '박대표' -and $wsIn.Range('X44').Value2 -eq '02-000-0000') '불러오기()가 R-08 대표자 성명·R-09 대표자 연락처를 업체 반복행에 복원함(2026-09-23 신설)'
    $wsIn.Range('C59').Value2 = '위원 ○*'
    $wsIn.Range('D73').Value2 = 59.5
    $wsIn.Range('B60').Value2 = '교원위원'
    $wsIn.Range('C60').Value2 = '위원 ○○'
    Assert-Check ([bool]$excel.Run('검증_위원중복경고')) '수정 전 위원 역할 중복 경고를 감지함'
    $wsIn.Range('B60:C60').ClearContents()
    Invoke-ValidationMacro -macroName '검증_수정하기'
    $wsIn.Range('B58:C67').ClearContents()
    $wsIn.Range('B72:D81').ClearContents()
    Invoke-ValidationMacro -macroName '검증_첫레코드불러오기'
    Assert-Check ($wsIn.Range('C59').Value2 -eq '위원 ○*' -and $wsIn.Range('D73').Value2 -eq 59.5 -and $wsScore.Cells.Item(3, 5).Value2 -eq 59.5 -and $wsIn.Range('D82').Value2 -eq 98) '수정하기()가 위원·평가 반복행을 교체하고 총점을 재계산함'
    # 업체 2의 학교 정량·정성 평가와 F-018 자기평점을 함께 비워야 업체명 없는 고아행이 남지 않는다.
    $wsIn.Range('B45:F45').ClearContents()
    $wsIn.Range('H45:L45').ClearContents()
    $wsIn.Range('N45:Q45').ClearContents()
    $wsIn.Range('B47').Value2 = '검증업체다'
    Invoke-ValidationMacro -macroName '검증_수정하기'
    $wsIn.Range('B44:B53').ClearContents()
    Invoke-ValidationMacro -macroName '검증_첫레코드불러오기'
    Assert-Check ($wsIn.Range('B44').Value2 -eq '검증업체가' -and [string]::IsNullOrWhiteSpace([string]$wsIn.Range('B45').Value2) -and $wsIn.Range('B47').Value2 -eq '검증업체다' -and $wsVendors.Cells.Item(3, 2).Value2 -eq 4) '수정하기()가 기존 업체 반복행을 교체하고 저장 행번호를 보존함'
    $wsIn.Range('B46').Value2 = '검증업체가'
    Assert-Check ([bool]$excel.Run('검증_업체행검증') -and [bool]$excel.Run('검증_업체중복경고')) '중복 업체명은 저장 전 경고 신호를 반환함'
    # 2026-09-24 사용자 결정(재확인): 형식·완성도 문제는 검증 함수 자체는 여전히 감지해
    # 경고를 표시하지만, 저장 자체는 항상 진행된다. 아래는 그 정책이 실제로 적용되는지
    # 확인하고, 매 반복마다 DB 3행을 비워 다음 시나리오가 항상 같은 새 행(seq=2)에서
    # 시작하도록 정리한다(자식 테이블의 부수 기록은 이 1회성 검증 사본에만 남고 downstream
    # 검사가 그 행 수를 참조하지 않음을 확인했다).
    $wsIn.Range('B46').Value2 = 'X'
    Assert-Check (-not [bool]$excel.Run('검증_업체행검증')) '1자 업체명 입력을 검증 함수가 여전히 감지함(경고 용도로 유지)'
    Invoke-ValidationMacro -macroName '검증_저장하기'
    Assert-Check (-not [string]::IsNullOrWhiteSpace([string]$wsDB.Cells.Item(3, 1).Value2)) '2026-09-24 사용자 결정: 형식 오류 업체명도 경고만 하고 저장이 진행됨'
    $wsDB.Range('A3:T3').ClearContents()
    $wsIn.Range('B46').Value2 = ('가' * 151)
    Assert-Check (-not [bool]$excel.Run('검증_업체행검증')) '151자 이상 업체명 입력을 검증 함수가 여전히 감지함(경고 용도로 유지)'
    $wsIn.Range('B46').ClearContents()
    foreach ($forbiddenVendorValue in @('010-1234-5678', '010.1234.5678', '123-45-67890', '123 45 67890', '대표자 홍길동')) {
        $wsIn.Range('B46').Value2 = $forbiddenVendorValue
        Assert-Check (-not [bool]$excel.Run('검증_업체행검증')) "업체명 입력란이 금지된 식별·연락처 표기($forbiddenVendorValue)를 검증 함수가 여전히 감지함(경고 용도로 유지)"
        Invoke-ValidationMacro -macroName '검증_저장하기'
        Assert-Check (-not [string]::IsNullOrWhiteSpace([string]$wsDB.Cells.Item(3, 1).Value2)) "2026-09-24 사용자 결정: 금지된 식별·연락처 표기($forbiddenVendorValue)도 경고만 하고 저장이 진행됨"
        $wsDB.Range('A3:T3').ClearContents()
    }
    $wsIn.Range('B46').ClearContents()
    $wsIn.Range('B31:D31').ClearContents()
    $wsIn.Range('C31').Value2 = 1
    Assert-Check (-not [bool]$excel.Run('검증_품목행검증')) '품목명 없는 수량·단가 고아 입력행을 검증 함수가 여전히 감지함(경고 용도로 유지)'
    $wsIn.Range('B31:D31').ClearContents()
    $wsIn.Range('B31').Value2 = '불완전 품목'
    Assert-Check (-not [bool]$excel.Run('검증_품목행검증')) '수정 경로와 공유하는 품목 검증도 불완전 행을 여전히 감지함(경고 용도로 유지)'
    Invoke-ValidationMacro -macroName '검증_저장하기'
    Assert-Check (-not [string]::IsNullOrWhiteSpace([string]$wsDB.Cells.Item(3, 1).Value2)) '2026-09-24 사용자 결정: 불완전 품목 행도 경고만 하고 저장이 진행됨'
    $wsDB.Range('A3:T3').ClearContents()
    $wsIn.Range('B31:D31').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_품목행검증') -and $wsIn.Application.WorksheetFunction.CountA($wsIn.Range('B29:B38')) -ge 1 -and $wsIn.Application.WorksheetFunction.CountA($wsIn.Range('B29:B38')) -le 6) 'F-024 출력용 품목 입력이 완전하고 1~6개임'
    Assert-Check ([bool]$excel.Run('검증_F024출력가능') -eq $true) 'F-024 선택 출력이 정상 품목 1~6개를 허용함'
    for ($r = 31; $r -le 35; $r++) {
        $wsIn.Cells.Item($r, 2).Value2 = "추가 품목 $r"
        $wsIn.Cells.Item($r, 3).Value2 = 1
        $wsIn.Cells.Item($r, 4).Value2 = 1000
    }
    Assert-Check ([bool]$excel.Run('검증_F024출력가능') -eq $false) 'F-024 선택 출력이 완전한 7개 품목을 거부함'
    $wsIn.Range('B31:D35').ClearContents()

    # F-007/F-014/F-024 수식이 HWPX 원문 대조 후 확정한 배치에서 기초자료입력을 정상 참조하는지 확인
    $wsF7 = $wb.Worksheets.Item("F-007_구매요청기안문")
    L "F-007 제목 셀(C9) 계산값: $($wsF7.Range('C9').Value2)"
    Assert-Check ($wsF7.Range('C9').Value2 -eq '2026학년도 동복 학교주관구매 요청') 'F-007 공통 제목이 입력값을 참조함'
    $wsF24 = $wb.Worksheets.Item("F-024_단가비율표")
    L "F-024 수량 합계(D15) 계산값: $($wsF24.Range('D15').Value2), 비율(E9/E10): $($wsF24.Range('E9').Value2)/$($wsF24.Range('E10').Value2)"
    Assert-Check ($wsF24.Range('D15').Value2 -eq 204 -and $wsF24.Range('E9').Text -eq '55.6%' -and $wsF24.Range('E10').Text -eq '44.4%') 'F-024 수량 합계와 단가비율이 입력값을 참조함'
    Assert-Check ($wsF24.Range('C9').Value2 -eq '동복 상의' -and $wsF24.Range('C10').Value2 -eq '동복 하의') 'F-024 품목명이 입력 반복행을 참조함'
    $wsF17 = $wb.Worksheets.Item('F-017_제출서류자기확인서')
    $wsF17.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF17.Range('B8').Value2 -eq '검증업체가' -and $wsF17.Range('C8').Value2 -eq '박대표' -and $wsF17.Range('D8').Value2 -eq '02-000-0000' -and [string]::IsNullOrWhiteSpace([string]$wsF17.Range('E8').Value2) -and $wsF17.Range('B10').Value2 -eq '제출서류') 'F-017은 업체명·대표자·연락처(R-08·R-09, 2026-09-23부터 자동 반영)를 반영하고 대리인 정보는 빈 양식으로 유지함'
    Assert-Check ([bool]$excel.Run('검증_F017출력가능')) 'F-017 선택 출력이 공통 필수값과 업체 반복행을 확인함'
    $longTitle = ('평가 기준 & 특수문자 <>()[]{} / ' * 4)
    $wsIn.Range('C4').Value2 = ('검증학교 & 특수문자 <>()[]{} ' * 3)
    $wsIn.Range('C22').Value2 = $longTitle
    # DB_정량평가 연동 제거로 F-015 점수는 불러오기() 이후 유지되지 않으므로(위 DB_정량평가 주석 참고),
    # 수식 검증을 위해 여기서 다시 직접 입력함.
    $wsIn.Range('B44').Value2 = '검증업체가'
    $wsIn.Range('C44').Value2 = 10
    $wsIn.Range('D44').Value2 = 10
    $wsIn.Range('E44').Value2 = 15
    $wsIn.Range('F44').Value2 = 15
    $wsF15.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF14.Range('C5').Value2 -eq $wsIn.Range('C4').Value2 -and $wsF14.Range('B6').Value2 -eq $longTitle -and $wsF14.Range('F20').Value2 -eq 98) 'F-014 최대 길이·특수문자 공통값과 점수 합계가 수식 오류 없이 반영됨'
    Assert-Check ($wsF15.Range('C5').Value2 -eq $wsIn.Range('C4').Value2 -and $wsF15.Range('F22').Value2 -eq 50) 'F-015 최대 길이·특수문자 학교명과 정량평가 합계가 수식 오류 없이 반영됨(같은 세션 직접 입력값 기준)'
    $wsIn.Range('C4').Value2 = '테스트초등학교'
    $wsIn.Range('C22').Value2 = '2026학년도 동복 학교주관구매 요청'
    $excel.CalculateFullRebuild()

    # 6a. 학교명·지역·급별 검색 및 선택 결과 C4 반영 (공개 기관정보 표본, 검증 사본만 사용)
    # 원본 표본의 외부 쿼리 캐시에 검색 열 값이 비어 있을 수 있으므로, 검증 사본에만 마스킹 기관정보를 넣어
    # 학교명·지역·급별 조건의 복합 검색과 C4 반영 흐름을 독립적으로 재현함.
    $wsSchool.Cells.Item(2, 2).Value2 = '검증지역'
    $wsSchool.Cells.Item(2, 3).Value2 = '초등'
    $wsSchool.Cells.Item(2, 5).Value2 = '검증학교'
    $wsSearch.Range('B3').Value2 = '검증학교'
    $wsSearch.Range('D3').Value2 = '검증지역'
    $wsSearch.Range('F3').Value2 = '초등'
    Invoke-ValidationMacro -macroName '검증_학교검색_실행'
    Assert-Check ($wsSearch.Cells.Item(7, 4).Value2 -eq '검증학교') '학교명·지역·급별 복합 검색 결과가 생성됨'
    # 2026-09-24부터 학교검색·입찰계약절차안내는 화면에서 숨김(Visible=0) 처리되어,
    # 숨긴 상태로 Activate()한 뒤 Select()를 호출하면 COM 예외가 발생함(실측 확인).
    # 검증 중에만 임시로 보이게 했다가 원래 숨김 상태로 되돌린다.
    $wsSearch.Visible = -1
    $wsSearch.Activate()
    $wsSearch.Cells.Item(7, 1).Select()
    Invoke-ValidationMacro -macroName '검증_선택학교_기초자료반영'
    Assert-Check ($wsIn.Range('C4').Value2 -eq '검증학교') '선택 학교명이 기초자료입력!C4에 반영됨'
    $wsSearch.Visible = 0

    # 6b. 9단계 절차 안내에서 단계별 대표 Form ID로 이동
    $wsFlow.Visible = -1
    $wsFlow.Activate()
    $wsFlow.Cells.Item(7, 1).Select()  # 3단계: F-007
    Invoke-ValidationMacro -macroName '검증_관련FormID로이동'
    Assert-Check ($excel.ActiveSheet.Name -eq '서식선택_출력' -and $excel.ActiveCell.Value2 -eq 'F-007') '3단계 절차안내가 F-007 선택 행으로 이동함'
    $wsFlow.Visible = 0

    # 7. 서식선택_출력 체크 + PDF 내보내기 매크로 테스트
    $wsSel = $wb.Worksheets.Item("서식선택_출력")
    Assert-Check ($wsSel.Columns.Item(1).Hidden -eq $true) '서식선택_출력의 내부 선택값 A열이 사용자 화면에서 숨김 처리됨'
    $legacyTopButtons = @('btn인쇄미리보기', 'btnPDF저장', 'btnHWPX작성') | Where-Object {
        try { $null = $wsSel.Buttons($_); $true } catch { $false }
    }
    Assert-Check ($legacyTopButtons.Count -eq 0) '서식선택_출력의 중복 상단 일괄 출력 버튼 3개가 제거됨'
    # F-007, F-014, F-024 행 찾기
    $lastRow = $wsSel.Cells.Item($wsSel.Rows.Count, 2).End(-4162).Row  # xlUp
    for ($r = 5; $r -le $lastRow; $r++) {
        $fid = $wsSel.Cells.Item($r, 2).Value2
        if ($fid -eq "F-001" -or $fid -eq "F-002" -or $fid -eq "F-003" -or $fid -eq "F-004" -or $fid -eq "F-005" -or $fid -eq "F-006" -or $fid -eq "F-007" -or $fid -eq "F-008" -or $fid -eq "F-009" -or $fid -eq "F-010" -or $fid -eq "F-011" -or $fid -eq "F-012" -or $fid -eq "F-014" -or $fid -eq "F-015" -or $fid -eq "F-016" -or $fid -eq "F-017" -or $fid -eq "F-018" -or $fid -eq "F-019" -or $fid -eq "F-020" -or $fid -eq "F-021" -or $fid -eq "F-022" -or $fid -eq "F-023" -or $fid -eq "F-024" -or $fid -eq "F-025" -or $fid -eq "F-032" -or $fid -eq "F-033" -or $fid -eq "F-034" -or $fid -eq "F-035" -or $fid -eq "F-036" -or $fid -eq "F-037" -or $fid -eq "F-038" -or $fid -eq "F-039" -or $fid -eq "F-040" -or $fid -eq "F-041" -or $fid -eq "F-042" -or $fid -eq "F-043" -or $fid -eq "F-044" -or $fid -eq "F-045" -or $fid -eq "F-046" -or $fid -eq "F-047" -or $fid -eq "F-048" -or $fid -eq "F-049" -or $fid -eq "F-050" -or $fid -eq "F-027" -or $fid -eq "F-029" -or $fid -eq "F-030" -or $fid -eq "F-031") {
            $wsSel.Cells.Item($r, 1).Value2 = $true
        }
    }
    $pdfStartTime = Get-Date
    # Err.Raise가 Application.Run 경계를 넘으면 예측 불가능하게 실패하는 것을 반복 확인해
    # (2026-09-21) Boolean 반환으로 바꿈. 실패 시 pdf_error_log.txt에서 실제 원인을 읽어 보고함.
    $pdfBatchResult = $false
    try {
        $pdfBatchResult = [bool](Invoke-ValidationMacro -macroName '검증_선택서식_PDF저장')
    } catch {
        L "FAIL: 매크로 선택서식_PDF저장() 호출 자체가 예외로 실패함: $($_.Exception.Message)"
    }
    if ($pdfBatchResult) {
        L "매크로 선택서식_PDF저장(): 정상 실행"
    } else {
        $errLogPath = Join-Path $validationDirectory 'pdf_error_log.txt'
        $errDetail = if (Test-Path $errLogPath) { (Get-Content $errLogPath -Raw -Encoding UTF8).Trim() } else { '(pdf_error_log.txt 없음)' }
        $failures.Add("매크로 선택서식_PDF저장() 실패: $errDetail")
        L "FAIL: 매크로 선택서식_PDF저장() 실패: $errDetail"
    }
    $leftoverF015TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F015_임시_*' })
    Assert-Check ($leftoverF015TempSheets.Count -eq 0) 'F-015 업체별 임시 인쇄 시트가 PDF 저장 성공·실패와 관계없이 정리됨'
    $leftoverF016TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F016_임시_*' })
    Assert-Check ($leftoverF016TempSheets.Count -eq 0) 'F-016 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF017TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F017_임시_*' })
    Assert-Check ($leftoverF017TempSheets.Count -eq 0) 'F-017 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF018TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F018_임시_*' })
    Assert-Check ($leftoverF018TempSheets.Count -eq 0) 'F-018 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF019TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F019_임시_*' })
    Assert-Check ($leftoverF019TempSheets.Count -eq 0) 'F-019 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF020TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F020_임시_*' })
    Assert-Check ($leftoverF020TempSheets.Count -eq 0) 'F-020 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF021TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F021_임시_*' })
    Assert-Check ($leftoverF021TempSheets.Count -eq 0) 'F-021 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF022TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F022_임시_*' })
    Assert-Check ($leftoverF022TempSheets.Count -eq 0) 'F-022 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF023TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F023_임시_*' })
    Assert-Check ($leftoverF023TempSheets.Count -eq 0) 'F-023 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF025TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F025_임시_*' })
    Assert-Check ($leftoverF025TempSheets.Count -eq 0) 'F-025 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'

    $outputDir = Join-Path $validationDirectory 'output'
    if (Test-Path $outputDir) {
        $pdfs = Get-ChildItem $outputDir -Filter "*.pdf" | Where-Object { $_.LastWriteTime -ge $pdfStartTime } | Sort-Object LastWriteTime -Descending
        if ($pdfs.Count -gt 0) {
            L "PDF 생성 확인: $($pdfs[0].FullName) ($([math]::Round($pdfs[0].Length/1024,1)) KB)"
            Assert-Check ($pdfs[0].Length -gt 0) '선택 서식 PDF가 비어 있지 않음'
        } else {
            $failures.Add('PDF 생성 실패: output 폴더에 PDF 없음')
            L 'FAIL: PDF 생성 실패: output 폴더에 PDF 없음'
        }
    } else {
        $failures.Add('PDF 생성 실패: output 폴더 자체가 없음')
        L 'FAIL: PDF 생성 실패: output 폴더 자체가 없음'
    }

    # 7b. 행별 미리보기/PDF 버튼: 구현상태=Y인 모든 행에 버튼이 있고, 개별 실행이 기존
    # 일괄 선택 상태를 건드리지 않으며, 반복행 서식(F-017)의 임시 시트도 정리되는지 확인
    $implementedRowCount = 0
    $missingButtonRows = @()
    for ($r = 5; $r -le $lastRow; $r++) {
        if ($wsSel.Cells.Item($r, 5).Value2 -eq "Y") {
            $implementedRowCount++
            $fid = $wsSel.Cells.Item($r, 2).Value2
            $hasPrevBtn = $false
            $hasPdfBtn = $false
            $hasHwpxBtn = $false
            try { $null = $wsSel.Buttons("btnPrev_$fid"); $hasPrevBtn = $true } catch { $hasPrevBtn = $false }
            try { $null = $wsSel.Buttons("btnPdf_$fid"); $hasPdfBtn = $true } catch { $hasPdfBtn = $false }
            try { $null = $wsSel.Buttons("btnHwpx_$fid"); $hasHwpxBtn = $true } catch { $hasHwpxBtn = $false }
            if (-not $hasPrevBtn -or -not $hasPdfBtn -or -not $hasHwpxBtn) { $missingButtonRows += $fid }
        }
    }
    Assert-Check ($missingButtonRows.Count -eq 0) "구현상태=Y인 $implementedRowCount 개 행 전부에 행별 미리보기/PDF/HWPX 버튼이 존재함(HWPX 열은 2026-09-24 신설)"
    # HWPX 버튼은 실제 클릭 시 npx kordoc을 외부 프로세스로 호출해 완료까지 대기하므로(2026-09-24
    # 동기 실행으로 변경), 빠른 COM 전용 자동 검증 루프에서는 배치·배선(OnAction)만 확인하고
    # 실제 생성 실행은 하지 않는다(사람이 버튼을 눌러 확인하는 범위로 남김).
    Assert-Check ($wsSel.Buttons('btnHwpx_F-007').OnAction -eq '행별_HWPX작성') 'F-007 행별 HWPX 버튼이 행별_HWPX작성 매크로에 연결됨'
    $outputModuleCode = $wb.VBProject.VBComponents.Item('Module_출력').CodeModule.Lines(1, $wb.VBProject.VBComponents.Item('Module_출력').CodeModule.CountOfLines)
    Assert-Check ($outputModuleCode -match 'ThisWorkbook\.FollowHyperlink outputPath' -and $outputModuleCode -match '작성할 서식을 체크하세요') 'HWPX 작성은 선택한 서식만 생성하고, 단일 결과는 작성 후 직접 열도록 구현됨'

    $wsGuide = $wb.Worksheets.Item('사용설명서')
    Assert-Check ($wsGuide.Shapes('imgGuideBanner').Type -eq 13 -and $wsGuide.Shapes('hotGuideMethod').OnAction -eq '이동_계약방법안내' -and $wsGuide.Shapes('hotGuideInput').OnAction -eq '이동_기초자료입력' -and $wsGuide.Shapes('hotGuideOutput').OnAction -eq '이동_서식선택출력') '사용설명서 시안 이미지와 3개 단계 클릭 영역이 실제 이동 동작에 연결됨'

    # 기초자료입력 "빠른 출력 선택" 패널(2026-09-21): 체크박스 LinkedCell 양방향 연동과
    # 보기 버튼 존재 여부를 검증함. 버튼 클릭 체인(Application.Caller)은 행별_실행/폼ID별_실행이
    # 이미 검증됐으므로 여기서는 배치·연동 정합성만 확인함.
    Assert-Check ($wsIn.Buttons('btn빠른미리보기').OnAction -eq '선택서식_인쇄미리보기' -and $wsIn.Buttons('btn빠른PDF').OnAction -eq '선택서식_PDF저장') '기초자료입력의 빠른 출력 선택 패널에 기존 일괄 미리보기/PDF 버튼과 동일한 매크로가 연결됨'
    $quickCheckboxCount = 0
    $quickViewButtonCount = 0
    $quickLinkMismatch = @()
    for ($r = 5; $r -le $lastRow; $r++) {
        $qFid = $wsSel.Cells.Item($r, 2).Value2
        try {
            $chkObj = $wsIn.CheckBoxes("chkQuick_$qFid")
            $quickCheckboxCount++
            if ($chkObj.LinkedCell -notmatch "\`$A\`$$r$") { $quickLinkMismatch += $qFid }
        } catch { $quickLinkMismatch += $qFid }
        if ($wsSel.Cells.Item($r, 5).Value2 -eq "Y") {
            try { $null = $wsIn.Buttons("btnQuickView_$qFid"); $quickViewButtonCount++ } catch { $null = $_ }
        }
    }
    Assert-Check ($quickCheckboxCount -eq ($lastRow - 4)) "빠른 출력 선택 패널에 서식선택_출력과 동일한 $($lastRow - 4)개 체크박스가 배치됨"
    Assert-Check ($quickLinkMismatch.Count -eq 0) '빠른 출력 선택 패널의 모든 체크박스가 서식선택_출력!A열의 같은 행에 LinkedCell로 연결됨'
    Assert-Check ($quickViewButtonCount -eq $implementedRowCount) "빠른 출력 선택 패널의 보기 버튼이 구현상태=Y인 $implementedRowCount 개 행과 정확히 일치함"

    # 양방향 연동 실측: 서식선택_출력!A5를 TRUE로 바꾸면 기초자료입력 체크박스도 즉시 체크됨을 확인
    $wsSel.Cells.Item(5, 1).Value2 = $true
    $excel.CalculateFullRebuild()
    $firstFormId = $wsSel.Cells.Item(5, 2).Value2
    $firstChk = $wsIn.CheckBoxes("chkQuick_$firstFormId")
    Assert-Check ($firstChk.Value -eq 1) '서식선택_출력에서 체크하면 기초자료입력의 연결된 체크박스도 즉시 체크됨(양방향 연동)'
    $wsSel.Cells.Item(5, 1).Value2 = $false
    $excel.CalculateFullRebuild()
    # CheckBox.Value는 Boolean 0/1이 아니라 xlOn(1)/xlOff(-4146)/xlMixed(2) 상수를 반환함
    # (2026-09-21 확인, 이전에 0을 기대한 것이 테스트 자체의 오류였음 — 기능은 정상 동작).
    $firstChkRecheck = $wsIn.CheckBoxes("chkQuick_$firstFormId")
    Assert-Check ($firstChkRecheck.Value -eq -4146) '서식선택_출력에서 체크 해제하면 기초자료입력의 연결된 체크박스도 즉시 해제됨(xlOff)'
    Assert-Check ($wsSel.Buttons('btnPrev_F-007').OnAction -eq '행별_미리보기') 'F-007 행별 미리보기 버튼이 사용자용 행별_미리보기 매크로에 연결됨'

    $checkSnapshotBefore = @{}
    for ($r = 5; $r -le $lastRow; $r++) { $checkSnapshotBefore[$r] = $wsSel.Cells.Item($r, 1).Value2 }

    $rowPdfStartTime = Get-Date
    $rowResult1 = $excel.Run('검증_행별출력', 'F-007', $true)
    Assert-Check ([bool]$rowResult1) '매크로 검증_행별출력(F-007, PDF)이 정상 실행됨'
    # 업체별 F-017은 위 통합 PDF 검사에서 실제 임시 시트 생성·삭제를 이미 확인한다.
    # 같은 Excel 인스턴스에서 행별 호출을 한 번 더 하면 COM이 종료되는 현상이 있어,
    # 행별 출력 경로는 임시 시트를 만들지 않는 F-007 단일 PDF로 검증한다.

    $checkStateRestored = $true
    for ($r = 5; $r -le $lastRow; $r++) {
        if ($wsSel.Cells.Item($r, 1).Value2 -ne $checkSnapshotBefore[$r]) { $checkStateRestored = $false }
    }
    Assert-Check $checkStateRestored '행별 출력 실행 후 다른 행의 선택(TRUE/FALSE) 상태가 그대로 보존됨'

    $rowPdfs = @(Get-ChildItem $outputDir -Filter "*.pdf" | Where-Object { $_.LastWriteTime -ge $rowPdfStartTime })
    Assert-Check ($rowPdfs.Count -ge 1) '행별 PDF 출력이 F-007 단일 파일로 생성됨'

    $rowResultBlocked = $excel.Run('검증_행별출력', 'F-013', $true)
    Assert-Check (-not [bool]$rowResultBlocked) '미구현 서식(F-013)의 행별 출력이 의도대로 거부됨'

    # 8. 매크로 실행 후 테스트값 원복(초기화) — 배포본에 테스트 데이터가 남지 않도록
    Invoke-ValidationMacro -macroName '검증_초기화'
    $excel.CalculateFullRebuild()
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF17.Range('B8').Value2)) 'F-017 빈 배포본은 업체명 미선택 시 0 대신 빈 칸을 표시함'
    $wsDB2 = $wb.Worksheets.Item("DB")
    foreach ($internalSheet in @($wsDB2, $wb.Worksheets.Item('DB_품목'), $wb.Worksheets.Item('DB_업체'), $wb.Worksheets.Item('DB_위원'), $wb.Worksheets.Item('DB_평가'), $wb.Worksheets.Item('DB_정량평가'), $wb.Worksheets.Item('DB_정성평가'), $wb.Worksheets.Item('DB_자기평점'))) {
        $internalSheet.Unprotect("")
    }
    $wsDB2.Range("A2:T2").ClearContents()
    $wb.Worksheets.Item('DB_품목').Range('A2:F100').ClearContents()
    $wb.Worksheets.Item('DB_업체').Range('A2:C100').ClearContents()
    $wb.Worksheets.Item('DB_위원').Range('A2:D100').ClearContents()
    $wb.Worksheets.Item('DB_평가').Range('A2:E100').ClearContents()
    $wb.Worksheets.Item('DB_정량평가').Range('A2:F100').ClearContents()
    $wb.Worksheets.Item('DB_정성평가').Range('A2:G100').ClearContents()
    $wb.Worksheets.Item('DB_자기평점').Range('A2:F100').ClearContents()
    $missingInputAccepted = [bool]$excel.Run('검증_필수값검증')
    Assert-Check (-not $missingInputAccepted) '필수값이 비어 있으면 저장 검증이 거부됨'
    L "테스트 데이터 정리(기초자료입력 초기화, DB 2행 삭제) 완료"

    L '테스트 데이터 정리 완료(검증 사본만 변경, 배포본 저장 없음)'

} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($excelProcessId) {
        Start-Sleep -Milliseconds 500
        if (Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue) {
            Stop-Process -Id $excelProcessId -Force -ErrorAction SilentlyContinue
            L "PID $excelProcessId Excel 프로세스가 Quit() 이후에도 남아 있어 강제 종료함"
        }
    }
    RemovePathWithRetry -path $validationDirectory -maxAttempts 10 -delayMs 500
}

if ($failures.Count -gt 0) {
    throw ('Excel v1 검증 실패: ' + ($failures -join '; '))
}

L 'PASS: Excel v1 검증 전체 통과'
