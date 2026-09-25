# QA 전용: 버그 2(F-014/F-031/F-036 표 칸 정리·수식 수정)와 요청 4(DB 시트 이동+버튼)의
# 실제 렌더링/화면 동작을 육안 확인하기 위한 1회성 스크립트. 원본 v11.xlsm은 건드리지 않고
# 사본에서만 값을 채우고 PDF·스크린샷을 뽑는다. 마스킹 값만 사용함.
$ErrorActionPreference = 'Stop'
$root = 'C:\Users\최익창\Downloads\나연수_26.9.10'
$srcXlsm = Join-Path $root 'artifacts\excel\교복구매_길라잡이_20260923_v11.xlsm'
$qaDir = Join-Path $root 'artifacts\excel\_diagnostic\bug2_req4_qa'
New-Item -ItemType Directory -Path $qaDir -Force | Out-Null
$qaXlsm = Join-Path $qaDir 'qa_copy.xlsm'
Copy-Item -LiteralPath $srcXlsm -Destination $qaXlsm -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($qaXlsm, [Type]::Missing, $false)
    $wsIn = $wb.Worksheets.Item('기초자료입력')

    # 마스킹 골든 값 채우기(개인정보 아님, 프로젝트 관례값)
    $wsIn.Range('C4').Value2 = '검증초등학교'
    $wsIn.Range('C5').Value2 = 2026
    $wsIn.Range('C12').Value2 = '2026학년도 동복 구매'
    $wsIn.Range('B44').Value2 = '검증업체가'
    $wsIn.Range('H44').Value2 = '탁월'
    $wsIn.Range('I44').Value2 = '탁월'
    $wsIn.Range('J44').Value2 = '탁월'
    $wsIn.Range('K44').Value2 = '탁월'
    $wsIn.Range('L44').Value2 = 5
    $wsIn.Range('C44').Value2 = 10
    $wsIn.Range('D44').Value2 = 10
    $wsIn.Range('E44').Value2 = 15
    $wsIn.Range('F44').Value2 = 15
    $excel.CalculateFullRebuild()

    $wsF14 = $wb.Worksheets.Item('F-014_평가항목배점기준')
    $wsF31 = $wb.Worksheets.Item('F-031_교복품목별금액표')
    $wsF36 = $wb.Worksheets.Item('F-036_제안서평가결과')

    Write-Output "F-014 G20 (버그2 정리 확인): '$($wsF14.Range('G20').Value2)'"
    Write-Output "F-031 G18 (버그2 정리 확인): '$($wsF31.Range('G18').Value2)'"
    Write-Output "F-036 D19 (정성 합계, 기대 55): $($wsF36.Range('D19').Value2)"
    Write-Output "F-036 E19 (총점, 기대 105): $($wsF36.Range('E19').Value2)"
    Write-Output "F-036 F19 (판정, 기대 적격): $($wsF36.Range('F19').Value2)"

    foreach ($item in @(
        @{ ws = $wsF14; name = 'F-014' },
        @{ ws = $wsF31; name = 'F-031' },
        @{ ws = $wsF36; name = 'F-036' }
    )) {
        $pdfPath = Join-Path $qaDir ("{0}_QA.pdf" -f $item.name)
        $item.ws.ExportAsFixedFormat(0, $pdfPath)
        Write-Output "PDF 저장: $pdfPath"
    }

    # 요청 4: DB 열기 -> 화면 캡처용 상태 확인 -> 버튼 좌표/겹침 육안 확인을 위해 DB 시트를
    # 스크린샷 대신 각 셀의 값·버튼 좌표를 텍스트로 남긴다(레이아웃 겹침 여부는 좌표로 판단).
    $excel.Run('검증_DB시트열기') | Out-Null
    $wsDb = $wb.Worksheets.Item('DB')
    Write-Output "DB.Visible after open: $($wsDb.Visible)"
    $btn = $wsDb.Buttons('btnDB닫기')
    Write-Output "DB 닫기 버튼 좌표: Left=$($btn.Left) Top=$($btn.Top) Width=$($btn.Width) Height=$($btn.Height)"
    Write-Output "DB 헤더 W1(23열) 오른쪽 끝 대략 위치와 버튼 Left 비교로 겹침 여부 판단: 헤더 마지막 열(W=23) Right=$($wsDb.Columns.Item(23).Left + $wsDb.Columns.Item(23).Width), 버튼 Left=$($btn.Left)"
    $wsInBtn = $wsIn.Buttons('btn이동_DB열기')
    Write-Output "기초자료입력 DB 열기 버튼 좌표: Left=$($wsInBtn.Left) Top=$($wsInBtn.Top)"
    $wsInBtnPrev = $wsIn.Buttons('btn이동_계약방법안내')
    Write-Output "기초자료입력 계약방법안내 버튼(바로 왼쪽) 좌표: Left=$($wsInBtnPrev.Left) Right=$($wsInBtnPrev.Left + $wsInBtnPrev.Width)"
    $excel.Run('검증_DB시트닫기') | Out-Null
    Write-Output "DB.Visible after close: $($wsDb.Visible)"

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Write-Output 'QA 스크립트 완료'
