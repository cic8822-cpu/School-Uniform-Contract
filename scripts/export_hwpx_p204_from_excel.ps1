<#
.SYNOPSIS
  Excel 기초자료입력 시트의 허용 공통값을 읽어 P2-04의 47종(+F-013 보너스)
  HWPX 템플릿 전체에 명시 토큰 채움을 일괄 적용한다("③ Excel→HWPX 연동 스크립트").

.DESCRIPTION
  계획.md 8.2절에 명시된 설계("47개 서식 각각의 기초자료입력/DB_* 현재 값을
  JSON으로 모아 해당 서식의 템플릿에 넘겨 실제 HWPX를 생성")를 구현한다.
  `scripts/export_hwpx_from_excel.ps1`(P2-01/F-007 전용, 셀 6개만 읽는 단일
  서식 CLI)을 계승하되, P2-04 표준 산출물 구조(`artifacts/hwpx/P2-04/
  F-0XX_<설명>_템플릿.hwpx`)에 맞춰 폴더 전체를 스캔해 일반화했다.

  허용 입력은 기초자료입력 시트의 6개 셀뿐이다:
    C4(C-01 학교명) · C5(C-02 학년도) · C8(C-05 문서 발행일) ·
    C9(C-06 문서번호) · C18(B-07 제출기한) · C23(D-03 관련문서)
  그 외 담당자·위원·업체·개인 성명·연락처·서명·직인 등은 어떤 셀도 읽지 않는다
  (입력데이터_사전.md "보호" 행, `hwpx_token_fill.ps1`의 $AllowedTokens·
  $ForbiddenTokenPattern이 이중으로 강제함). Excel VBA를 실행·수정하지 않는다
  (읽기 전용 COM Open, UpdateLinks=0, AutomationSecurity=3, 저장 없음).

  각 템플릿이 실제로 포함한 `{{토큰}}` 종류는 파일마다 다르므로(예: F-002는
  {{학교명}}만, F-001은 {{학교명}}·{{학년도}}·{{문서번호}}·{{발행일}}), 템플릿을
  직접 스캔해 그 서식이 필요로 하는 토큰만 골라 최소 JSON을 만든다. 필요한 값이
  비어 있으면(아직 Excel에 입력 안 됨) 그 서식만 건너뛰고 사유를 기록하며,
  전체 배치를 중단하지 않는다.

  실제 채움은 검증된 `hwpx_token_fill.ps1`(원본 재직렬화 없이 hp:t 텍스트만
  바이트 치환, 허용 토큰 화이트리스트·금지 패턴 이중 검사, 원자적 파일 교체)을
  그대로 재사용한다 — 이 스크립트는 "Excel 값 수집 → 서식별 JSON 조립 →
  hwpx_token_fill.ps1 호출 → kordoc validate" 오케스트레이션만 담당한다.

.PARAMETER ExcelPath
  읽을 XLSM 경로(원본을 포함해 어떤 XLSM도 절대 저장하지 않으므로 원본 직접
  지정도 안전하나, 관례상 배포 후보 사본을 권장).

.PARAMETER TemplateDir
  P2-04 템플릿이 있는 디렉터리. 기본값 artifacts/hwpx/P2-04.

.PARAMETER OutputDir
  생성 결과를 저장할 디렉터리. 기본값 artifacts/hwpx/P2-04_생성 — 품질검증용
  템플릿/골든결과/원본대조 3종 세트(artifacts/hwpx/P2-04)와 섞이지 않도록 별도
  폴더에 둔다.

.PARAMETER FormIds
  선택: 처리할 Form ID 배열(예: 'F-007','F-014'). 생략하면 TemplateDir의 모든
  *_템플릿.hwpx를 처리한다.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts/export_hwpx_p204_from_excel.ps1 `
    -ExcelPath artifacts/excel/교복구매_길라잡이_20260921_v9.xlsm
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ExcelPath,
    [string]$TemplateDir = 'artifacts/hwpx/P2-04',
    [string]$OutputDir = 'artifacts/hwpx/P2-04_생성',
    [string[]]$FormIds
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$KordocVersion = '4.13.1'

function Get-CellText($Sheet, [string]$Address) { return (($Sheet.Range($Address).Text -as [string]).Trim()) }
function Read-ZipEntryText($Entry) { $reader=[IO.StreamReader]::new($Entry.Open(),[Text.UTF8Encoding]::new($false,$true),$true); try{$reader.ReadToEnd()}finally{$reader.Dispose()} }
function Invoke-Kordoc([string[]]$KordocArguments) { & npx -y "kordoc@$KordocVersion" @KordocArguments; if($LASTEXITCODE -ne 0){throw "kordoc $KordocVersion 실패(exit $LASTEXITCODE): $($KordocArguments -join ' ')"} }
function Get-SharedFileHash([string]$Path) {
    # Get-FileHash는 FileShare.Read로 파일을 여는데, Excel은 통합문서를 열 때
    # FileAccess.ReadWrite로 잠가 다른 프로세스가 쓰기 공유까지 허용해야 열 수 있게
    # 한다. HWPX 작성 버튼은 지금 열려 있는(사용 중인) 바로 그 xlsm을 대상으로 이
    # 해시를 구하므로, 표준 Get-FileHash를 쓰면 "다른 프로세스에서 사용 중이므로
    # 액세스할 수 없습니다"(종료 코드 1)로 항상 실패한다(2026-09-24 실측, 스크린샷
    # 오류 재현). FileShare.ReadWrite로 직접 스트림을 열어 계산하면 통합문서가 열려
    # 있어도 안전하게 읽기 전용으로 해시를 구할 수 있다.
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    try {
        $sha256=[Security.Cryptography.SHA256]::Create()
        try { return ([BitConverter]::ToString($sha256.ComputeHash($stream)) -replace '-','') }
        finally { $sha256.Dispose() }
    } finally { $stream.Dispose() }
}

function Get-TemplateTokens([string]$HwpxPath) {
    # 템플릿의 Contents/section*.xml만 열어 실제로 쓰인 {{토큰}} 이름 집합을 반환한다.
    # 원본을 수정하지 않는 읽기 전용 접근이다.
    $zip = [IO.Compression.ZipFile]::OpenRead($HwpxPath)
    try {
        $tokens = [Collections.Generic.HashSet[string]]::new()
        foreach ($entry in $zip.Entries | Where-Object { $_.FullName -match '(^|/)Contents/section\d+\.xml$' }) {
            $xml = Read-ZipEntryText $entry
            foreach ($m in [regex]::Matches($xml, '<hp:t(?:\s[^>]*)?>(?<text>.*?)</hp:t>', [Text.RegularExpressions.RegexOptions]::Singleline)) {
                foreach ($tokenMatch in [regex]::Matches($m.Groups['text'].Value, '\{\{([^{}]+)\}\}')) {
                    [void]$tokens.Add($tokenMatch.Groups[1].Value)
                }
            }
        }
        return @($tokens)
    } finally { $zip.Dispose() }
}

$excelFullPath = (Resolve-Path -LiteralPath $ExcelPath).Path
$excelHash = Get-SharedFileHash $excelFullPath
$templateDirFull = (Resolve-Path -LiteralPath $TemplateDir).Path
[IO.Directory]::CreateDirectory($OutputDir) | Out-Null
$outputDirFull = (Resolve-Path -LiteralPath $OutputDir).Path

# ---- 1. Excel에서 허용 6개 셀만 읽는다(읽기 전용, 저장 없음, VBA 미실행) ----
$excel=$null; $workbook=$null; $sheet=$null; $schoolDirectory=$null
$values = [ordered]@{}
$schoolAddress = ''; $schoolPhone = ''
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($excelFullPath, 0, $true)  # UpdateLinks=0, ReadOnly=$true
    $sheet = $workbook.Worksheets.Item('기초자료입력')

    $schoolName = Get-CellText $sheet 'C4'
    $schoolYearRaw = Get-CellText $sheet 'C5'
    $issueDate = Get-CellText $sheet 'C8'
    $documentNumber = Get-CellText $sheet 'C9'
    $submitDeadline = Get-CellText $sheet 'C18'
    $relatedDocument = Get-CellText $sheet 'C23'
    # 2026-09-24 사용자 요청: HWPX 하단 결재란의 담당자·교장 성명도 자동반영한다.
    # 기초자료입력!C10(B10 라벨 "담당자 성명")·F10(E10 라벨 "교장 성명(결재권자)")는
    # 2026-09-23 결정(CLAUDE.md)으로 이미 자동반영 허용 범위에 포함된 값이다.
    $managerName = Get-CellText $sheet 'C10'
    $principalName = Get-CellText $sheet 'F10'
    foreach ($v in @($managerName, $principalName)) {
        if ($v -match '[\r\n\x00-\x1F]') { throw "허용값에 줄바꿈/제어 문자가 있습니다: $v" }
    }
    $schoolDirectory = $workbook.Worksheets.Item('학교정보')
    $schoolLastRow = $schoolDirectory.Cells($schoolDirectory.Rows.Count, 5).End(-4162).Row
    for ($schoolRow = 2; $schoolRow -le $schoolLastRow; $schoolRow++) {
        if ((Get-CellText $schoolDirectory ("E" + $schoolRow)) -eq $schoolName) {
            $schoolAddress = Get-CellText $schoolDirectory ("F" + $schoolRow)
            $schoolPhone = Get-CellText $schoolDirectory ("G" + $schoolRow)
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($schoolName)) { throw '필수 허용값이 비어 있습니다: 학교명(C4)' }
    if ([string]::IsNullOrWhiteSpace($schoolYearRaw)) { throw '필수 허용값이 비어 있습니다: 학년도(C5)' }
    foreach ($v in @($schoolName, $schoolYearRaw, $issueDate, $documentNumber, $submitDeadline, $relatedDocument)) {
        if ($v -match '[\r\n\x00-\x1F]') { throw "허용값에 줄바꿈/제어 문자가 있습니다: $v" }
    }
    if ($schoolYearRaw -notmatch '^\d{4}(학년도)?$') { throw "학년도(C-02)는 4자리 연도여야 합니다: $schoolYearRaw" }
    $schoolYear = if ($schoolYearRaw.EndsWith('학년도')) { $schoolYearRaw } else { $schoolYearRaw + '학년도' }

    $values['학교명'] = $schoolName
    $values['학년도'] = $schoolYear
    if (-not [string]::IsNullOrWhiteSpace($issueDate)) { $values['발행일'] = $issueDate }
    if (-not [string]::IsNullOrWhiteSpace($documentNumber)) { $values['문서번호'] = $documentNumber }
    if (-not [string]::IsNullOrWhiteSpace($submitDeadline)) { $values['제출기한'] = $submitDeadline }
    if (-not [string]::IsNullOrWhiteSpace($relatedDocument)) { $values['관련문서'] = $relatedDocument }
} finally {
    if ($schoolDirectory) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($schoolDirectory) | Out-Null }
    if ($sheet) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) | Out-Null }
    if ($workbook) { $workbook.Close($false); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ((Get-SharedFileHash $excelFullPath) -ne $excelHash) { throw '입력 XLSM 해시가 변경되었습니다(읽기 전용 위반 의심).' }

function Format-ReadValue([string]$Key) {
    if ($values.Contains($Key)) { return $values[$Key] }
    return '비어있음'
}
$documentNumberDisplay = Format-ReadValue '문서번호'
$relatedDocumentDisplay = Format-ReadValue '관련문서'
$issueDateDisplay = Format-ReadValue '발행일'
$submitDeadlineDisplay = Format-ReadValue '제출기한'
Write-Output "읽은 허용값: 학교명=$($values['학교명']) 학년도=$($values['학년도']) 문서번호=$documentNumberDisplay 관련문서=$relatedDocumentDisplay 발행일=$issueDateDisplay 제출기한=$submitDeadlineDisplay 학교주소=$schoolAddress 학교전화=$schoolPhone 담당자성명=$managerName 교장성명=$principalName"

# ---- 2. 템플릿 목록을 확정한다 ----
$templates = @(Get-ChildItem -LiteralPath $templateDirFull -Filter '*_템플릿.hwpx' | Sort-Object Name)
if ($FormIds) {
    $templates = @($templates | Where-Object { $formId = ($_.Name -split '_')[0]; $formId -in $FormIds })
}
if ($templates.Count -eq 0) { throw "처리할 템플릿을 찾지 못했습니다: $templateDirFull" }
Write-Output "대상 템플릿 $($templates.Count)개 확인 -> $templateDirFull"

# ---- 3. 서식별로 필요한 토큰만 골라 hwpx_token_fill.ps1을 호출한다 ----
$succeeded = @(); $skipped = @(); $failed = @()
$jsonWorkDir = Join-Path $outputDirFull '_json_tmp'
[IO.Directory]::CreateDirectory($jsonWorkDir) | Out-Null

foreach ($tpl in $templates) {
    $formId = ($tpl.Name -split '_')[0]
    try {
        # @()로 감싸 0개/1개 토큰일 때 PowerShell이 $null/스칼라로 축약하는 것을 방지한다.
        $requiredTokens = @(Get-TemplateTokens $tpl.FullName)
        if ($requiredTokens.Count -eq 0) {
            # 토큰이 전혀 없는 서식(F-014 등)은 템플릿을 그대로 복사만 한다.
            $outName = $tpl.Name -replace '_템플릿\.hwpx$', '_생성본.hwpx'
            $outPath = Join-Path $outputDirFull $outName
            Copy-Item -LiteralPath $tpl.FullName -Destination $outPath -Force
            & (Join-Path $PSScriptRoot 'hwpx_school_footer_fill.ps1') -HwpxPath $outPath -SchoolAddress $schoolAddress -SchoolPhone $schoolPhone -ManagerName $managerName -PrincipalName $principalName
            if (-not $?) { throw 'hwpx_school_footer_fill.ps1 실패' }
            Invoke-Kordoc @('validate', $outPath)
            $succeeded += "$formId(토큰없음, 원문 그대로)"
            continue
        }
        $missing = @($requiredTokens | Where-Object { -not $values.Contains($_) })
        if ($missing.Count -gt 0) {
            $skipped += "$formId(Excel 미입력: $($missing -join ', '))"
            continue
        }
        $payload = [ordered]@{}
        foreach ($t in $requiredTokens) { $payload[$t] = $values[$t] }
        $jsonPath = Join-Path $jsonWorkDir ("$formId.json")
        [IO.File]::WriteAllText($jsonPath, ($payload | ConvertTo-Json -Compress), [Text.UTF8Encoding]::new($false))

        $outName = $tpl.Name -replace '_템플릿\.hwpx$', '_생성본.hwpx'
        $outPath = Join-Path $outputDirFull $outName
        & (Join-Path $PSScriptRoot 'hwpx_token_fill.ps1') -TemplatePath $tpl.FullName -JsonPath $jsonPath -OutputPath $outPath
        if (-not $?) { throw 'hwpx_token_fill.ps1 실패' }
        & (Join-Path $PSScriptRoot 'hwpx_school_footer_fill.ps1') -HwpxPath $outPath -SchoolAddress $schoolAddress -SchoolPhone $schoolPhone -ManagerName $managerName -PrincipalName $principalName
        if (-not $?) { throw 'hwpx_school_footer_fill.ps1 실패' }
        Invoke-Kordoc @('validate', $outPath)
        $succeeded += "$formId($($requiredTokens -join ','))"
    } catch {
        $failed += "$formId : $($_.Exception.Message)"
    }
}

if (Test-Path -LiteralPath $jsonWorkDir) { Remove-Item -LiteralPath $jsonWorkDir -Recurse -Force -ErrorAction SilentlyContinue }

Write-Output ''
Write-Output "===== 결과 요약 ====="
Write-Output "성공 $($succeeded.Count)건:"
$succeeded | ForEach-Object { Write-Output "  PASS $_" }
Write-Output "건너뜀 $($skipped.Count)건(Excel 값 미입력, 오류 아님):"
$skipped | ForEach-Object { Write-Output "  SKIP $_" }
Write-Output "실패 $($failed.Count)건:"
$failed | ForEach-Object { Write-Output "  FAIL $_" }

if ($succeeded.Count -eq 0) { throw '성공한 서식이 하나도 없습니다.' }
if ($failed.Count -gt 0) { throw "$($failed.Count)건이 실패했습니다. 위 FAIL 목록을 확인하세요." }
Write-Output ''
Write-Output "PASS: P2-04 연동 생성 완료 -> $outputDirFull ($($succeeded.Count)/$($templates.Count) 서식)"
