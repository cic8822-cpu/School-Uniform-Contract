<#
.SYNOPSIS
  원본 보존형 HWPX 사본을 한컴 NEO로 PDF로 내보낸다.

.DESCRIPTION
  원본 HWPX를 수정하지 않고 HWPX→PDF의 한컴 실제 조판을 확보한다.
  입력은 artifacts/hwpx/P2-01, 출력은 artifacts/pdf/P2-01로 제한한다.
  이 자동 출력은 사람의 최종 사실관계·조판 검토를 대체하지 않는다.
#>
[CmdletBinding()]
param(
    [string]$InputPath = 'artifacts/hwpx/P2-01/교복매뉴얼_F007_원본보존형_골든결과.hwpx',
    [string]$OutputPath = 'artifacts/pdf/P2-01/교복매뉴얼_F007_원본보존형_골든결과_한컴.pdf'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$inputFullPath = [IO.Path]::GetFullPath($InputPath)
$outputFullPath = [IO.Path]::GetFullPath($OutputPath)
$inputRoot = [IO.Path]::GetFullPath('artifacts/hwpx/P2-01').TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
$outputRoot = [IO.Path]::GetFullPath('artifacts/pdf/P2-01').TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
if (-not $inputFullPath.StartsWith($inputRoot, [StringComparison]::OrdinalIgnoreCase)) { throw "입력 HWPX 경로는 $inputRoot 아래여야 합니다: $inputFullPath" }
if (-not $outputFullPath.StartsWith($outputRoot, [StringComparison]::OrdinalIgnoreCase)) { throw "출력 PDF 경로는 $outputRoot 아래여야 합니다: $outputFullPath" }
if (-not (Test-Path -LiteralPath $inputFullPath -PathType Leaf)) { throw "입력 HWPX를 찾을 수 없습니다: $inputFullPath" }
if ([IO.Path]::GetExtension($inputFullPath) -ne '.hwpx') { throw '입력 파일은 .hwpx여야 합니다.' }
if ([IO.Path]::GetExtension($outputFullPath) -ne '.pdf') { throw '출력 파일은 .pdf여야 합니다.' }

[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($outputFullPath)) | Out-Null
$inputHash = (Get-FileHash -LiteralPath $inputFullPath -Algorithm SHA256).Hash
$automationBefore = @(
    Get-CimInstance Win32_Process -Filter "Name='hwp.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -match '(?i)-automation\s+-embedding' } |
        ForEach-Object { $_.ProcessId }
)

$hwp = $null
try {
    $hwp = New-Object -ComObject 'HWPFrame.HwpObject'
    $null = $hwp.SetMessageBoxMode(0x1FFFF1)
    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null }
    catch { Write-Warning "한컴 경로 검사 모듈을 등록하지 못했습니다: $($_.Exception.Message)" }
    $hwp.XHwpWindows.Item(0).Visible = $false
    $null = $hwp.Open($inputFullPath, 'HWPX', 'forceopen:true')
    $null = $hwp.SaveAs($outputFullPath, 'PDF', '')
    $stableLength = -1
    $stableCount = 0
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        if (Test-Path -LiteralPath $outputFullPath -PathType Leaf) {
            $length = (Get-Item -LiteralPath $outputFullPath).Length
            if ($length -gt 0 -and $length -eq $stableLength) { $stableCount++ } else { $stableCount = 0 }
            $stableLength = $length
            if ($stableCount -ge 2) { break }
        }
        Start-Sleep -Milliseconds 500
    }
    if ($stableLength -le 0 -or $stableCount -lt 2) {
        throw '한컴에서 PDF를 저장하지 못했습니다.'
    }
    if ((Get-FileHash -LiteralPath $inputFullPath -Algorithm SHA256).Hash -ne $inputHash) { throw 'PDF 내보내기 중 입력 HWPX가 변경되었습니다.' }
    Write-Output "PASS: 한컴 HWPX→PDF 내보내기 및 입력 해시 불변 확인 -> $outputFullPath"
} finally {
    if ($hwp) {
        try { $hwp.Clear(1) } catch { Write-Verbose "hwp.Clear 무시 가능한 예외: $($_.Exception.Message)" }
        try { $hwp.Quit() } catch { Write-Verbose "hwp.Quit 무시 가능한 예외: $($_.Exception.Message)" }
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($hwp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # 이 실행이 새로 만든 헤드리스 COM 서버만 정리한다. 사용자가 연 일반 한컴 창은 대상이 아니다.
    $automationAfter = @(
        Get-CimInstance Win32_Process -Filter "Name='hwp.exe'" -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandLine -match '(?i)-automation\s+-embedding' } |
            ForEach-Object { $_.ProcessId }
    )
    foreach ($processId in $automationAfter) {
        if ($processId -notin $automationBefore) {
            Stop-Process -Id $processId -ErrorAction SilentlyContinue
        }
    }
}
