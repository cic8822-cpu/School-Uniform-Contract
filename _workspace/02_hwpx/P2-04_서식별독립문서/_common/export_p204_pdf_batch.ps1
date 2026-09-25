<#
.SYNOPSIS
  P2-04 서식별 독립 HWPX(템플릿/골든) 여러 건을 한 번의 한컴 COM 세션에서
  순차적으로 PDF로 내보낸다(F-032~F-043 그룹 재사용).

.DESCRIPTION
  scripts/export_hwpx_hancom_pdf.ps1(P2-01용, 입출력 경로가 P2-01로 고정됨)을
  P2-04 배치 처리용으로 일반화한 버전. 입력은 artifacts/hwpx/P2-04, 출력은
  artifacts/pdf/P2-04로 제한한다. hwp.exe COM 프로세스를 한 번만 띄워 여러
  파일을 처리해 반복 기동 오버헤드를 줄인다. 각 입력 파일의 SHA-256을
  변환 전후로 비교해 원본 불변을 확인한다.

.PARAMETER Names
  확장자 없는 파일 이름(예: "F-033_제안서접수대장_템플릿") 목록. 쉼표 구분.
#>
[CmdletBinding()]
param(
    [string[]]$Names,
    [string]$NamesFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($NamesFile) {
    $Names = Get-Content -LiteralPath $NamesFile -Encoding UTF8 | Where-Object { $_.Trim() -ne '' }
}
if (-not $Names -or $Names.Count -eq 0) { throw 'Names 또는 NamesFile 중 하나를 지정해야 합니다.' }

$inputRoot = [IO.Path]::GetFullPath('artifacts/hwpx/P2-04').TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
$outputRoot = [IO.Path]::GetFullPath('artifacts/pdf/P2-04').TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
[IO.Directory]::CreateDirectory($outputRoot) | Out-Null

$automationBefore = @(
    Get-CimInstance Win32_Process -Filter "Name='hwp.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -match '(?i)-automation\s+-embedding' } |
        ForEach-Object { $_.ProcessId }
)

$hwp = $null
$results = @()
try {
    $hwp = New-Object -ComObject 'HWPFrame.HwpObject'
    $null = $hwp.SetMessageBoxMode(0x1FFFF1)
    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null }
    catch { Write-Warning "한컴 경로 검사 모듈을 등록하지 못했습니다: $($_.Exception.Message)" }
    $hwp.XHwpWindows.Item(0).Visible = $false

    foreach ($name in $Names) {
        $inputFullPath = Join-Path $inputRoot "$name.hwpx"
        $outputFullPath = Join-Path $outputRoot "$name`_한컴.pdf"
        if (-not (Test-Path -LiteralPath $inputFullPath -PathType Leaf)) {
            Write-Warning "입력 HWPX 없음, 건너뜀: $inputFullPath"
            continue
        }
        $inputHash = (Get-FileHash -LiteralPath $inputFullPath -Algorithm SHA256).Hash
        if (Test-Path -LiteralPath $outputFullPath) { Remove-Item -LiteralPath $outputFullPath -Force }

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
        try { $hwp.Clear(1) } catch { Write-Verbose "hwp.Clear 무시 가능: $($_.Exception.Message)" }

        $outputHashOk = $stableLength -gt 0 -and $stableCount -ge 2
        $inputUnchanged = (Get-FileHash -LiteralPath $inputFullPath -Algorithm SHA256).Hash -eq $inputHash
        $status = if ($outputHashOk -and $inputUnchanged) { 'PASS' } else { 'FAIL' }
        $results += [PSCustomObject]@{ Name = $name; Status = $status; OutputPath = $outputFullPath; InputUnchanged = $inputUnchanged }
        Write-Output "$status : $name -> $outputFullPath (입력 불변: $inputUnchanged)"
    }
} finally {
    if ($hwp) {
        try { $hwp.Quit() } catch { Write-Verbose "hwp.Quit 무시 가능: $($_.Exception.Message)" }
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($hwp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

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

$fail = $results | Where-Object { $_.Status -ne 'PASS' }
if ($fail) {
    Write-Error "실패 항목 있음: $($fail.Name -join ', ')"
    exit 1
}
exit 0
