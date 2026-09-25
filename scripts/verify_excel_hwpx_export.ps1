<# .SYNOPSIS F-007 Excel→HWPX 허용값 내보내기 회귀 검증 #>
[CmdletBinding()]
param([string]$ExcelPath='artifacts/excel/교복구매_길라잡이_20260919_v7.xlsm',[string]$WorkDirectory='_workspace/02_hwpx/P2-01/excel_hwpx_export_검증산출')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$source=(Resolve-Path -LiteralPath $ExcelPath).Path
$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
[IO.Directory]::CreateDirectory($WorkDirectory)|Out-Null
$fixture=Join-Path $WorkDirectory 'F007_비식별_입력사본.xlsm'
$tokenDir='_workspace/02_hwpx/P2-01/토큰치환_검증산출'
$tokenTemplate=Join-Path $tokenDir 'F007_명시적토큰_템플릿.hwpx'
$output=Join-Path $WorkDirectory 'F007_Excel연계_골든결과.hwpx'
[IO.File]::Copy($source,$fixture,$true)
# 생성 입력본만 안전한 공통값으로 채운다. 원본 XLSM은 열거나 수정하지 않는다.
$excel=$null;$book=$null;$sheet=$null
try{$excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$book=$excel.Workbooks.Open(([IO.Path]::GetFullPath($fixture)),0,$false);$sheet=$book.Worksheets.Item('기초자료입력');$sheet.Range('C4').Value2='검증초등학교';$sheet.Range('C5').Value2='2026';$sheet.Range('C8').Value2='2026. 3. 2.';$sheet.Range('C9').Value2='검증초-2026-1';$sheet.Range('C21').Value2='내부결재';$sheet.Range('C22').Value2='2026학년도 교복 학교주관구매 요청';$book.Save()}finally{if($book){$book.Close($false)};if($sheet){[Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)|Out-Null};if($excel){$excel.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
# 기존 검증기는 기준 F-007을 사본 토큰 템플릿으로 만들며, 이 템플릿을 이후 읽기 전용으로 사용한다.
& (Join-Path $PSScriptRoot 'verify_hwpx_token_fill.ps1') -WorkDirectory $tokenDir
if(-not $?) {throw '명시 토큰 템플릿 준비 실패'}
& (Join-Path $PSScriptRoot 'export_hwpx_from_excel.ps1') -ExcelPath $fixture -TokenTemplatePath $tokenTemplate -OutputPath $output -WorkingDirectory $WorkDirectory
if(-not $?) {throw 'Excel→HWPX 내보내기 실패'}
if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash){throw '원본 XLSM 해시가 변경되었습니다.'}
$jsonCount=@(Get-ChildItem -LiteralPath $WorkDirectory -Filter '.hwpx_export_*.json' -File).Count
if($jsonCount -ne 0){throw '임시 JSON이 삭제되지 않았습니다.'}
Write-Output "PASS: 원본 XLSM 불변/허용 6값만 사용/임시 JSON 삭제/F-007 HWPX 출력 -> $output"
