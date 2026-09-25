$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$excelDir = Join-Path $root 'artifacts\excel'
# Keep the workbook base name as Unicode code points for Windows PowerShell 5.1
# compatibility when this UTF-8 script is read without a BOM.
$documentName = -join [char[]](0xAD50, 0xBCF5, 0xAD6C, 0xB9E4, 0x005F, 0xAE38, 0xB77C, 0xC7A1, 0xC774)
$dateStamp = Get-Date -Format 'yyyyMMdd'
$versionPattern = "^{0}_{1}_v(\d+)\.xlsm$" -f [regex]::Escape($documentName), $dateStamp
$existingVersions = @(Get-ChildItem -LiteralPath $excelDir -File | ForEach-Object {
	$match = [regex]::Match($_.Name, $versionPattern)
	if ($match.Success) { [int]$match.Groups[1].Value }
})
$nextVersion = if ($existingVersions.Count -eq 0) { 1 } else { (($existingVersions | Measure-Object -Maximum).Maximum + 1) }
$finalPath = Join-Path $excelDir ('{0}_{1}_v{2}.xlsm' -f $documentName, $dateStamp, $nextVersion)
$stagingDir = Join-Path $root 'artifacts\excel\_staging'
$stagingPath = Join-Path $root 'artifacts\excel\build.xlsm'
New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null
Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue

$env:UNIFORM_EXCEL_BUILD_PATH = $stagingPath
try {
	$scriptPath = Join-Path $PSScriptRoot 'build_excel_v1_structure.ps1'
	$scriptText = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
	$escapedRoot = $PSScriptRoot.Replace("'", "''")
	$scriptText = $scriptText.Replace('$PSScriptRoot', "'$escapedRoot'")
	& ([scriptblock]::Create($scriptText))
	& (Join-Path $PSScriptRoot 'build_excel_v1_vba.ps1')
	& (Join-Path $PSScriptRoot 'apply_manual_formatting_overlay.ps1') -TargetPath $stagingPath
	& (Join-Path $PSScriptRoot 'verify_excel_v1.ps1')

	[System.IO.File]::Move($stagingPath, $finalPath)
	Write-Output "PASS: Saved verified workbook with name_date_version format: $finalPath"
} finally {
	Remove-Item Env:UNIFORM_EXCEL_BUILD_PATH -ErrorAction SilentlyContinue
}
