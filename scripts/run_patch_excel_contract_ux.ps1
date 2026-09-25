$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$documentName = -join [char[]](0xAD50, 0xBCF5, 0xAD6C, 0xB9E4, 0x005F, 0xAE38, 0xB77C, 0xC7A1, 0xC774)
$source = Join-Path $root ("artifacts\excel\{0}_20260923_v9.xlsm" -f $documentName)
$staging = Join-Path $root 'artifacts\excel\_staging\contract_ux_build.xlsm'
$final = Join-Path $root ("artifacts\excel\{0}_20260923_v10.xlsm" -f $documentName)

$patchPath = Join-Path $PSScriptRoot 'patch_excel_contract_ux.ps1'
$patchText = [IO.File]::ReadAllText($patchPath, [Text.Encoding]::UTF8)
$temporaryPatch = Join-Path ([IO.Path]::GetTempPath()) 'uniform_patch_excel_contract_ux_utf8.ps1'
$bom = [byte[]](0xEF, 0xBB, 0xBF)
$body = [Text.Encoding]::UTF8.GetBytes($patchText)
[IO.File]::WriteAllBytes($temporaryPatch, $bom + $body)
try {
    & $temporaryPatch -InputPath $source -OutputPath $staging
} finally {
    Remove-Item -LiteralPath $temporaryPatch -Force -ErrorAction SilentlyContinue
}
# 기존 검증본의 VBA 모듈은 보존하고, 구조 패치가 추가한 전용 모듈만 사용한다.
Move-Item -LiteralPath $staging -Destination $final -Force
Write-Output "PASS: 계약방식 UX 검증본 생성 -> $final"
