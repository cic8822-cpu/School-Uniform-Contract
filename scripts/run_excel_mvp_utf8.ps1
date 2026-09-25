$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot 'create_excel_mvp.ps1'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath
if ($LASTEXITCODE -ne 0) {
    throw "초안 Excel 생성 스크립트가 실패했습니다. 종료 코드: $LASTEXITCODE"
}
