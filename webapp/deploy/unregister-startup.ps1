# 교복 학교주관구매 길라잡이 - PC 부팅 시 자동 실행 해제
#
# register-startup.ps1(자동시작_등록.bat)이 만든 시작프로그램 바로가기를 삭제한다.

$startupDir = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
$shortcutPath = Join-Path $startupDir '교복길라잡이.lnk'

if (Test-Path -LiteralPath $shortcutPath) {
    Remove-Item -LiteralPath $shortcutPath -Force
    Write-Host "자동 실행이 해제되었습니다." -ForegroundColor Green
} else {
    Write-Host "등록된 자동 실행이 없습니다."
}
