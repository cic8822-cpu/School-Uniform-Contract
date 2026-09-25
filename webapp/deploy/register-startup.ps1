# 교복 학교주관구매 길라잡이 - PC 부팅(로그인) 시 자동 실행 등록
#
# Windows 시작프로그램 폴더(shell:startup)에 "교복길라잡이_실행.bat"를 가리키는
# 바로가기를 만든다. 한 번만 실행하면 되며, 이후로는 PC에 로그인할 때마다
# 자동으로 로컬 서버가 켜지고 기본 브라우저가 열린다.
#
# 보통 이 스크립트를 직접 실행하지 않고 같은 폴더의 "자동시작_등록.bat"를
# 더블클릭해서 사용한다.
#
# 자동 실행을 그만두려면: 같은 폴더의 "자동시작_해제.bat"를 실행하거나,
# 아래에서 안내하는 바로가기 파일을 직접 삭제한다.

$ErrorActionPreference = 'Stop'

$target = Join-Path $PSScriptRoot '교복길라잡이_실행.bat'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Host "[오류] $target 을 찾을 수 없습니다. 이 스크립트를 옮기지 말고 원래 폴더에서 실행하세요." -ForegroundColor Red
    exit 1
}

$startupDir = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
$shortcutPath = Join-Path $startupDir '교복길라잡이.lnk'

try {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $target
    $shortcut.WorkingDirectory = $PSScriptRoot
    $shortcut.WindowStyle = 7  # 최소화 상태로 시작
    $shortcut.Description = '교복 학교주관구매 길라잡이 자동 실행'
    $shortcut.Save()
} catch {
    Write-Host "[오류] 바로가기 생성에 실패했습니다: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

if (Test-Path -LiteralPath $shortcutPath) {
    Write-Host "등록되었습니다. 다음부터 이 PC에 로그인할 때마다 자동으로 서버가 실행되고 브라우저가 열립니다." -ForegroundColor Green
    Write-Host "바로가기 위치: $shortcutPath"
    Write-Host "자동 실행을 그만두려면 이 폴더의 자동시작_해제.bat를 실행하세요."
} else {
    Write-Host "[오류] 바로가기 생성에 실패했습니다." -ForegroundColor Red
    exit 1
}
