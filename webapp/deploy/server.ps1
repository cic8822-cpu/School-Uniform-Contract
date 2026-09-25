# 교복 학교주관구매 길라잡이 - 로컬 정적 서버
#
# Windows에 기본 내장된 PowerShell + .NET System.Net.HttpListener만 사용한다.
# Node.js·Python 등 추가 설치가 전혀 필요 없다(ADR-002: 서버 없는 정적 웹앱 —
# 이 서버는 "정적 파일을 그대로 내려주는 역할"만 하며, 어떤 비즈니스 로직·DB도
# 갖지 않는다. 기초자료·엑셀 파싱·PDF/HWPX 생성은 여전히 전부 브라우저 안에서
# 실행된다).
#
# 사용법(직접 실행):
#   powershell -ExecutionPolicy Bypass -File server.ps1
#   powershell -ExecutionPolicy Bypass -File server.ps1 -Port 9000
#
# 보통은 이 스크립트를 직접 실행하지 않고, 같은 폴더의
# "교복길라잡이_실행.bat"를 더블클릭해서 사용한다.

param(
    [int]$Port = 8973
)

$ErrorActionPreference = 'Stop'

# dist 폴더 경로: 이 스크립트(webapp/deploy/server.ps1) 기준 상위의 dist
$distPath = Join-Path $PSScriptRoot '..\dist'
if (-not (Test-Path $distPath)) {
    Write-Host "[오류] dist 폴더를 찾을 수 없습니다: $distPath" -ForegroundColor Red
    Write-Host "먼저 webapp 폴더에서 'npm run build'를 실행해 dist를 생성해야 합니다." -ForegroundColor Yellow
    exit 1
}
$distPath = (Resolve-Path $distPath).Path

$mimeTypes = @{
    '.html'  = 'text/html; charset=utf-8'
    '.js'    = 'text/javascript; charset=utf-8'
    '.mjs'   = 'text/javascript; charset=utf-8'
    '.css'   = 'text/css; charset=utf-8'
    '.json'  = 'application/json; charset=utf-8'
    '.svg'   = 'image/svg+xml'
    '.png'   = 'image/png'
    '.jpg'   = 'image/jpeg'
    '.jpeg'  = 'image/jpeg'
    '.ico'   = 'image/x-icon'
    '.woff'  = 'font/woff'
    '.woff2' = 'font/woff2'
    '.hwpx'  = 'application/octet-stream'
    '.xlsx'  = 'application/octet-stream'
}

$listener = New-Object System.Net.HttpListener
$prefix = "http://localhost:$Port/"
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
} catch {
    Write-Host "[오류] 포트 $Port 를 사용할 수 없습니다(이미 다른 프로그램이 쓰고 있을 수 있습니다)." -ForegroundColor Red
    Write-Host "다른 포트로 다시 시도하세요: powershell -File server.ps1 -Port 9000" -ForegroundColor Yellow
    exit 1
}

Write-Host "교복 학교주관구매 길라잡이 서버를 시작합니다." -ForegroundColor Green
Write-Host "주소: $prefix" -ForegroundColor Green
Write-Host "서빙 폴더: $distPath"
Write-Host "이 창을 닫거나 Ctrl+C를 누르면 서버가 종료됩니다."
Write-Host ""

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        try {
            $urlPath = [System.Uri]::UnescapeDataString($request.Url.AbsolutePath)
            if ([string]::IsNullOrEmpty($urlPath) -or $urlPath -eq '/') { $urlPath = '/index.html' }

            $relativePath = $urlPath.TrimStart('/') -replace '/', [System.IO.Path]::DirectorySeparatorChar
            $candidatePath = Join-Path $distPath $relativePath

            $resolvedFile = $null
            if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
                $resolved = (Resolve-Path -LiteralPath $candidatePath).Path
                # 경로 순회(path traversal) 방지: dist 폴더 밖 파일은 절대 내려주지 않는다.
                if ($resolved.StartsWith($distPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $resolvedFile = $resolved
                }
            }

            if (-not $resolvedFile) {
                # 이 웹앱은 클라이언트 상태 기반 화면 전환만 쓰고 경로 라우팅은
                # 쓰지 않는다(react-router 등 없음). 그래도 알 수 없는 경로 요청은
                # index.html로 안전하게 되돌려준다.
                $indexPath = Join-Path $distPath 'index.html'
                if (Test-Path -LiteralPath $indexPath -PathType Leaf) {
                    $resolvedFile = (Resolve-Path -LiteralPath $indexPath).Path
                }
            }

            if ($resolvedFile) {
                $ext = [System.IO.Path]::GetExtension($resolvedFile).ToLowerInvariant()
                $contentType = $mimeTypes[$ext]
                if (-not $contentType) { $contentType = 'application/octet-stream' }
                $bytes = [System.IO.File]::ReadAllBytes($resolvedFile)
                $response.ContentType = $contentType
                $response.ContentLength64 = $bytes.Length
                $response.OutputStream.Write($bytes, 0, $bytes.Length)
            } else {
                $response.StatusCode = 404
                $notFoundBytes = [System.Text.Encoding]::UTF8.GetBytes('404 Not Found')
                $response.OutputStream.Write($notFoundBytes, 0, $notFoundBytes.Length)
            }
        } catch {
            try { $response.StatusCode = 500 } catch { }
        } finally {
            $response.OutputStream.Close()
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
