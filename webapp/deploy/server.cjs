#!/usr/bin/env node
// 교복 학교주관구매 길라잡이 - 단일 실행 파일(exe)용 정적 서버
//
// 이 파일은 @yao-pkg/pkg로 webapp/deploy/dist-exe/*.exe 하나로 번들링된다.
// Node.js가 설치되지 않은 학교 PC에서도 실행 파일 자체에 Node 런타임이
// 포함되어 있어 그대로 동작한다. 서버는 "정적 파일을 그대로 내려주는 역할"만
// 하며 어떤 비즈니스 로직·DB도 갖지 않는다(ADR-002: 서버 없는 정적 웹앱).
// 기초자료 입력값·엑셀 파싱·PDF/HWPX 생성은 여전히 전부 브라우저 안에서
// 실행되고, 이 서버는 그 요청 바디를 받지 않는다.
//
// 개발 중(패키징 전)에는 `node deploy/server.cjs`로 그대로 실행할 수 있고,
// 이때는 이 파일 기준 상위의 dist(webapp/dist)를 그대로 서빙한다.
// pkg로 패키징된 exe에서는 같은 경로(__dirname 기준 상대 경로)가 pkg의
// 가상 스냅샷 파일시스템(/snapshot/webapp/dist)으로 자동 매핑된다 — 이
// 매핑이 깨지지 않으려면 package.json의 `pkg.assets`(`dist/**/*`)와 이
// 서버가 참조하는 상대 경로 구조가 항상 일치해야 한다.

const http = require('node:http')
const fs = require('node:fs')
const path = require('node:path')
const { exec } = require('node:child_process')

const DEFAULT_PORT = 8973
const port = Number(process.env.PORT) || DEFAULT_PORT

const distPath = path.join(__dirname, '..', 'dist')

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.hwpx': 'application/octet-stream',
  '.xlsx': 'application/octet-stream',
}

/** 요청 경로를 dist 폴더 안의 실제 파일로 해석한다(경로 순회 방지, SPA 폴백 포함). */
function resolveRequestedFile(urlPath) {
  const decodedPath = decodeURIComponent(urlPath)
  const relativePath = decodedPath === '/' || decodedPath === '' ? 'index.html' : decodedPath.replace(/^\/+/, '')
  const candidatePath = path.join(distPath, relativePath)

  if (fs.existsSync(candidatePath) && fs.statSync(candidatePath).isFile()) {
    const resolved = path.resolve(candidatePath)
    const resolvedRoot = path.resolve(distPath)
    if (resolved === resolvedRoot || resolved.startsWith(resolvedRoot + path.sep)) {
      return resolved
    }
  }

  // 이 웹앱은 클라이언트 상태 기반 화면 전환만 쓰고 경로 라우팅은 쓰지 않지만,
  // 알 수 없는 경로 요청은 index.html로 안전하게 되돌려준다(SPA 폴백).
  const indexPath = path.join(distPath, 'index.html')
  return fs.existsSync(indexPath) ? indexPath : null
}

function handleRequest(request, response) {
  try {
    const urlPath = new URL(request.url ?? '/', 'http://localhost').pathname
    const resolvedFile = resolveRequestedFile(urlPath)

    if (!resolvedFile) {
      response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' })
      response.end('404 Not Found')
      return
    }

    const ext = path.extname(resolvedFile).toLowerCase()
    const contentType = MIME_TYPES[ext] ?? 'application/octet-stream'
    const fileBytes = fs.readFileSync(resolvedFile)
    response.writeHead(200, { 'Content-Type': contentType, 'Content-Length': fileBytes.length })
    response.end(fileBytes)
  } catch (error) {
    console.error('[오류] 요청 처리 중 예외가 발생했습니다:', error instanceof Error ? error.message : error)
    response.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' })
    response.end('500 Internal Server Error')
  }
}

/** Windows 기본 브라우저로 주소를 연다(실패해도 서버 동작에는 영향 없음). */
function openBrowser(url) {
  if (process.platform !== 'win32') return
  exec(`start "" "${url}"`, (error) => {
    if (error) {
      console.log(`[안내] 브라우저를 자동으로 열지 못했습니다. 주소창에 직접 입력하세요: ${url}`)
    }
  })
}

function main() {
  if (!fs.existsSync(distPath)) {
    console.error(`[오류] dist 폴더를 찾을 수 없습니다: ${distPath}`)
    console.error('먼저 webapp 폴더에서 "npm run build"를 실행해 dist를 생성해야 합니다.')
    process.exitCode = 1
    return
  }

  startServer(port, { isFallback: false })
}

/**
 * 지정한 포트로 서버를 띄운다. 고정 포트가 이미 사용 중이거나(EADDRINUSE)
 * OS/보안 소프트웨어가 예약해 거부하는 경우(EACCES — 학교 PC의 Hyper-V·VPN
 * 등이 임의 포트를 동적으로 예약해두는 사례가 실측으로 확인됨) 포트 0으로
 * 재시도해 OS가 비어 있는 포트를 자동으로 골라주게 한다. PORT 환경변수로
 * 직접 포트를 지정했을 때는 자동 대체를 하지 않고 원인만 안내한다.
 */
function startServer(requestedPort, { isFallback }) {
  const server = http.createServer(handleRequest)

  server.on('error', (error) => {
    const canFallback = !isFallback && !process.env.PORT && (error.code === 'EADDRINUSE' || error.code === 'EACCES')
    if (canFallback) {
      console.log(`[안내] 포트 ${requestedPort}를 쓸 수 없어(${error.code}) 빈 포트를 자동으로 찾습니다...`)
      startServer(0, { isFallback: true })
      return
    }

    if (error.code === 'EADDRINUSE') {
      console.error(`[오류] 포트 ${requestedPort}를(을) 사용할 수 없습니다(이미 다른 프로그램이 쓰고 있을 수 있습니다).`)
      console.error('다른 포트로 다시 시도하세요: PORT=9000 (환경변수로 지정)')
    } else if (error.code === 'EACCES') {
      console.error(`[오류] 포트 ${requestedPort}를(을) 쓸 권한이 없습니다(시스템이 예약해둔 포트일 수 있습니다).`)
      console.error('다른 포트로 다시 시도하세요: PORT=9000 (환경변수로 지정)')
    } else {
      console.error(`[오류] 서버를 시작하지 못했습니다: ${error.message}`)
    }
    process.exitCode = 1
  })

  server.listen(requestedPort, () => {
    const actualPort = server.address().port
    const url = `http://localhost:${actualPort}/`
    console.log('교복 학교주관구매 길라잡이 서버를 시작합니다.')
    console.log(`주소: ${url}`)
    console.log(`서빙 폴더: ${distPath}`)
    console.log('이 창을 닫으면 서버가 종료됩니다.')
    console.log('')
    openBrowser(url)
  })
}

main()
