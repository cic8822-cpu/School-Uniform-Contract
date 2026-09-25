import http from 'node:http'

const upstreamPort = 4174
const port = 4175

http
  .createServer((request, response) => {
    const upstream = http.request(
      { host: '127.0.0.1', port: upstreamPort, path: request.url, method: request.method, headers: request.headers },
      (upstreamResponse) => {
        response.writeHead(upstreamResponse.statusCode ?? 502, upstreamResponse.headers)
        const delay = request.url?.includes('/templates/F-008_') ? 1200 : 0
        if (delay > 0) {
          setTimeout(() => upstreamResponse.pipe(response), delay)
        } else {
          upstreamResponse.pipe(response)
        }
      }
    )
    upstream.on('error', (error) => {
      response.statusCode = 502
      response.end(error.message)
    })
    request.pipe(upstream)
  })
  .listen(port, '127.0.0.1', () => console.log(`slow-template-proxy listening on ${port}`))
