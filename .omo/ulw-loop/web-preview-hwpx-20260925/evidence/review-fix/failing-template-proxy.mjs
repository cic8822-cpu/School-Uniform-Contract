import http from 'node:http'

const upstreamPort = 4174
const port = 4176

http
  .createServer((request, response) => {
    if (request.url?.includes('/templates/F-008_')) {
      response.statusCode = 503
      response.end('temporary template failure')
      return
    }
    const upstream = http.request(
      { host: '127.0.0.1', port: upstreamPort, path: request.url, method: request.method, headers: request.headers },
      (upstreamResponse) => {
        response.writeHead(upstreamResponse.statusCode ?? 502, upstreamResponse.headers)
        upstreamResponse.pipe(response)
      }
    )
    upstream.on('error', (error) => {
      response.statusCode = 502
      response.end(error.message)
    })
    request.pipe(upstream)
  })
  .listen(port, '127.0.0.1', () => console.log(`failing-template-proxy listening on ${port}`))
