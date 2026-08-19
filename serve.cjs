// Local static server for the tax-refund web app (temporary, pre-deploy).
const http = require('http');
const fs = require('fs');
const path = require('path');
const ROOT = __dirname;
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8', '.wasm': 'application/wasm',
  '.json': 'application/json', '.whl': 'application/octet-stream',
  '.zip': 'application/octet-stream', '.xlsx': 'application/octet-stream',
  '.docx': 'application/octet-stream', '.data': 'application/octet-stream',
  '.py': 'text/plain; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.png': 'image/png', '.ico': 'image/x-icon',
};
const server = http.createServer((req, res) => {
  let url = decodeURIComponent(req.url.split('?')[0]);
  if (url === '/') url = '/index.html';
  const fp = path.join(ROOT, url);
  // prevent path traversal
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end('forbidden'); return; }
  fs.stat(fp, (e, st) => {
    if (e || !st.isFile()) { res.writeHead(404); res.end('not found: ' + url); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(fp).toLowerCase()] || 'application/octet-stream',
                         'Access-Control-Allow-Origin': '*' });
    fs.createReadStream(fp).pipe(res);
  });
});
const ports = [8090, 8091, 8092, 8093, 8094, 8095];
(function tryPort(i) {
  if (i >= ports.length) { console.error('no free port'); process.exit(1); }
  const p = ports[i];
  server.once('error', err => { if (err.code === 'EADDRINUSE') tryPort(i + 1); else throw err; });
  server.listen(p, '127.0.0.1', () => console.log('SERVER_READY http://127.0.0.1:' + p + '/  (root=' + ROOT + ')'));
})(0);
