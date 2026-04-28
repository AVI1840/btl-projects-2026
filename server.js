const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const PORT = 3456;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.js':   'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png':  'image/png',
  '.jpg':  'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg':  'image/svg+xml',
  '.ico':  'image/x-icon',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};

function resolveSafePath(url) {
  // Strip query string and decode URI
  const clean = decodeURIComponent(url.split('?')[0]);

  // Resolve to absolute path
  const abs = path.resolve(ROOT, clean === '/' ? 'portal-v2.html' : '.' + clean);

  // Ensure resolved path stays within ROOT
  if (!abs.startsWith(ROOT + path.sep) && abs !== ROOT) return null;

  return abs;
}

http.createServer((req, res) => {
  let filePath;

  try {
    filePath = resolveSafePath(req.url);
  } catch {
    res.writeHead(400);
    res.end('Bad request');
    return;
  }

  if (!filePath) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      if (err.code === 'ENOENT') {
        res.writeHead(404);
        res.end('Not found');
      } else {
        console.error(`[500] ${filePath}:`, err.message);
        res.writeHead(500);
        res.end('Server error');
      }
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
}).listen(PORT, () => console.log(`Serving on http://localhost:${PORT}`));
