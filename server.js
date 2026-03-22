const http = require('http');
const fs = require('fs');
const path = require('path');
const dir = __dirname;
http.createServer((req, res) => {
  let f = path.join(dir, req.url === '/' ? 'portal-v2.html' : req.url);
  fs.readFile(f, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    const ext = path.extname(f);
    const types = {'.html':'text/html','.css':'text/css','.js':'application/javascript','.docx':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'};
    res.writeHead(200, {'Content-Type': types[ext] || 'application/octet-stream'});
    res.end(data);
  });
}).listen(3456, () => console.log('Serving on http://localhost:3456'));
