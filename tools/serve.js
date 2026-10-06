#!/usr/bin/env node
/*
 * serve.js — 极简静态文件 server（Office Add-in 用）
 *
 * 所有 .pptx 读写都直接通过 Office.js 在 PowerPoint 进程内完成，
 * 这个 server 只负责给 dialog 提供 HTML / CSS / JS 资源。
 *
 * 启动：node tools/serve.js
 *      或 .app 双击
 */

const fs = require('fs');
const path = require('path');
const http = require('http');

const ROOT = path.resolve(__dirname, '..');
const PORT = parseInt(process.env.PORT || '3000', 10);
const HOST = process.env.HOST || '127.0.0.1';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml':  'application/xml; charset=utf-8',
  '.png':  'image/png',
  '.jpg':  'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg':  'image/svg+xml',
  '.ico':  'image/x-icon',
};

function createStaticServer(options) {
  const root = path.resolve(options && options.root || ROOT);
  const isPublicPath = (relative) => !relative.startsWith('../') && !path.isAbsolute(relative) &&
    (relative === 'manifest.xml' || relative.startsWith('src/') || relative.startsWith('assets/'));
  return http.createServer((req, res) => {
    const reply = (status, message) => {
      res.writeHead(status, { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' });
      res.end(message);
    };
    if (req.method !== 'GET' && req.method !== 'HEAD') return reply(405, 'Method Not Allowed');
    let urlPath;
    try {
      urlPath = decodeURIComponent((req.url || '/').split('?')[0]);
      if (urlPath.includes('\0')) return reply(400, 'Bad Request');
    } catch (_) { return reply(400, 'Bad Request'); }
    if (urlPath === '/') urlPath = '/manifest.xml';
    // Serve only add-in resources, never repository files or sibling paths.
    const filePath = path.resolve(root, '.' + urlPath);
    const relative = path.relative(root, filePath).split(path.sep).join('/');
    if (!isPublicPath(relative)) {
      return reply(403, 'Forbidden');
    }
    fs.realpath(filePath, (realError, realPath) => {
      if (realError) return reply(404, 'Not Found');
      const actualRelative = path.relative(root, realPath).split(path.sep).join('/');
      if (!isPublicPath(actualRelative)) return reply(403, 'Forbidden');
      fs.stat(realPath, (err, stat) => {
        if (err || !stat.isFile()) return reply(404, 'Not Found');
        const stream = fs.createReadStream(realPath);
        stream.on('error', () => {
          if (!res.headersSent) reply(500, 'Read Failed');
          else res.destroy();
        });
        stream.on('open', () => {
          res.writeHead(200, {
            'Content-Type': MIME[path.extname(realPath).toLowerCase()] || 'application/octet-stream',
            'Content-Length': stat.size,
            'Cache-Control': 'no-store',
          });
          if (req.method === 'HEAD') { stream.destroy(); res.end(); }
          else stream.pipe(res);
        });
        res.on('close', () => stream.destroy());
      });
    });
  });
}

if (require.main === module) {
  const server = createStaticServer();
  server.listen(PORT, HOST, () => console.log(`[serve] listening on http://${HOST}:${server.address().port}`));
  ['SIGINT', 'SIGTERM'].forEach((sig) => {
    process.on(sig, () => server.close(() => process.exit(0)));
  });
}
module.exports = { createStaticServer };
