#!/usr/bin/env node
/*
 * serve.js — 静态 HTTP server，给 Office 加载项用
 *
 * - 监听 http://localhost:3000
 * - 根目录 = 项目根（包含 manifest.xml、src/、assets/）
 * - Office Add-in 对 localhost 允许 HTTP（不要求 HTTPS）
 * - 零配置：开箱即用
 *
 * 启动：npm start   或   node tools/serve.js
 */

const fs = require('fs');
const path = require('path');
const http = require('http');

const ROOT = path.resolve(__dirname, '..');
const PORT = parseInt(process.env.PORT || '3000', 10);
const HOST = process.env.HOST || '127.0.0.1';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.md': 'text/markdown; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8',
};

function safeJoin(root, urlPath) {
  const p = path.normalize(path.join(root, urlPath));
  if (!p.startsWith(root)) return null;
  return p;
}

function serveStatic(req, res) {
  let urlPath = decodeURIComponent(req.url.split('?')[0]);
  if (urlPath === '/') urlPath = '/README.md'; // 默认首页
  const filePath = safeJoin(ROOT, urlPath);
  if (!filePath) {
    res.writeHead(403);
    return res.end('Forbidden');
  }
  fs.stat(filePath, (err, stat) => {
    if (err || !stat.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end('Not Found: ' + urlPath);
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      'Content-Type': MIME[ext] || 'application/octet-stream',
      'Content-Length': stat.size,
      // Office 加载项对缓存敏感，开发期关掉
      'Cache-Control': 'no-store',
    });
    fs.createReadStream(filePath).pipe(res);
  });
}

const server = http.createServer(serveStatic);
server.listen(PORT, HOST, () => {
  console.log(`[serve] HTTP listening on http://${HOST}:${PORT}`);
  console.log(`[serve] 加载项入口:  http://${HOST}:${PORT}/manifest.xml`);
  console.log(`[serve] Dialog 入口: http://${HOST}:${PORT}/src/dialog/dialog.html`);
});

// 优雅退出（让 .app 启动脚本能干净地 stop）
['SIGINT', 'SIGTERM'].forEach((sig) => {
  process.on(sig, () => {
    console.log(`[serve] received ${sig}, shutting down`);
    server.close(() => process.exit(0));
  });
});
