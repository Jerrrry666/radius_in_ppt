#!/usr/bin/env node
/*
 * serve.js — 静态 HTTPS server，给 Office 加载项用
 *
 * - 监听 https://localhost:3000
 * - 根目录 = 项目根（包含 manifest.xml、src/、assets/）
 * - 首次运行自动用 openssl 在 ./certs/ 下生成自签证书
 * - 静态文件 / 简易目录浏览（无 SPA fallback；需要精确路径）
 *
 * 启动：npm start
 */

const fs = require('fs');
const path = require('path');
const https = require('https');
const { execSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const PORT = parseInt(process.env.PORT || '3000', 10);
const HOST = process.env.HOST || 'localhost';
const CERTS_DIR = path.join(ROOT, 'certs');
const KEY_FILE = path.join(CERTS_DIR, 'server.key');
const CRT_FILE = path.join(CERTS_DIR, 'server.crt');

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
};

function ensureCerts() {
  if (fs.existsSync(KEY_FILE) && fs.existsSync(CRT_FILE)) {
    return;
  }
  console.log('[serve] 生成自签证书到 ./certs/ ...');
  fs.mkdirSync(CERTS_DIR, { recursive: true });

  // 用一个 SAN 友好的自签证书（覆盖 localhost / 127.0.0.1 / ::1）
  const subj = '/CN=localhost';
  const san = 'subjectAltName=DNS:localhost,IP:127.0.0.1,IP:::1';
  const cmd = [
    'openssl', 'req', '-x509', '-nodes',
    '-newkey', 'rsa:2048',
    '-keyout', `"${KEY_FILE}"`,
    '-out', `"${CRT_FILE}"`,
    '-days', '365',
    '-subj', `"${subj}"`,
    '-addext', `"${san}"`,
  ].join(' ');
  execSync(cmd, { stdio: 'inherit', shell: true });
  console.log('[serve] 证书生成完成。\n  ⚠️  首次访问 https://localhost:3000 时浏览器会提示不安全，');
  console.log('      需要手动「显示详细信息 → 访问此网站」以信任自签证书。');
}

function safeJoin(root, urlPath) {
  // 防止 ../
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
      'Cache-Control': 'no-store',
    });
    fs.createReadStream(filePath).pipe(res);
  });
}

function main() {
  ensureCerts();
  const server = https.createServer(
    {
      key: fs.readFileSync(KEY_FILE),
      cert: fs.readFileSync(CRT_FILE),
    },
    serveStatic
  );
  server.listen(PORT, HOST, () => {
    console.log(`[serve] HTTPS listening on https://${HOST}:${PORT}`);
    console.log(`[serve] 加载项入口:  https://${HOST}:${PORT}/manifest.xml`);
    console.log(`[serve] Dialog 入口: https://${HOST}:${PORT}/src/dialog/dialog.html`);
  });
}

main();
