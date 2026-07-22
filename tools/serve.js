#!/usr/bin/env node
/*
 * serve.js — 静态 HTTP server + 写回 .pptx 的 API
 *
 * - 监听 http://localhost:3000
 * - 静态文件 = 项目根（manifest.xml、src/、assets/）
 * - API:
 *     GET  /api/pptx-path         → 当前 PowerPoint 打开的 .pptx 路径
 *     POST /api/save-pptx?path=…  body=base64 → 写回磁盘
 *     GET  /api/health            → {"ok":true}
 *
 * 启动：npm start   或   node tools/serve.js
 */

const fs = require('fs');
const path = require('path');
const http = require('http');
const { execSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const PORT = parseInt(process.env.PORT || '3000', 10);
const HOST = process.env.HOST || '127.0.0.1';
const PPTX_PATH_FILE = '/tmp/radius_in_ppt_pptx_path.txt';

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

// ---------------- 拿当前 PowerPoint .pptx 路径（从 /tmp 缓存文件读） ----------------

function getCachedPptxPath() {
  try {
    const p = fs.readFileSync(PPTX_PATH_FILE, 'utf8').trim();
    if (p && fs.existsSync(p)) return p;
  } catch (_) {}
  return null;
}

function tryDetectPptxFromProcess() {
  // 备选：lsof 查 powerpoint 进程
  try {
    const pids = execSync('pgrep -f "Microsoft PowerPoint" || true', { encoding: 'utf8' }).trim();
    if (!pids) return null;
    for (const pid of pids.split('\n')) {
      try {
        const lsofOut = execSync(`lsof -p ${pid} 2>/dev/null | grep -E '\\.pptx$|\\.ppt$' || true`, { encoding: 'utf8' });
        const m = lsofOut.match(/(\/[\w\/.\-]+\.pptx)/);
        if (m) return m[1];
      } catch (_) {}
    }
  } catch (_) {}
  return null;
}

// ---------------- HTTP handlers ----------------

function readBodyBase64(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      const buf = Buffer.concat(chunks);
      // 去掉可能的 data:application/zip;base64, 前缀
      let s = buf.toString('utf8');
      const comma = s.indexOf(',');
      if (s.startsWith('data:') && comma !== -1) s = s.substring(comma + 1);
      resolve(s.replace(/\s+/g, ''));
    });
    req.on('error', reject);
  });
}

function sendJSON(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body),
    'Access-Control-Allow-Origin': '*',
  });
  res.end(body);
}

function handleApi(req, res) {
  const url = req.url.split('?')[0];
  const query = {};
  req.url.split('?')[1]?.split('&').forEach((kv) => {
    const [k, v] = kv.split('=');
    if (k) query[decodeURIComponent(k)] = decodeURIComponent(v || '');
  });

  if (url === '/api/health') {
    return sendJSON(res, 200, { ok: true });
  }

  if (url === '/api/pptx-path') {
    let p = getCachedPptxPath();
    if (!p) p = tryDetectPptxFromProcess();
    if (!p) return sendJSON(res, 404, { error: 'no pptx detected; please launch R 角调整.app first' });
    return sendJSON(res, 200, { path: p });
  }

  if (url === '/api/save-pptx') {
    const targetPath = query.path || getCachedPptxPath();
    if (!targetPath) {
      return sendJSON(res, 400, { error: 'no target path' });
    }
    readBodyBase64(req).then((b64) => {
      try {
        const bin = Buffer.from(b64, 'base64');
        // 备份一份
        try {
          fs.copyFileSync(targetPath, targetPath + '.bak');
        } catch (_) {}
        fs.writeFileSync(targetPath, bin);
        console.log(`[serve] wrote ${bin.length} bytes to ${targetPath}`);
        sendJSON(res, 200, { ok: true, bytes: bin.length, path: targetPath });
      } catch (err) {
        console.error('[serve] save-pptx failed:', err);
        sendJSON(res, 500, { error: err.message });
      }
    }).catch((err) => sendJSON(res, 400, { error: err.message }));
    return;
  }

  sendJSON(res, 404, { error: 'unknown api' });
}

function safeJoin(root, urlPath) {
  const p = path.normalize(path.join(root, urlPath));
  if (!p.startsWith(root)) return null;
  return p;
}

function serveStatic(req, res) {
  let urlPath = decodeURIComponent(req.url.split('?')[0]);
  if (urlPath === '/') urlPath = '/README.md';
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

const server = http.createServer((req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    return res.end();
  }

  if (req.url.startsWith('/api/')) {
    return handleApi(req, res);
  }
  return serveStatic(req, res);
});

server.listen(PORT, HOST, () => {
  console.log(`[serve] HTTP listening on http://${HOST}:${PORT}`);
  console.log(`[serve] 加载项入口:  http://${HOST}:${PORT}/manifest.xml`);
  console.log(`[serve] Dialog 入口: http://${HOST}:${PORT}/src/dialog/dialog.html`);
  console.log(`[serve] API:`);
  console.log(`[serve]   GET  /api/health`);
  console.log(`[serve]   GET  /api/pptx-path`);
  console.log(`[serve]   POST /api/save-pptx?path=…`);
});

['SIGINT', 'SIGTERM'].forEach((sig) => {
  process.on(sig, () => {
    console.log(`[serve] received ${sig}, shutting down`);
    server.close(() => process.exit(0));
  });
});

