#!/usr/bin/env node
/*
 * serve.js — 静态 HTTP server + .pptx 解析/写回的 API
 *
 * 监听 http://localhost:3000
 *
 * 静态文件 = 项目根（manifest.xml、src/、assets/）
 *
 * API:
 *   GET  /api/health             → {"ok":true}
 *   GET  /api/pptx-path          → 当前 .pptx 路径
 *   GET  /api/scan               → 扫描 .pptx，返回所有圆角矩形
 *   POST /api/apply              body: {shapes:[{slideNum,id,newVal}], path?}
 *                                → 改 XML + 写回
 *
 * 启动：npm start  或  node tools/serve.js
 */

const fs = require('fs');
const path = require('path');
const http = require('http');
const { execSync } = require('child_process');
const JSZip = require('jszip');

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

const EMU_PER_CM = 360000;
const ADJ_VAL_MAX = 50000; // 0~50000 对应 0%~50%

// ---------------- 拿 .pptx 路径 ----------------

function getCachedPptxPath() {
  try {
    const p = fs.readFileSync(PPTX_PATH_FILE, 'utf8').trim();
    if (p && fs.existsSync(p)) return p;
  } catch (_) {}
  return null;
}

function tryDetectPptxFromProcess() {
  try {
    const pids = execSync('pgrep -f "Microsoft PowerPoint" || true', { encoding: 'utf8' }).trim();
    if (!pids) return null;
    for (const pid of pids.split('\n')) {
      try {
        const lsofOut = execSync(
          `lsof -p ${pid} 2>/dev/null | awk '/\\.pptx$|\\.ppt$/ {print $NF; exit}' || true`,
          { encoding: 'utf8' }
        );
        const line = lsofOut.split('\n')[0].trim();
        if (line && fs.existsSync(line)) return line;
      } catch (_) {}
    }
  } catch (_) {}
  return null;
}

function resolvePptxPath(queryPath) {
  if (queryPath) return queryPath;
  return getCachedPptxPath() || tryDetectPptxFromProcess();
}

// ---------------- 解析 .pptx 里的圆角矩形 ----------------

async function scanRoundedRects(pptxPath) {
  const buf = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(buf, { stringFileName: false });
  const out = [];

  const slideFiles = [];
  zip.forEach((p, file) => {
    const m = p.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) slideFiles.push({ slideNum: parseInt(m[1], 10), path: p, file });
  });
  slideFiles.sort((a, b) => a.slideNum - b.slideNum);

  for (const { slideNum, path } of slideFiles) {
    const xml = await zip.file(path).async('string');
    const rects = parseRectsInSlideXml(xml);
    for (const r of rects) {
      out.push({ slideNum, slideFile: path, ...r });
    }
  }
  return out;
}

function parseRectsInSlideXml(xml) {
  const out = [];
  let idx = 0;
  while (true) {
    const spStart = xml.indexOf('<p:sp>', idx);
    if (spStart === -1) break;
    const spEnd = xml.indexOf('</p:sp>', spStart);
    if (spEnd === -1) break;
    const spEndClose = spEnd + '</p:sp>'.length;
    const spXml = xml.substring(spStart, spEndClose);
    if (spXml.includes('prst="roundRect"')) {
      const r = parseRoundedRect(spXml);
      if (r) out.push(r);
    }
    idx = spEndClose;
  }
  return out;
}

function parseRoundedRect(spXml) {
  const idMatch = spXml.match(/<p:cNvPr[^>]*\bid="(\d+)"/);
  if (!idMatch) return null;
  const id = idMatch[1];
  const nameMatch = spXml.match(/<p:cNvPr[^>]*\bname="([^"]*)"/);
  const name = nameMatch ? nameMatch[1] : '';
  const spPrMatch = spXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
  if (!spPrMatch) return null;
  const spPr = spPrMatch[1];
  const extMatch = spPr.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
  if (!extMatch) return null;
  const cx = parseInt(extMatch[1], 10);
  const cy = parseInt(extMatch[2], 10);
  const shortSideEmu = Math.min(cx, cy);
  const fmlaMatch = spPr.match(/<a:gd[^>]*\bname="adj"[^>]*\bfmla="val\s+(\d+)"/)
                || spPr.match(/<a:gd[^>]*\bfmla="val\s+(\d+)"[^>]*\bname="adj"/);
  const currentVal = fmlaMatch ? parseInt(fmlaMatch[1], 10) : 0;
  const currentCm = currentVal === 0 ? 0 : (currentVal / 100000) * shortSideEmu / EMU_PER_CM;
  return { id, name, shortSideEmu, currentVal, currentCm };
}

function cmToVal(cm, shortSideEmu) {
  if (shortSideEmu <= 0 || !Number.isFinite(cm)) return 0;
  const ratio = (cm * EMU_PER_CM) / shortSideEmu;
  if (ratio < 0) return 0;
  if (ratio > 0.5) return ADJ_VAL_MAX;
  return Math.round(ratio * 100000);
}

// ---------------- 修改 .pptx ----------------

async function applyRadius(pptxPath, items) {
  // items: [{slideNum, id, newVal}]
  const buf = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(buf, { stringFileName: false });

  // 按 slideFile 分组
  const byFile = {};
  for (const it of items) {
    // 先解析一次拿到 slideFile 路径
    const all = await scanRoundedRects(pptxPath);
    const found = all.find((s) => s.slideNum === it.slideNum && s.id === it.id);
    if (!found) continue;
    if (!byFile[found.slideFile]) byFile[found.slideFile] = [];
    byFile[found.slideFile].push({ id: it.id, newVal: it.newVal });
  }

  for (const [filePath, mods] of Object.entries(byFile)) {
    let xml = await zip.file(filePath).async('string');
    for (const m of mods) {
      xml = modifyShapeAdj(xml, m.id, m.newVal);
    }
    zip.file(filePath, xml);
  }

  const out = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  // 备份 + 写回
  try { fs.copyFileSync(pptxPath, pptxPath + '.bak'); } catch (_) {}
  fs.writeFileSync(pptxPath, out);
  return { ok: true, bytes: out.length, path: pptxPath, modified: items.length };
}

function modifyShapeAdj(slideXml, shapeId, newVal) {
  let idx = 0;
  while (true) {
    const spStart = slideXml.indexOf('<p:sp>', idx);
    if (spStart === -1) return slideXml;
    const spEnd = slideXml.indexOf('</p:sp>', spStart);
    if (spEnd === -1) return slideXml;
    const spEndClose = spEnd + '</p:sp>'.length;
    const spXml = slideXml.substring(spStart, spEndClose);
    if (spXml.includes(`id="${shapeId}"`) && spXml.includes('prst="roundRect"')) {
      const newSpXml = spXml.replace(
        /(<a:gd[^>]*\bname="adj"[^>]*\bfmla="val\s+)\d+(")/,
        `$1${newVal}$2`
      );
      return slideXml.substring(0, spStart) + newSpXml + slideXml.substring(spEndClose);
    }
    idx = spEndClose;
  }
}

// ---------------- HTTP ----------------

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      try {
        const txt = Buffer.concat(chunks).toString('utf8');
        resolve(txt ? JSON.parse(txt) : {});
      } catch (e) {
        reject(e);
      }
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

async function handleApi(req, res) {
  const url = req.url.split('?')[0];
  const query = {};
  (req.url.split('?')[1] || '').split('&').forEach((kv) => {
    if (!kv) return;
    const [k, v] = kv.split('=');
    if (k) query[decodeURIComponent(k)] = decodeURIComponent(v || '');
  });

  try {
    if (url === '/api/health') {
      return sendJSON(res, 200, { ok: true });
    }

    if (url === '/api/pptx-path') {
      const p = resolvePptxPath(query.path);
      if (!p) return sendJSON(res, 404, { error: 'no pptx detected; please launch R 角调整.app first' });
      return sendJSON(res, 200, { path: p });
    }

    if (url === '/api/scan') {
      const p = resolvePptxPath(query.path);
      if (!p) return sendJSON(res, 404, { error: 'no pptx detected' });
      const shapes = await scanRoundedRects(p);
      return sendJSON(res, 200, { shapes, path: p });
    }

    if (url === '/api/apply') {
      const body = await readJsonBody(req);
      const items = body.shapes || [];
      const p = resolvePptxPath(body.path);
      if (!p) return sendJSON(res, 404, { error: 'no pptx detected' });
      if (items.length === 0) return sendJSON(res, 400, { error: 'no items' });

      // 先扫一次拿到 shortSide（用于 cm → val 转换）
      const all = await scanRoundedRects(p);

      // 给每个 item 计算 newVal（如果传的是 cm）
      const finalItems = items.map((it) => {
        if (Number.isFinite(it.newVal)) return it;
        if (Number.isFinite(it.cm)) {
          const found = all.find((s) => s.slideNum === it.slideNum && s.id === it.id);
          if (!found) return null;
          return { slideNum: it.slideNum, id: it.id, newVal: cmToVal(it.cm, found.shortSideEmu) };
        }
        return null;
      }).filter(Boolean);

      if (finalItems.length === 0) {
        return sendJSON(res, 400, { error: 'no valid items' });
      }
      const result = await applyRadius(p, finalItems);
      return sendJSON(res, 200, result);
    }

    sendJSON(res, 404, { error: 'unknown api' });
  } catch (err) {
    console.error('[serve] API error:', err);
    sendJSON(res, 500, { error: err.message || String(err) });
  }
}

// 缓存：避免每次 /api/apply 都重新解析整个 .pptx
const shapesCache = new Map();

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
  console.log(`[serve]   GET  /api/scan`);
  console.log(`[serve]   POST /api/apply`);
});

['SIGINT', 'SIGTERM'].forEach((sig) => {
  process.on(sig, () => {
    console.log(`[serve] received ${sig}, shutting down`);
    server.close(() => process.exit(0));
  });
});

