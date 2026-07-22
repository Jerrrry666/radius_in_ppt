/*
 * radius.js
 * OOXML 工具函数 + R 角单位换算
 *
 * 关键事实（PowerPoint 圆角矩形 OOXML）：
 *   - shape `<p:sp>` 里有 `<a:prstGeom prst="roundRect">`
 *   - R 角值在 `<a:gd name="adj" fmla="val X"/>`，X 是 0~50000（表示 0%~50%）
 *   - 形状尺寸在 `<a:ext cx="..." cy="..."/>`，单位是 EMU
 *   - 1 厘米 = 360000 EMU
 *   - 所以 绝对值(cm) = (X / 100000) * min(cx, cy) / 360000
 *
 * "锁定"语义：
 *   - 锁定信息存在 localStorage，key = (slideNum, shapeId)
 *   - 锁定时 R 角绝对值（厘米）固定；改变大小按比例重新计算
 */

const CM_PER_EMU = 1 / 360000;          // EMU -> 厘米
const EMU_PER_CM = 360000;              // 厘米 -> EMU
const ADJUSTMENT_MAX = 0.5;             // ratio 上限
const ADJUSTMENT_VAL_MAX = 50000;       // OOXML 中 val 的上限（表示 50%）
const STORAGE_KEY = 'radius_in_ppt_locks_v1';

// ---------------- 锁定文件读写 ----------------

function loadLocks() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (_) {
    return {};
  }
}

function saveLocks(locks) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(locks));
  } catch (_) {}
}

function getLockKey(slideNum, shapeId) {
  return `${slideNum}|${shapeId}`;
}

function setLockEntry(slideNum, shapeId, entry) {
  const locks = loadLocks();
  const k = getLockKey(slideNum, shapeId);
  if (entry === null) {
    delete locks[k];
  } else {
    locks[k] = entry;
  }
  saveLocks(locks);
}

function getLockEntry(slideNum, shapeId) {
  return loadLocks()[getLockKey(slideNum, shapeId)] || null;
}

// ---------------- 换算 ----------------

/**
 * 把厘米值转成 OOXML 的 val 整数（0 ~ 50000）
 * @param {number} cm
 * @param {number} shortSideEmu 形状短边（EMU）
 */
function cmToVal(cm, shortSideEmu) {
  if (shortSideEmu <= 0 || !Number.isFinite(cm)) return 0;
  const ratio = (cm * EMU_PER_CM) / shortSideEmu;
  if (ratio < 0) return 0;
  if (ratio > ADJUSTMENT_MAX) return ADJUSTMENT_VAL_MAX;
  return Math.round(ratio * 100000);
}

/**
 * 把 OOXML 的 val 转成厘米
 */
function valToCm(val, shortSideEmu) {
  if (shortSideEmu <= 0) return 0;
  const ratio = val / 100000;
  return ratio * shortSideEmu * CM_PER_EMU;
}

// ---------------- OOXML 解析 ----------------

/**
 * 从 .pptx ZIP buffer 解析所有圆角矩形。
 * @param {JSZip} zip
 * @returns {Array<{slideNum:number, shapeId:string, shapeName:string, shortSideEmu:number, currentVal:number, currentCm:number, slideFile:string}>}
 */
function parseRoundedRects(zip) {
  const results = [];
  // 找所有 slideN.xml
  const slideFiles = [];
  zip.forEach((path, file) => {
    const m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) {
      slideFiles.push({ slideNum: parseInt(m[1], 10), path });
    }
  });
  slideFiles.sort((a, b) => a.slideNum - b.slideNum);

  for (const { slideNum, path } of slideFiles) {
    const xml = zip.file(path).asText();
    const rects = parseRoundedRectsInSlideXml(xml, slideNum, path);
    results.push(...rects);
  }
  return results;
}

/**
 * 解析单个 slide XML 里的所有圆角矩形
 */
function parseRoundedRectsInSlideXml(xml, slideNum, slideFile) {
  const out = [];
  // 找所有 <p:sp>...</p:sp>
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
      if (r) {
        out.push({
          slideNum,
          slideFile,
          ...r,
        });
      }
    }
    idx = spEndClose;
  }
  return out;
}

function parseRoundedRect(spXml) {
  // 提取 id（cNvPr id="..."）
  const idMatch = spXml.match(/<p:cNvPr[^>]*\bid="(\d+)"/);
  if (!idMatch) return null;
  const id = idMatch[1];
  // 提取 name
  const nameMatch = spXml.match(/<p:cNvPr[^>]*\bname="([^"]*)"/);
  const name = nameMatch ? nameMatch[1] : '';
  // 提取 ext cx/cy（注意：可能有多个 <a:ext>，取 spPr 里的那个）
  const spPrMatch = spXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
  if (!spPrMatch) return null;
  const spPr = spPrMatch[1];
  const extMatch = spPr.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
  if (!extMatch) return null;
  const cx = parseInt(extMatch[1], 10);
  const cy = parseInt(extMatch[2], 10);
  const shortSideEmu = Math.min(cx, cy);
  // 提取 adj fmla val
  const fmlaMatch = spPr.match(/<a:gd[^>]*\bname="adj"[^>]*\bfmla="val\s+(\d+)"/)
                || spPr.match(/<a:gd[^>]*\bfmla="val\s+(\d+)"[^>]*\bname="adj"/);
  const currentVal = fmlaMatch ? parseInt(fmlaMatch[1], 10) : 0;
  const currentCm = valToCm(currentVal, shortSideEmu);
  return { id, name, shortSideEmu, currentVal, currentCm };
}

// ---------------- OOXML 修改 ----------------

/**
 * 修改指定 shape 的 adj val（直接字符串替换，简化版）
 * @param {string} slideXml 整个 slide XML
 * @param {string} shapeId 目标 shape id
 * @param {number} newVal 新的 val（0~50000）
 * @returns {string} 修改后的 slide XML
 */
function modifyShapeAdj(slideXml, shapeId, newVal) {
  // 找 <p:sp> ... </p:sp> 包含 id="shapeId" 的那个
  let idx = 0;
  while (true) {
    const spStart = slideXml.indexOf('<p:sp>', idx);
    if (spStart === -1) return slideXml;
    const spEnd = slideXml.indexOf('</p:sp>', spStart);
    if (spEnd === -1) return slideXml;
    const spEndClose = spEnd + '</p:sp>'.length;
    const spXml = slideXml.substring(spStart, spEndClose);
    if (spXml.includes(`id="${shapeId}"`) && spXml.includes('prst="roundRect"')) {
      // 找 prstGeom 里的 <a:gd name="adj" fmla="val X"/>
      const newSpXml = spXml.replace(
        /(<a:gd[^>]*\bname="adj"[^>]*\bfmla="val\s+)\d+(")/,
        `$1${newVal}$2`
      );
      return slideXml.substring(0, spStart) + newSpXml + slideXml.substring(spEndClose);
    }
    idx = spEndClose;
  }
}

// ---------------- 暴露到全局 ----------------

window.RadiusCore = {
  CM_PER_EMU,
  EMU_PER_CM,
  ADJUSTMENT_MAX,
  ADJUSTMENT_VAL_MAX,
  STORAGE_KEY,
  loadLocks,
  saveLocks,
  setLockEntry,
  getLockEntry,
  getLockKey,
  cmToVal,
  valToCm,
  parseRoundedRects,
  parseRoundedRectsInSlideXml,
  parseRoundedRect,
  modifyShapeAdj,
};
