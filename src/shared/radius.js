/*
 * radius.js
 * 锁定信息管理 + R 角单位换算
 *
 * 锁定信息存在 localStorage，key = "slideNum|shapeId"
 * OOXML 解析/修改现在由 server 端（tools/serve.js）处理。
 */

const EMU_PER_CM = 360000;              // 厘米 -> EMU
const CM_PER_EMU = 1 / 360000;          // EMU -> 厘米
const ADJUSTMENT_MAX = 0.5;             // ratio 上限
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

function setLockEntry(slideNum, shapeId, entry) {
  const locks = loadLocks();
  const k = `${slideNum}|${shapeId}`;
  if (entry === null) {
    delete locks[k];
  } else {
    locks[k] = entry;
  }
  saveLocks(locks);
}

function getLockEntry(slideNum, shapeId) {
  return loadLocks()[`${slideNum}|${shapeId}`] || null;
}

// ---------------- 暴露到全局 ----------------

window.RadiusCore = {
  EMU_PER_CM,
  CM_PER_EMU,
  ADJUSTMENT_MAX,
  STORAGE_KEY,
  loadLocks,
  saveLocks,
  setLockEntry,
  getLockEntry,
};
