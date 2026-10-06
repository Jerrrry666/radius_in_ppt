/*
 * radius-core.js — R 角调整 v1.2+ 实现层（业务逻辑，可被 mock 测）
 *
 * 这是 AGENTS.md 2.2 描述的**实现层**：
 *   - 所有 feature 函数（writeRadius / applyLayout / ...）都把 driver 作为第一参数
 *   - 通过 driver 间接访问 Office.js shape
 *   - 纯逻辑 + 业务判断（strict 拦截、lock 同步、padding 公式）
 *   - 零 Office.js import → mock 一个 driver 对象就能 100% 单元测试
 *
 * 模块边界：
 *   - dialog.js 是 UI 层（事件绑定 / 渲染 / 调 feature）
 *   - ppt-driver.js 是交互层（Office.js 薄封装）
 *   - radius-core 不知道任何 Office.js 概念
 */

const PT_PER_CM = 28.3464567;        // 1 cm = 28.3464567 pt
const ADJ_SCALE = 1;                 // Mac LTSC: adjustments.get(0).value 是 0~1 比例

// Tag keys（与 dialog.js 中的常量保持一致）
const LOCK_TAG_KEY = 'radiusLock_v1';
const LOCK_STRICT_TAG_KEY = 'radiusLockStrict_v1';
const LAYOUT_PARENT_TAG_KEY = 'layoutParent_v1';
const LAYOUT_CHILD_TAG_KEY = 'layoutChild_v1';

// PowerPoint 在 OOXML 中把 tag key 统一存成大写，但 mock / 旧内存对象可能保留
// caller 传入的驼峰写法。批量读取后的业务查找必须大小写不敏感。
function getBulkTagValue(tags, key) {
  if (!tags || key == null) return null;
  if (Object.prototype.hasOwnProperty.call(tags, key)) return tags[key];
  const wanted = String(key).toUpperCase();
  for (const actualKey of Object.keys(tags)) {
    if (String(actualKey).toUpperCase() === wanted) return tags[actualKey];
  }
  return null;
}

function lockStateFromTags(tags) {
  if (!tags || typeof tags !== 'object') throw new Error('Protection tags were not loaded');
  const raw = getBulkTagValue(tags, LOCK_TAG_KEY);
  const cm = raw == null || String(raw).trim() === '' ? NaN : Number(raw);
  const isLocked = Number.isFinite(cm) && cm >= 0;
  return { isLocked, lockedCm: isLocked ? cm : null,
    isStrict: getBulkTagValue(tags, LOCK_STRICT_TAG_KEY) === '1' };
}

function decideLockMonitorUpdate(shape, sample, previous) {
  const epsilon = 0.0001;
  const next = { width: sample.width, height: sample.height, adj: sample.adj,
    candidateAdj: sample.adj, stableCount: 0 };
  if (!shape.locked || !previous || previous.adj == null) return { action: 'none', next };
  const minSideCm = Math.min(sample.width, sample.height) / PT_PER_CM;
  const targetAdj = Math.min(shape.lockedCm, minSideCm / 2) / minSideCm;
  const sizeChanged = Math.abs(sample.width - previous.width) > 0.001 || Math.abs(sample.height - previous.height) > 0.001;
  if (sizeChanged) {
    next.adj = targetAdj;
    return { action: Math.abs(sample.adj - targetAdj) > epsilon ? 'restore' : 'none', targetAdj, next };
  }
  if (Math.abs(sample.adj - previous.adj) > epsilon) {
    next.adj = previous.adj;
    next.stableCount = Math.abs(sample.adj - previous.candidateAdj) <= epsilon ? (previous.stableCount || 0) + 1 : 1;
    if (next.stableCount < 4) return { action: 'none', next };
    next.stableCount = 0;
    next.adj = shape.strictLocked ? targetAdj : sample.adj;
    return { action: shape.strictLocked ? 'restore' : 'updateLock', targetAdj,
      lockedCm: sample.adj * minSideCm, next };
  }
  if (Math.abs(sample.adj - targetAdj) > epsilon) {
    next.adj = targetAdj;
    return { action: 'restore', targetAdj, next };
  }
  return { action: 'none', next };
}

async function reapplySelectionLocks(driver, ids, updates, legacyDefaults, opts) {
  let applied = 0;
  return withWritableShapes(driver, ids, async (shapes) => {
    const tags = await driver.loadTagsBulk(shapes);
    if (opts && opts.isCurrent && !opts.isCurrent()) throw new Error('stale-selection');
    for (const sh of shapes) {
      const id = driver.shapeId(sh);
      const state = lockStateFromTags(tags[id]);
      // Older brushes could leave strict without a fixed-radius tag. Capture
      // the radius observed when selected, then persist it through the same
      // safe transaction used to restore protected native edits.
      const legacy = legacyDefaults && legacyDefaults[id];
      if (state.isStrict && !state.isLocked && Number.isFinite(legacy) && legacy >= 0) {
        const r = await writeLockState(driver, sh, { lockedCm: legacy });
        if (!r.ok) throw new Error(r.error);
        state.isLocked = true;
        state.lockedCm = legacy;
      }
      if (!state.isLocked) continue;
      const newValue = updates && updates[id];
      if (!state.isStrict && Number.isFinite(newValue) && newValue >= 0) {
        const r = await writeLockState(driver, sh, { lockedCm: newValue });
        if (!r.ok) throw new Error(r.error);
        state.lockedCm = newValue;
      }
      const r = await reapplyLock(driver, sh, state.lockedCm);
      if (r.ok) applied++;
      else if (r.reason !== 'not-roundRect') throw new Error(r.error || r.reason);
    }
    await driver.sync();
    return { ok: true, applied, failed: 0 };
  }, opts);
}

// Ordinary radius/brush operations also use fresh top-level proxies whenever
// targets belong to a group. Never write into a transformed group descendant.
async function withWritableShapes(driver, ids, action, opts) {
  const checkCurrent = () => {
    if (opts && opts.isCurrent && !opts.isCurrent()) throw new Error('stale-selection');
  };
  if (ids.length === 0) return action([]);
  const slide = driver.activeSlide();
  const collection = driver.slideShapes(slide);
  let leaves = await driver.loadShapeTree(collection, 'id, name, width, height, adjustments, tags');
  let byId = new Map(leaves.map((sh) => [driver.shapeId(sh), sh]));
  const targets = ids.map((id) => byId.get(id));
  checkCurrent();
  if (targets.some((sh) => !sh)) throw new Error('目标形状在当前幻灯片找不到');
  const ordinaryIds = targets.filter((sh) => !driver.parentGroupOf(sh)).map((sh) => driver.shapeId(sh));
  const groups = new Map();
  for (const sh of targets) {
    let group = driver.parentGroupOf(sh);
    const seen = new Set();
    while (group) {
      const id = driver.shapeId(group);
      if (seen.has(id)) throw new Error('组合层级存在循环');
      seen.add(id);
      groups.set(id, group);
      group = driver.parentGroupOf(group);
    }
  }
  if (groups.size === 0) return action(targets);
  const snapshots = [];
  // Finish all metadata reads before the first structural mutation.
  const tags = await driver.loadTagsBulk(Array.from(groups.values()));
  for (const [id, group] of groups) {
    const parent = driver.parentGroupOf(group);
    let depth = 0, ancestor = parent;
    while (ancestor) { depth++; ancestor = driver.parentGroupOf(ancestor); }
    snapshots.push({ id, group, members: driver.groupShapes(group).map((s) => driver.shapeId(s)),
      parentId: parent ? driver.shapeId(parent) : null, depth,
      name: driver.shapeName(group), tags: tags[id], ungrouped: false });
  }
  snapshots.sort((a, b) => a.depth - b.depth);
  const findNode = (items, id) => {
    for (const node of items) {
      if (driver.shapeId(node) === id) return node;
      if (driver.isGroup(node)) {
        const found = findNode(driver.groupShapes(node), id);
        if (found) return found;
      }
    }
    return null;
  };
  let result, operationError;
  try {
    checkCurrent();
    for (const snapshot of snapshots) {
      await driver.loadShapeTree(collection, 'id, name, width, height, adjustments, tags');
      const freshGroup = findNode(collection.items, snapshot.id);
      if (!freshGroup) throw new Error('无法找回组合');
      checkCurrent();
      driver.ungroupShapeGroup(freshGroup);
      snapshot.ungrouped = true;
      await driver.sync();
    }
    leaves = await driver.loadShapeTree(collection, 'id, name, width, height, adjustments, tags');
    byId = new Map(leaves.map((sh) => [driver.shapeId(sh), sh]));
    const fresh = ids.map((id) => byId.get(id));
    if (fresh.some((sh) => !sh)) throw new Error('解除组合后无法找回目标形状');
    checkCurrent();
    result = await action(fresh);
  } catch (e) { operationError = e; }
  const restoredIds = [];
  const replacements = new Map();
  const restoreErrors = [];
  for (const snapshot of snapshots.slice().reverse()) {
    if (!snapshot.ungrouped) continue;
    try {
      const group = driver.addGroup(collection, snapshot.members.map((id) => replacements.get(id) || id));
      driver.load(group, 'id');
      await driver.sync();
      const id = driver.shapeId(group);
      replacements.set(snapshot.id, id);
      if (snapshot.parentId == null) restoredIds.push(id);
      // Commit the structure before restoring metadata. A name/tag failure
      // must not leave the outer group referring to a removed inner group ID.
      if (snapshot.name) {
        try { driver.setShapeName(group, snapshot.name); }
        catch (e) { restoreErrors.push(e.message || String(e)); }
      }
      for (const [key, value] of Object.entries(snapshot.tags || {})) {
        try { driver.addTag(group, key, value); }
        catch (e) { restoreErrors.push(e.message || String(e)); }
      }
      await driver.sync();
    } catch (e) { restoreErrors.push(e.message || String(e)); }
  }
  if (restoredIds.length) {
    try {
      driver.selectShapes(slide, restoredIds.concat(ordinaryIds));
      await driver.sync();
    } catch (e) { restoreErrors.push(e.message || String(e)); }
  }
  if (restoreErrors.length) {
    throw new Error((operationError ? (operationError.message || String(operationError)) + '；' : '') +
      '组合恢复失败：' + restoreErrors.join('; '));
  }
  if (operationError) throw operationError;
  return result;
}

async function applyRadiusToSelection(driver, shapes, cm, opts) {
  opts = opts || {};
  let applied = 0, failed = 0, lockedSynced = 0;
  try {
    if (!Number.isFinite(cm) || cm < 0) throw new Error('Invalid radius');
    const targets = shapes.filter((sh) => driver.isRoundRect(sh));
    const tags = await driver.loadTagsBulk(targets);
    if (targets.some((sh) => lockStateFromTags(tags[driver.shapeId(sh)]).isStrict)) {
      return { ok: false, applied, failed, lockedSynced, reason: 'strict', error: '选区中有形状启用了防误触' };
    }
    if (opts.isCurrent && !opts.isCurrent()) return { ok: false, reason: 'stale-selection', applied, failed, lockedSynced };
    await withWritableShapes(driver, targets.map((sh) => driver.shapeId(sh)), async (fresh) => {
      if (opts.isCurrent && !opts.isCurrent()) throw new Error('stale-selection');
      for (const sh of fresh) {
        const r = await writeRadius(driver, sh, cm, { knownLockState: lockStateFromTags(tags[driver.shapeId(sh)]) });
        if (r.ok) { applied++; if (r.wasLocked) lockedSynced++; }
        else failed++;
      }
      await driver.sync();
    }, opts);
    return { ok: true, applied, failed, lockedSynced };
  } catch (e) {
    const error = e.message || String(e);
    console.log('[applyRadiusToSelection] EXCEPTION:', error);
    return { ok: false, applied, failed, lockedSynced, error };
  }
}

// ---------------- 布局 math ----------------

/**
 * Legacy single-inset padding bound, retained for existing pure callers.
 * It applies only to 1×1 with no gutter. applyLayout no longer uses it:
 * user centimetre spacing is preserved in every mode, and radius is clamped.
 *
 * @param {number} parentWidthCm - 父宽 (cm)
 * @param {number} parentHeightCm - 父高 (cm)
 * @param {number} parentRcm - 父 R 角 (cm)
 * @param {number} dInitCm - 用户初始设定的 padding (cm)
 * @returns {Object} { effectivePaddingCm, dMaxCm, clamped }
 *   - effectivePaddingCm: 实际 padding（可能 < dInitCm）
 *   - dMaxCm: d 的上限（保证子 R 角不 clamp）
 *   - clamped: true 表示 dInitCm > dMaxCm，padding 被自动减小
 */
function computeAutoPadding(parentWidthCm, parentHeightCm, parentRcm, dInitCm) {
  const minSideCm = Math.min(parentWidthCm, parentHeightCm);
  const dMaxCm = minSideCm / 2 - parentRcm;
  if (!Number.isFinite(dMaxCm) || dMaxCm <= 0) {
    // R 父 >= min/2：d_max < 0，clamp padding 到 0
    return { effectivePaddingCm: 0, dMaxCm, clamped: true };
  }
  if (dInitCm > dMaxCm) {
    return { effectivePaddingCm: dMaxCm, dMaxCm, clamped: true };
  }
  return { effectivePaddingCm: Math.max(0, dInitCm), dMaxCm, clamped: false };
}

/**
 * 纯函数：给定父 box + rows/cols/padding/gutter，算出子形状的尺寸 + 位置
 * @param {Object} parent - { left, top, width, height } (pt)
 * @param {number} rows - 行数 (1-5)
 * @param {number} cols - 列数 (1-5)
 * @param {number} paddingCm - 边距 (cm)
 * @param {number} gutterCm - 间距 (cm)
 * @returns {Object} { subW, subH, positions, feasible, reason }
 *   - positions: [{ left, top, w, h, idx }] (pt)，row-major 排（i*cols + j）
 *   - feasible: false 表示 padding/gutter 太大，子尺寸 ≤ 0
 */
function computeLayout(parent, rows, cols, paddingCm, gutterCm) {
  const paddingPt = paddingCm * PT_PER_CM;
  const gutterPt = gutterCm * PT_PER_CM;
  const totalW = parent.width - 2 * paddingPt - (cols - 1) * gutterPt;
  const totalH = parent.height - 2 * paddingPt - (rows - 1) * gutterPt;
  if (totalW <= 0 || totalH <= 0) {
    return { subW: 0, subH: 0, positions: [], feasible: false, reason: 'padding/gutter 太大，挤不下' };
  }
  if (rows < 1 || cols < 1) {
    return { subW: 0, subH: 0, positions: [], feasible: false, reason: '行/列必须 ≥ 1' };
  }
  const subW = totalW / cols;
  const subH = totalH / rows;
  const positions = [];
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      positions.push({
        left: parent.left + paddingPt + j * (subW + gutterPt),
        top: parent.top + paddingPt + i * (subH + gutterPt),
        w: subW,
        h: subH,
        idx: i * cols + j,
      });
    }
  }
  return { subW, subH, positions, feasible: true, reason: '' };
}

/**
 * v1.2.11：行/列互斥联动（行 × 列 = 子数 N）
 *
 * 用户拖 rows 时 cols 自动 = max(1, ceil(N / rows))，反之亦然。
 * 解决"2×2 布局改 rows=1 → 期望 1×4 而不是 1×2"的 bug。
 *
 * @param {string} changed - 哪个维度被用户改了，'rows' 或 'cols'
 * @param {number} value - 用户改的那个值
 * @param {number} N - 子形状总数（layout 创建后固定）
 * @returns {Object} { rows, cols } 联动后的两个值（都 ≥ 1，rows × cols ≥ N）
 *
 * 例子（N=4）：
 *   changed='rows', value=1 → { rows:1, cols:4 }
 *   changed='rows', value=3 → { rows:3, cols:2 }（4/3 上取整 = 2）
 *   changed='cols', value=3 → { rows:2, cols:3 }
 *   changed='rows', value=10 → clamp 到 4 → { rows:4, cols:1 }
 *   N=0（防御）→ { rows:1, cols:1 }
 */
function computeGridCoupledRowsCols(changed, value, N) {
  // 兼容字符串输入（slider / number input 的 value 都是 string）
  // 之前用 Number.isFinite(value) → "2" 算 false → fallback 到 1，导致 2×2 改到 1×4 后改不回去
  const safeN = (() => {
    const n = Number(N);
    return Number.isFinite(n) && n > 0 ? Math.floor(n) : 1;
  })();
  const numValue = Number(value);
  let v = Number.isFinite(numValue) ? Math.floor(numValue) : 1;
  // clamp 到 [1, N]
  v = Math.max(1, Math.min(safeN, v));
  const otherVal = Math.max(1, Math.ceil(safeN / v));
  return changed === 'rows'
    ? { rows: v, cols: otherVal }
    : { rows: otherVal, cols: v };
}

/**
 * v1.2.13：算出 rows 的「合法可取值列表」= N 的所有正因子
 *
 * 用户要求：行滑块不应该是连续范围，而是离散列表。
 * 例：N=4 → rows 可选 [1, 2, 4]（不含 3，因为 3×2=6>4 会留 2 个空位）
 *     N=6 → rows 可选 [1, 2, 3, 6]
 *     N=12 → rows 可选 [1, 2, 3, 4, 6, 12]
 *
 * 推导：rows × cols = N 严格成立（不留空位）→ cols = N/rows 必须是整数
 *       → rows 必须是 N 的因子
 *
 * 用途：
 *   - UI：renderLayoutPanel 把这些值设到 datalist 的 <option>，slider 显示 tick
 *   - input 事件：snap 用户输入到最近的合法值（防"拖到 3"）
 *   - 单测：保证行为可预测
 *
 * @param {number} N - 子数
 * @returns {number[]} 升序的 N 的正因子列表
 *   - N=1 → [1]
 *   - N=0 / 非正数 → [1]（防御：默认 1×1）
 */
function computeGridFactors(N) {
  const safeN = (() => {
    const n = Number(N);
    return Number.isFinite(n) && n > 0 ? Math.floor(n) : 1;
  })();
  const factors = [];
  for (let i = 1; i * i <= safeN; i++) {
    if (safeN % i === 0) {
      factors.push(i);
      if (i !== safeN / i) factors.push(safeN / i);
    }
  }
  factors.sort((a, b) => a - b);
  return factors;
}

/**
 * v1.2.13：把任意 v snap 到 N 的最近因子
 *
 * 用途：slider 拖动时（连续值）snap 到最近的合法离散值
 *   - N=4，factors=[1,2,4]，snap(3) → 2（3 离 2 比离 4 近）
 *   - snap 边界：snap(1) → 1，snap(4) → 4
 *   - 越界（snap(0) / snap(99)）→ clamp 到 [1, lastFactor]
 *
 * @param {number} v - 任意正数（字符串会 Number() 转换）
 * @param {number[]} factors - N 的因子列表（升序）
 * @returns {number} 最近的因子
 */
function snapToNearestGridFactor(v, factors) {
  if (!Array.isArray(factors) || factors.length === 0) return 1;
  const num = Number(v);
  if (!Number.isFinite(num)) return factors[0];
  const sorted = factors.slice().sort((a, b) => a - b);
  const clamped = Math.max(sorted[0], Math.min(sorted[sorted.length - 1], num));
  // 找最近的
  let best = sorted[0];
  let bestDist = Math.abs(clamped - best);
  for (const f of sorted) {
    const d = Math.abs(clamped - f);
    if (d < bestDist) {
      best = f;
      bestDist = d;
    }
  }
  return best;
}

// ---------------- 单位换算 ----------------

/**
 * 输入值按单位换算到 cm
 * @param {number} val - 输入值
 * @param {string} unit - 'cm' | '%'
 * @param {number} refMinSideCm - % 模式参考的形状短边（cm）
 * @returns {number} cm
 */
function valueToCm(val, unit, refMinSideCm) {
  if (unit === '%') {
    return (val / 100) * refMinSideCm;
  }
  return val;
}

/**
 * cm 按单位换算到显示值
 * @param {number} cm
 * @param {string} unit - 'cm' | '%'
 * @param {number} refMinSideCm
 * @returns {number} 显示值
 */
function cmToValue(cm, unit, refMinSideCm) {
  if (unit === '%') {
    if (refMinSideCm <= 0) return 0;
    return (cm / refMinSideCm) * 100;
  }
  return cm;
}

// ---------------- R 角联动公式 ----------------

/**
 * 计算子 R 角（按 linkRMode）
 * @param {number} parentRcm - 父 R 角（cm）
 * @param {number} paddingCm - 边距（cm）
 * @param {string} linkRMode - 'subtract' | 'same' | 'off'
 * @returns {number} subRcm（off 时返回 0）
 */
function computeLinkedSubR(parentRcm, paddingCm, linkRMode) {
  if (linkRMode === 'off') return 0;
  if (linkRMode === 'same') return parentRcm;
  // subtract（v1.0 公式）：r = max(0, 父 R − 边距)
  return Math.max(0, parentRcm - paddingCm);
}

/**
 * 计算 adj 值（0~1 比例，给 PowerPoint set 用）
 * @param {number} subRcm - 子 R 角（cm）
 * @param {number} childMinSideCm - 子短边（cm）
 * @returns {number} adj 0~1
 */
function cmToAdj(subRcm, childMinSideCm) {
  if (childMinSideCm <= 0) return 0;
  return subRcm / childMinSideCm * ADJ_SCALE;
}

/**
 * clamp：把 R 角限制到不超过子短边一半（PowerPoint 几何约束）
 * @param {number} targetCm
 * @param {number} minSideCm
 * @returns {number} clampedCm
 */
function clampRadius(targetCm, minSideCm) {
  if (minSideCm <= 0) return Math.max(0, targetCm);
  return Math.min(targetCm, minSideCm / 2);
}

/**
 * 写 R 角的完整流程：按 linkRMode 算 subR → clamp → adj 计算
 * @param {number} childMinSideCm - 子短边（cm）
 * @param {string} linkRMode - 'subtract' | 'same' | 'off'
 * @param {number} parentRcm - 父 R 角（cm）
 * @param {number} paddingCm - 边距（cm）
 * @returns {Object} { finalCm, adj, skipped }
 *   - skipped: true 如果 linkRMode = 'off'
 */
function computeFinalRadius(childMinSideCm, linkRMode, parentRcm, paddingCm) {
  if (linkRMode === 'off') return { finalCm: 0, adj: 0, skipped: true };
  // 按 linkRMode 算 subR（v1.0 公式：r = max(0, 父R − 边距)）
  const subRcm = computeLinkedSubR(parentRcm, paddingCm, linkRMode);
  // clamp：不超过子短边一半
  const finalCm = clampRadius(subRcm, childMinSideCm);
  // adj 转换
  const adj = cmToAdj(finalCm, childMinSideCm);
  return { finalCm, adj, skipped: false };
}

// ---------------- 写 R 角的行为规则（纯函数） ----------------

/**
 * pushHistory 纯函数：去重 + 移到最前 + 限 MAX_HISTORY 条
 * 跟 dialog.js v1.0 userHistory 行为完全一致（v1.3.6 抽到 radius-core）
 * @param {Array} history - 当前 history 列表
 * @param {number} value - 新的 R 角值
 * @param {string} unit - 'cm' | '%'
 * @param {number} maxLen - 默认 5
 * @returns {Array} 新 history 列表（不修改入参）
 */
function pushHistory(history, value, unit, maxLen) {
  const limit = Number.isFinite(maxLen) ? maxLen : 5;
  const filtered = (history || []).filter((h) => !(h.value === value && h.unit === unit));
  filtered.unshift({ value, unit, ts: Date.now() });
  return filtered.slice(0, limit);
}

/**
 * 决定是否应该拒绝写 R 角
 * 模拟 writeRadius 的 strict 拦截逻辑（供 dialog.js 在 PowerPoint.run 之前做第一道防线）
 * @param {Object} shape - { isStrict, isRoundRect, minSideCm }
 * @returns {Object} { allow, reason }
 *   - allow: false + reason='strict' → 拒绝
 *   - allow: false + reason='no-size' → 拒绝
 *   - allow: false + reason='not-roundRect' → 拒绝
 *   - allow: true → 可以写
 */
function shouldRejectWriteRadius(shape) {
  // 防误触：永远拦截（最高优先级）
  if (shape.isStrict) {
    return { allow: false, reason: 'strict', message: '🔒 此形状启用了防误触，必须用户手动关闭' };
  }
  if (!shape.minSideCm || shape.minSideCm <= 0) {
    return { allow: false, reason: 'no-size' };
  }
  if (!shape.isRoundRect) {
    return { allow: false, reason: 'not-roundRect' };
  }
  return { allow: true };
}

/**
 * 决定 onApply 是否应该「全部拒绝」（选区里有任何 strict → 拒绝）
 * @param {Array} selectedShapes
 * @returns {Object} { shouldReject, strictCount }
 */
function shouldRejectOnApply(selectedShapes) {
  const strictCount = selectedShapes.filter((s) => s.isRoundRect && s.isStrict).length;
  if (strictCount > 0) {
    return { shouldReject: true, strictCount };
  }
  return { shouldReject: false, strictCount: 0 };
}

/**
 * 决定 applyLayoutToChildren 是否应该「拒绝整个 apply」（选区子有 strict）
 * @param {Array} selectedShapes - 全部选区
 * @param {string} parentId - 父 id
 * @param {Array} childIds - 子 id 列表
 * @returns {Object} { shouldReject, strictShapes }
 */
function shouldRejectLayoutApply(selectedShapes, parentId, childIds) {
  const strictShapes = selectedShapes.filter((s) =>
    s.layoutRole !== 'parent' && childIds.indexOf(s.id) >= 0 && s.isStrict
  );
  if (strictShapes.length > 0) {
    return { shouldReject: true, strictShapes };
  }
  return { shouldReject: false, strictShapes: [] };
}

/**
 * 检测哪些 layout 父的 R 角变了（用于 monitorTick → syncLayoutChildrenRIfNeeded 联动）
 *
 * 场景：用户在 PPT 里**直接拖父的 R 角黄色滑块**（不走 task pane 的 onApply），
 *       monitorTick 检测到 currentCm 变了，但不会主动同步子 R 角。
 *       这个函数告诉 caller「哪些 layout 父的 R 角相对上次记的 knownCm 变了」，
 *       caller 拿到结果后调 syncLayoutChildrenRIfNeeded() 同步子 R 角。
 *
 * 行为：
 *   - 遍历 selectedShapes，找 layoutRole === 'parent' 的形状
 *   - 对每个父，比较 currentCm vs knownCmMap[id]（容差 0.01 cm）
 *   - currentCm 为 null/undefined 的（还没读到 R 角的）→ 跳过
 *   - knownCmMap[id] 为 null/undefined 的（首次见到）→ 算"变了"，让 caller 触发首次同步
 *
 * 为什么不放 dialog.js：
 *   - 纯函数，零副作用（不读 driver / 不写任何状态）
 *   - 单测覆盖"哪些算变了"的边界（NaN、null、容差边界、首次）
 *   - dialog.js 只负责维护 knownCmMap + 调 syncLayoutChildrenRIfNeeded
 *
 * @param {Object} knownCmMap - { [shapeId]: lastKnownCm }（dialog.js 维护）
 * @param {Array} selectedShapes - dialog.js 内存里的 selectedShapes
 *   - 每项至少含 { id, layoutRole, currentCm }
 * @returns {Array<{parentId, lastCm, newCm}>} 变了哪些父（empty = 没变）
 */
function detectLayoutParentChanges(knownCmMap, selectedShapes) {
  const changes = [];
  if (!Array.isArray(selectedShapes)) return changes;
  for (const s of selectedShapes) {
    if (!s || s.layoutRole !== 'parent') continue;
    if (s.currentCm == null) continue;
    const newCm = s.currentCm;
    const lastCm = knownCmMap && knownCmMap[s.id] != null ? knownCmMap[s.id] : null;
    // 首次见到（lastCm null）→ 算"变了"（caller 可以选择忽略，或立即同步一次）
    // 浮点容差：monitorTick 自己用 ADJ_EPSILON=0.0001 做 adj 比较，但 adj 换 cm 后误差是 0.0001 * minSideCm
    //   对最小 1cm 形状误差 = 0.0001cm，10cm 形状 = 0.001cm —— 1e-3 cm 是合理阈值
    if (lastCm == null) {
      changes.push({ parentId: s.id, lastCm: null, newCm });
      continue;
    }
    if (!Number.isFinite(newCm) || !Number.isFinite(lastCm)) continue;
    if (Math.abs(newCm - lastCm) > 1e-3) {
      changes.push({ parentId: s.id, lastCm, newCm });
    }
  }
  return changes;
}

/**
 * v1.2.9：检测哪些 layout 父的 size (width/height) 变了（cm 单位）
 *
 * 场景：用户在 PPT 里**直接拖父的边缘**改 width/height（不走 task pane 的 onApply），
 *       子图形的 size 应该按 layout 公式 (subW = (W - 2d - (cols-1)*g) / cols) 重算。
 *       monitorTick 检测到 size 变化，触发 applyLayoutToChildren → 重写子位置/尺寸/R 角。
 *
 * 跟 detectLayoutParentChanges 的关系：
 *   - detectLayoutParentChanges 看 R 角变化（触发 syncR）
 *   - detectLayoutParentSizeChanges 看 size 变化（触发重算 layout 几何 + syncR）
 *   - 两个并行检测，任一 fire 都触发 scheduleParentRSync → applyLayoutToChildren
 *   - applyLayoutToChildren 已经会同时重算子位置/尺寸 + R 角，所以 R/size 触发同一路径
 *
 * 为什么不合并到 detectLayoutParentChanges：
 *   - R 角 map 和 size map 是两套独立状态（lastCm / lastSize），caller 维护成本不同
 *   - 合并会让函数返回类型变复杂、caller 判断逻辑变长
 *   - 两个独立函数 + caller merge「任一变了就 fire」更清晰
 *
 * 容差：
 *   - 1e-3 cm = 0.001 cm = 0.028 pt
 *   - monitorTick PPT 读 width/height 走 driver.size，ppt 浮点可能有 0.01 pt 误差
 *   - 1e-3 cm 是合理阈值（< 1 像素 / 96 dpi）
 *
 * @param {Object} knownSizeMap - { [parentId]: { widthCm, heightCm } }（dialog.js 维护）
 * @param {Array} selectedShapes - dialog.js 内存里的 selectedShapes
 *   - 每项至少含 { id, layoutRole, widthCm, heightCm }
 *   - 兼容 width/height 字段（pt 单位，自动除 PT_PER_CM）
 * @returns {Array<{parentId, lastSize, newSize}>} 变了哪些父
 *   - lastSize 为 null 表示首次见到（caller 决定是否 fire）
 */
function detectLayoutParentSizeChanges(knownSizeMap, selectedShapes) {
  const changes = [];
  if (!Array.isArray(selectedShapes)) return changes;
  for (const s of selectedShapes) {
    if (!s || s.layoutRole !== 'parent') continue;
    // 兼容 widthCm/heightCm（cm）和 width/height（pt）两种入参
    const wCm = Number.isFinite(s.widthCm)
      ? s.widthCm
      : (Number.isFinite(s.width) ? s.width / PT_PER_CM : NaN);
    const hCm = Number.isFinite(s.heightCm)
      ? s.heightCm
      : (Number.isFinite(s.height) ? s.height / PT_PER_CM : NaN);
    if (!Number.isFinite(wCm) || !Number.isFinite(hCm)) continue;
    const newSize = { widthCm: wCm, heightCm: hCm };
    const lastSize = knownSizeMap && knownSizeMap[s.id] ? knownSizeMap[s.id] : null;
    if (!lastSize) {
      // 首次见到：算"变了"（caller 可以选择忽略）
      changes.push({ parentId: s.id, lastSize: null, newSize });
      continue;
    }
    if (
      !Number.isFinite(lastSize.widthCm) ||
      !Number.isFinite(lastSize.heightCm)
    ) {
      // lastSize 损坏：当成首次
      changes.push({ parentId: s.id, lastSize: null, newSize });
      continue;
    }
    if (
      Math.abs(newSize.widthCm - lastSize.widthCm) > 1e-3 ||
      Math.abs(newSize.heightCm - lastSize.heightCm) > 1e-3
    ) {
      changes.push({ parentId: s.id, lastSize, newSize });
    }
  }
  return changes;
}

/**
 * 模拟写 R 角后：locked 形状的 fixed value 同步逻辑
 * @param {Object} shape - { isLocked, lockedCm }
 * @param {number} newCm - 写完后的 R 角（cm）
 * @returns {Object} { newLockedCm, synced }
 */
function syncFixedValueIfLocked(shape, newCm) {
  if (shape.isLocked) {
    return { newLockedCm: newCm, synced: true };
  }
  return { newLockedCm: shape.lockedCm || 0, synced: false };
}

// ---------------- 统一写 R 角函数（driver 版，Office.js 上下文） ----------------

/**
 * 写 R 角的 driver 版：操作真实的 Office.js shape proxy
 *
 * 通过 driver 间接访问 shape 属性（driver.size / driver.isRoundRect / driver.setAdjFraction / driver.readTag / driver.addTag）
 *   - 不直接 import Office.js
 *   - 跟 mock 测的纯函数版返回相同形状的 { ok, newCm, wasLocked, wasStrict, reason, error }
 *
 * 行为：
 *   1. 读 lock + strict（通过 driver.readTag）
 *   2. strict 永远拦截（最高优先级）
 *   3. clamp + 写 R 角（通过 driver.setAdjFraction）
 *   4. 如果 locked → 同步 fixed value（通过 driver.addTag）
 *   5. 如果 layoutParentId → 写子 tag
 *
 * @param {Object} driver - createDriver(ctx) 返回的 driver
 * @param {Object} shape - shape proxy（已 load 完所有需要的字段：id, width, height, adjustments, tags）
 * @param {number} targetCm - 目标 R 角（cm）
 * @param {Object} [opts] - { layoutParentId, clamp }
 * @returns {Promise<{ok, newCm?, wasLocked, wasStrict, reason?, error?}>}
 */
async function writeRadius(driver, shape, targetCm, opts) {
  opts = opts || {};
  const layoutParentId = opts.layoutParentId;
  const clamp = opts.clamp !== false;
  try {
    // 1. 读 lock + strict（优先用 caller 传的 knownLockState，避免 per-shape readTag + sync 在 for 循环内累积）
    //    Mac LTSC 实测坑：v1.3.6 修 #6 之前，syncLayoutChildrenR 调 readTag 4 次，每次都 await ctx.sync()，
    //                    第 3/4 个 shape 的 setAdjFraction 在真实 PPT 上会丢（mock 不模拟得到）
    //    修法：caller 用 driver.loadTagsBulk 一次 load + sync 拿全部 → 传 knownLockState
    //          → 跳过 per-shape readTag
    let isLocked = false;
    let lockedCm = 0;
    let isStrict = false;
    if (opts.knownLockState && typeof opts.knownLockState === 'object') {
      isLocked = !!opts.knownLockState.isLocked;
      lockedCm = Number.isFinite(opts.knownLockState.lockedCm) ? opts.knownLockState.lockedCm : 0;
      isStrict = !!opts.knownLockState.isStrict;
    } else {
      // 走 driver.readTag（per-shape sync，可能在 for 循环内累积 — v1.2.6 Mac LTSC 实测坑）
      const lockVal = await driver.readTag(shape, LOCK_TAG_KEY);
      if (lockVal) {
        const cm = parseFloat(lockVal);
        if (Number.isFinite(cm) && cm >= 0) {
          isLocked = true;
          lockedCm = cm;
        }
      }
      const strictVal = await driver.readTag(shape, LOCK_STRICT_TAG_KEY);
      isStrict = strictVal === '1';
    }

    // 2. strict 永远拦截
    if (isStrict) {
      return { ok: false, reason: 'strict', isStrict: true, wasLocked: isLocked };
    }

    // 3. clamp + 写 R 角
    if (!driver.isRoundRect(shape)) {
      return { ok: false, reason: 'not-roundRect', wasLocked: isLocked, wasStrict: false };
    }
    // 用 driver.size（只要 width/height），不要 driver.box（还要 left/top）——
    // 写 R 角只需要短边做 clamp，caller 可以只 load width/height 省掉 left/top
    const size = driver.size(shape);
    const minSideCm = Math.min(size.width, size.height) / PT_PER_CM;
    if (minSideCm <= 0) {
      return { ok: false, reason: 'no-size', wasLocked: isLocked, wasStrict: false };
    }
    // v1.3.5 修：Infinity/NaN 不能让 clamp 静默吞掉，提前 reject
    // - Math.min(Infinity, 30) = 30，clamp 会把 Infinity 当成有限值处理
    // - Math.min(NaN, 30) = NaN，虽然下面 !Number.isFinite(newAdj) 会兜住，但语义上更早 reject 更明确
    if (!Number.isFinite(targetCm)) {
      return { ok: false, reason: 'invalid-adj', wasLocked: isLocked, wasStrict: false };
    }
    let newCm = clamp ? Math.min(targetCm, minSideCm / 2) : targetCm;
    if (newCm < 0) newCm = 0;
    const newAdj = (newCm / minSideCm) * ADJ_SCALE;
    if (!Number.isFinite(newAdj)) {
      return { ok: false, reason: 'invalid-adj', wasLocked: isLocked, wasStrict: false };
    }
    driver.setAdjFraction(shape, newAdj);

    // 4. 同步 fixed value（如果 locked）— 用 driver.addTag（不需要额外 load）
    if (isLocked) {
      driver.addTag(shape, LOCK_TAG_KEY, String(newCm));
    }

    // 5. 写子 tag
    if (layoutParentId) {
      driver.addTag(shape, LAYOUT_CHILD_TAG_KEY, layoutParentId);
    }

    return { ok: true, newCm, wasLocked: isLocked, wasStrict: false, lockedCm };
  } catch (e) {
    // 把异常 message 主动 log 出来，免得以后被外层 caller 当成 reason='exception' 一吞了之
    const msg = e && e.message ? e.message : String(e);
    const stack = e && e.stack ? e.stack : '';
    if (typeof console !== 'undefined') {
      console.log('[writeRadius] EXCEPTION:', msg, '| stack:', stack, '| targetCm:', targetCm);
    }
    return { ok: false, reason: 'exception', error: msg, wasLocked: false, wasStrict: false };
  }
}

// ---------------- 读/写 shape 的 lock + strict 状态（driver 版） ----------------

/**
 * 读一个 shape 的 lock + strict 状态
 *
 * @param {Object} driver
 * @param {Object} shape - shape proxy（必须先 load 'items/tags' 才能 readTag）
 * @returns {Promise<{lockedCm: number|null, isStrict: boolean}>}
 *   - lockedCm: number 解析后的 cm 值；null = 没 lock tag
 *   - isStrict: true = 防误触开启
 *   - 任一 tag 不存在都返回默认（null/false），不 throw
 */
async function readLockState(driver, shape) {
  let lockedCm = null;
  let isStrict = false;
  const lockVal = await driver.readTag(shape, LOCK_TAG_KEY);
  if (lockVal != null) {
    const cm = parseFloat(lockVal);
    if (Number.isFinite(cm) && cm >= 0) lockedCm = cm;
  }
  const strictVal = await driver.readTag(shape, LOCK_STRICT_TAG_KEY);
  if (strictVal === '1') isStrict = true;
  return { lockedCm, isStrict };
}

/**
 * 写一个 shape 的 lock + strict 状态
 *
 * 语义（跟原 updateLockTagForShape 一致）：
 *   - lockedCm: number 写 / null 删 / undefined 不动
 *   - isStrict: true 写 '1' / false 删 / null/undefined 不动
 *
 * @param {Object} driver
 * @param {Object} shape
 * @param {Object} state - { lockedCm, isStrict }
 * @returns {Promise<{ok: boolean, error?: string}>}
 */
async function writeLockState(driver, shape, state) {
  state = state || {};
  try {
    if (state.lockedCm !== undefined) {
      if (state.lockedCm == null) {
        driver.deleteTag(shape, LOCK_TAG_KEY);
      } else {
        if (!Number.isFinite(state.lockedCm) || state.lockedCm < 0) throw new Error('Invalid fixed radius');
        driver.addTag(shape, LOCK_TAG_KEY, String(state.lockedCm));
      }
    }
    if (state.isStrict === true) {
      driver.addTag(shape, LOCK_STRICT_TAG_KEY, '1');
    } else if (state.isStrict === false) {
      driver.deleteTag(shape, LOCK_STRICT_TAG_KEY);
    }
    return { ok: true };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log('[writeLockState] EXCEPTION:', msg);
    return { ok: false, error: msg };
  }
}

/**
 * 重新应用 lock：按当前形状大小反算 adj = lockedCm（clamp 到短边一半）
 *
 * 给「使用数值固定 R 角 - 重新应用」按钮用：被 PPT 内编辑改了之后点这个恢复
 *
 * @param {Object} driver
 * @param {Object} shape - shape proxy
 * @param {number} lockedCm - 要反算回的 cm 值
 * @returns {Promise<{ok, newCm, reason?}>}
 */
async function reapplyLock(driver, shape, lockedCm) {
  try {
    if (!Number.isFinite(lockedCm) || lockedCm < 0) return { ok: false, reason: 'invalid-target' };
    if (!driver.isRoundRect(shape)) return { ok: false, reason: 'not-roundRect' };
    const size = driver.size(shape);
    const minSideCm = Math.min(size.width, size.height) / PT_PER_CM;
    if (minSideCm <= 0) return { ok: false, reason: 'no-size' };
    const newCm = Math.min(lockedCm, minSideCm / 2);
    if (newCm < 0) return { ok: false, reason: 'invalid-target' };
    const newAdj = (newCm / minSideCm) * ADJ_SCALE;
    if (!Number.isFinite(newAdj)) return { ok: false, reason: 'invalid-adj' };
    driver.setAdjFraction(shape, newAdj);
    return { ok: true, newCm };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    return { ok: false, reason: 'exception', error: msg };
  }
}

// ---------------- applyLayout driver 版（Office.js 上下文） ----------------

/**
 * 应用 layout 到父形状的子形状们（driver 版，端到端在真实 PPT 跑）
 *
 * 行为：
 *   1. 在当前 slide（getSelectedSlides().getItemAt(0)）找父 + 子
 *   2. 集合层 load 所有需要的字段（id, left, top, width, height, adjustments, tags）
 *      —— per-shape load 在 Mac LTSC 不 work，4.4.1 坑
 *   3. 算 layout（computeLayout）：子尺寸 + 位置
 *   4. 第一道防线：进 PowerPoint.run 之前 caller 必须检查 strict（这里不查 selectedShapes）
 *   5. 第二道防线：writeRadius 内部会查 strict tag，命中跳过
 *   6. 写每个子的位置 + 尺寸
 *   7. 写每个子的 R 角（按 linkRMode 公式）— 走 writeRadius，lock 同步 fixed value
 *   8. 写每个子的 child tag（LAYOUT_CHILD_TAG_KEY）
 *   9. 写父 tag（LAYOUT_PARENT_TAG_KEY），**过滤掉不在当前 slide 的 stale childIds**
 *
 * @param {Object} driver
 * @param {string} parentId
 * @param {Object} params - { rows, cols, padding, gutter, linkRMode }
 * @param {Array} childIds
 * @param {Object} [opts] - { writeParentTag, syncR, writeGeometry }
 * @returns {Promise<{ok, applied, failed, warn, strictOverridden, lockedCount, error?}>}
 */
async function applyLayout(driver, parentId, params, childIds, opts) {
  opts = opts || {};
  if (!params || !Number.isInteger(params.rows) || params.rows < 1 || params.rows > 5 ||
      !Number.isInteger(params.cols) || params.cols < 1 || params.cols > 5 ||
      !Number.isFinite(params.padding) || params.padding < 0 ||
      !Number.isFinite(params.gutter) || params.gutter < 0 || !Array.isArray(childIds) ||
      new Set(childIds).size !== childIds.length || childIds.includes(parentId)) {
    return { ok: false, applied: 0, failed: 0, error: '布局参数不合法' };
  }
  const writeParentTag = opts.writeParentTag !== false;
  const syncR = opts.syncR !== false;
  // 切换 linkRMode 只应改变子 R 和持久化 tag；不能顺带把已由用户缩放的
  // group 子位置/尺寸重新按 layout 公式“反算”一遍。
  const writeGeometry = opts.writeGeometry !== false;
  // Preserve the existing default and stored mode names. For a single inset
  // rectangle, subtract gives concentric arcs when parentR >= padding.
  const linkRMode = params.linkRMode || 'same';
  const expectedCount = params.rows * params.cols;

  let applied = 0;
  let failed = 0;
  let strictOverridden = 0;
  let lockedCount = 0;
  let warn = '';
  let lockedChildCm = [];
  let regroupState = null;

  // Mac LTSC 对已缩放 group 的后代直接写 box / adjustment 都不稳定。
  // 如果本次布局临时解组了，成功和异常路径都必须尽力恢复原 group。
  const restoreGroupIfNeeded = async () => {
    if (!regroupState || regroupState.restored || regroupState.restoreAttempted) return null;
    regroupState.restoreAttempted = true;
    const newGroup = driver.addGroup(regroupState.shapeCollection, regroupState.memberIds);
    driver.load(newGroup, 'id');
    await driver.sync();
    regroupState.restored = true;
    const restoreErrors = [];
    if (regroupState.name) {
      try { driver.setShapeName(newGroup, regroupState.name); }
      catch (e) { restoreErrors.push(e.message || String(e)); }
    }
    for (const [key, value] of Object.entries(regroupState.tags || {})) {
      try { driver.addTag(newGroup, key, value); }
      catch (e) { restoreErrors.push(e.message || String(e)); }
    }
    try { await driver.sync(); }
    catch (e) { restoreErrors.push(e.message || String(e)); }
    console.log('[applyLayout/driver] GROUP-TXN regroup done members=', regroupState.memberIds.length);
    // addGroup 不保证保留原选区。主动选中新 group，避免布局面板在事务后消失。
    try {
      const newGroupId = driver.shapeId(newGroup);
      driver.selectShapes(regroupState.slide, [newGroupId]);
      await driver.sync();
      console.log('[applyLayout/driver] GROUP-TXN selection restored groupId=', newGroupId);
    } catch (selectionError) {
      restoreErrors.push(selectionError.message || String(selectionError));
    }
    if (restoreErrors.length) throw new Error('组合名称、tag或选区恢复失败：' + restoreErrors.join('; '));
    return newGroup;
  };

  try {
    // 1. 当前 slide + shape tree collection-level load
    // Mac LTSC 不能把 group path 无条件 load 到普通 shape；driver 会先读 type，
    // 再只展开真实 Group 的 shapes collection。
    const slide = driver.activeSlide();
    const slideShapeCollection = driver.slideShapes(slide);
    let slideLeaves = await driver.loadShapeTree(
      slideShapeCollection,
      'id, name, left, top, width, height, level, adjustments, tags'
    );

    // 2. 建 id → shape 映射
    let idToShape = new Map();
    const rebuildIdMap = () => {
      idToShape = new Map();
      for (const sh of slideLeaves) {
        const id = driver.shapeId(sh);
        if (id != null) idToShape.set(id, sh);
      }
    };
    rebuildIdMap();

    // 3. 找父
    let parentSh = idToShape.get(parentId);
    if (!parentSh) {
      warn = '父形状在当前 slide 找不到（可能选了别的页）';
      console.log('[applyLayout/driver] WARN parent not found in current slide');
      return { ok: false, applied, failed, warn };
    }

    // 4. 收集存在的子（过滤掉 stale / 不在当前 slide 的）
    const validChildIds = [];
    let childShapes = [];
    for (let k = 0; k < expectedCount; k++) {
      const cid = childIds[k];
      const csh = idToShape.get(cid);
      if (!csh) {
        console.log('[applyLayout/driver] skip missing child id=', cid, '(stale)');
        continue;  // 跳过 stale（不在当前 slide / 已被删）
      }
      validChildIds.push(cid);
      childShapes.push(csh);
    }
    if (validChildIds.length < expectedCount) {
      warn = `子形状不足（需要 ${expectedCount}，找到 ${validChildIds.length}）`;
      console.log('[applyLayout/driver] WARN', warn);
      return { ok: false, applied, failed, warn };
    }

    // All actual targets, including children outside the selection, must be
    // checked before ungrouping, geometry writes, radius writes, or tag writes.
    let tagsById = await driver.loadTagsBulk(childShapes);
    if (childShapes.some((sh) => lockStateFromTags(tagsById[driver.shapeId(sh)]).isStrict)) {
      return { ok: false, applied, failed, reason: 'strict', warn: '布局子形状启用了防误触，请先关闭防误触' };
    }
    if (opts.isCurrent && !opts.isCurrent()) throw new Error('stale-selection');

    // 5. group 安全事务：
    //    已缩放 group 内直接改后代，哪怕逐子 sync 也会让 group transform 损坏。
    //    仅支持父 + 全部 layout 子是同一个「顶层 group 的直接成员」：
    //    临时 ungroup → 顶层坐标写布局/R → 按原全部成员 regroup。
    const layoutShapes = [parentSh, ...childShapes];
    const directGroups = layoutShapes.map((sh) =>
      typeof driver.parentGroupOf === 'function' ? driver.parentGroupOf(sh) : null
    );
    const groupedCount = directGroups.filter(Boolean).length;
    if (groupedCount > 0) {
      const commonGroup = directGroups[0];
      const commonGroupId = commonGroup ? driver.shapeId(commonGroup) : null;
      const allSameDirectGroup = !!commonGroup && directGroups.every((g) =>
        !!g && driver.shapeId(g) === commonGroupId
      );
      const groupIsTopLevel = allSameDirectGroup &&
        (!driver.parentGroupOf || !driver.parentGroupOf(commonGroup));

      if (!allSameDirectGroup || !groupIsTopLevel) {
        warn = '布局父/子位于不同或嵌套组合中；为避免 PowerPoint 损坏组合，请先解除嵌套/混合组合';
        console.log('[applyLayout/driver] WARN unsafe grouped layout:', warn);
        return { ok: false, applied, failed, warn };
      }

      const memberIds = driver.groupShapes(commonGroup)
        .map((sh) => driver.shapeId(sh))
        .filter((id) => id != null);
      if (memberIds.length < 2) {
        warn = '组合成员不足，无法安全重建组合';
        console.log('[applyLayout/driver] WARN', warn);
        return { ok: false, applied, failed, warn };
      }

      const groupName = driver.shapeName(commonGroup) || '';
      const groupTagsById = await driver.loadTagsBulk([commonGroup]);
      if (opts.isCurrent && !opts.isCurrent()) throw new Error('stale-selection');
      regroupState = {
        slide,
        shapeCollection: slideShapeCollection,
        memberIds,
        name: groupName,
        tags: groupTagsById[commonGroupId] || {},
        restored: false,
        restoreAttempted: false,
      };

      console.log('[applyLayout/driver] GROUP-TXN ungroup start members=', memberIds.length);
      driver.ungroupShapeGroup(commonGroup);
      await driver.sync();

      // ungroup 后旧的 group scoped proxy 不再可靠；重新从 slide 顶层按 id 获取。
      slideLeaves = await driver.loadShapeTree(
        slideShapeCollection,
        'id, name, left, top, width, height, level, adjustments, tags'
      );
      rebuildIdMap();
      parentSh = idToShape.get(parentId);
      childShapes = validChildIds.map((id) => idToShape.get(id));
      if (!parentSh || childShapes.some((sh) => !sh)) {
        throw new Error('解除组合后无法按原 id 找回布局父/子');
      }
      console.log('[applyLayout/driver] GROUP-TXN ungroup done; fresh top-level proxies ready');
    }

    // 6. 父 R 角（v1.0 per-shape get(0) + sync + 读）
    let parentRcm = 0;
    if (syncR && linkRMode !== 'off' && driver.isRoundRect(parentSh)) {
      parentRcm = await driver.readAdjFraction(parentSh) *
        Math.min(driver.size(parentSh).width, driver.size(parentSh).height) / PT_PER_CM;
    }

    // 7. 只有几何参数变化时才读父 box、算 layout。
    // linkRMode 切换走 R-only，不碰任何子 box。
    let layout = null;
    if (writeGeometry) {
      const parentBox = driver.box(parentSh);
      console.log('[applyLayout/driver] parent box:', JSON.stringify(parentBox), 'Rcm=', parentRcm);

      // Centimetre spacing is explicit. Only radius is clamped to the short
      // side; never silently change the requested padding to fit a radius.
      layout = computeLayout(parentBox, params.rows, params.cols, params.padding, params.gutter);
      if (!layout.feasible) {
        warn = layout.reason;
        await restoreGroupIfNeeded();
        return { ok: false, applied, failed, warn };
      }
    } else {
      console.log('[applyLayout/driver] R-ONLY: skip parent box + all child geometry');
    }

    // 8. 写每个子的位置 + 尺寸 + R 角 + child tag
    // 一次 loadTagsBulk 加载全部 TagCollection key/value，避免 per-call readTag + sync。
    tagsById = await driver.loadTagsBulk(childShapes);
    if (opts.isCurrent && !opts.isCurrent()) throw new Error('stale-selection');
    if (childShapes.some((sh) => lockStateFromTags(tagsById[driver.shapeId(sh)]).isStrict)) {
      throw new Error('布局子形状启用了防误触，请先关闭防误触');
    }
    for (let k = 0; k < childShapes.length; k++) {
      const csh = childShapes[k];
      if (writeGeometry) {
        const pos = layout.positions[k];
        if (!pos) continue;
        try {
          console.log(`[applyLayout/driver] write child #${k} (id=${driver.shapeId(csh)}) pos=`, JSON.stringify(pos));
          driver.setBox(csh, { left: pos.left, top: pos.top, width: pos.w, height: pos.h });
        } catch (e) {
          const msg = e && e.message ? e.message : String(e);
          console.log(`[applyLayout/driver] write fail #${k} (id=${driver.shapeId(csh)}):`, msg);
          throw new Error(`写子 #${k} 位置/尺寸失败: ${msg}`);
        }
      }
      // 写 R 角（走 writeRadius，第二道防线实时查 strict + 同步 lock fixed value）
      if (syncR && linkRMode !== 'off') {
        const subRcm = linkRMode === 'same' ? parentRcm : Math.max(0, parentRcm - params.padding);
        console.log(`[applyLayout/driver] R link #${k}: parentRcm=${parentRcm}, mode=${linkRMode}, padding=${params.padding}, target subRcm=${subRcm}`);
        // 从 readTagsBulk 拿的 tag 构造 knownLockState，传给 writeRadius（跳过 per-call readTag + sync）
        const r = await writeRadius(driver, csh, subRcm, {
          layoutParentId: parentId,
          knownLockState: lockStateFromTags(tagsById[driver.shapeId(csh)]),
        });
        if (r.ok) {
          console.log(`[applyLayout/driver] R link #${k}: written subRcm=${r.newCm}, wasLocked=${r.wasLocked}`);
          if (r.wasLocked) {
            lockedCount++;
            lockedChildCm.push({ id: driver.shapeId(csh), newCm: r.newCm });
          }
        } else {
          console.log(`[applyLayout/driver] R link #${k}: skipped, reason=${r.reason}${r.error ? ' error=' + r.error : ''}`);
          if (r.reason !== 'not-roundRect') throw new Error(r.error || r.reason);
        }
      }
      // 写 child tag
      driver.addTag(csh, LAYOUT_CHILD_TAG_KEY, parentId);
      applied++;
    }
    console.log('[applyLayout/driver] applied=', applied, 'failed=', failed);

    // 9. 写父 tag（**用 validChildIds 过滤后的版本**，stale childIds 自动清理）
    if (writeParentTag) {
      try {
        const payload = JSON.stringify({
          rows: params.rows,
          cols: params.cols,
          padding: params.padding,
          gutter: params.gutter,
          linkRMode,
          childIds: validChildIds,  // ← 关键：不是 caller 传的 childIds，是过滤后的
        });
        driver.addTag(parentSh, LAYOUT_PARENT_TAG_KEY, payload);
      } catch (e) {
        console.log('[applyLayout/driver] write parent tag fail:', e.message || e);
        throw e;
      }
    }
    await driver.sync();
    console.log('[applyLayout/driver] sync done, lockedChildCm count=', lockedChildCm.length);
    await restoreGroupIfNeeded();
    return {
      ok: true,
      applied,
      failed,
      warn,
      strictOverridden,
      lockedCount,
      regrouped: !!regroupState,
      geometryWritten: writeGeometry,
    };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log('[applyLayout/driver] OUTER ERROR:', msg);
    if (regroupState && !regroupState.restored && !regroupState.restoreAttempted) {
      try {
        await restoreGroupIfNeeded();
      } catch (restoreError) {
        const restoreMsg = restoreError && restoreError.message
          ? restoreError.message
          : String(restoreError);
        console.log('[applyLayout/driver] GROUP-TXN restore ERROR:', restoreMsg);
        return {
          ok: false,
          applied,
          failed,
          warn,
          error: `${msg}；组合恢复失败：${restoreMsg}`,
        };
      }
    }
    return { ok: false, applied, failed, warn, error: msg };
  }
}

// ---------------- syncLayoutChildrenR driver 版（Office.js 上下文） ----------------

/**
 * 同步 layout 子的 R 角（driver 版）
 *
 * 行为：
 *   1. 集合层 load slide shapes（id, width, height, adjustments, tags）
 *   2. 对每个 childId 找 shape（过滤掉 stale / 不在当前 slide 的）
 *   3. 按 linkRMode 算子 R = 父 R（same）或 父 R - padding（subtract）
 *   4. 调 writeRadius 写——自动处理 strict 拦截 + lock 同步 fixed value
 *
 * @param {Object} driver
 * @param {string} parentId
 * @param {Array} childIds
 * @param {number} paddingCm
 * @param {string} linkRMode - 'same' | 'subtract' | 'off'
 * @param {number} parentRcm
 * @returns {Promise<{ok, applied, failed, error?}>}
 */
async function syncLayoutChildrenR(driver, parentId, childIds, paddingCm, linkRMode, parentRcm) {
  if (linkRMode === 'off') return { ok: true, applied: 0, failed: 0 };
  if (!Number.isFinite(parentRcm) || parentRcm < 0) return { ok: false, applied: 0, failed: 0, error: 'Invalid parent radius' };
  let applied = 0;
  let failed = 0;
  try {
    // v1.3.6 修 #6 子 bug：4 个子只写 2 个（用户实测：调整父 R 角后上面 2 个子变了下面 2 个没变）
    // 根因：writeRadius 内部 readTag 调 ctx.sync()，4 次 readTag + 4 次 setAdjFraction 在同一个 PowerPoint.run
    //       Mac LTSC 上 per-shape sync 累积（v1.2.6 同样坑），后几个 shape 的 setAdjFraction 失败/丢失
    // 修法：sibling applyLayout 模式 —— 一次 load + sync 拿全部 tag（用 readTagsBulk 避开 per-call sync），
    //       写所有 setAdjFraction，final sync 一次
    const slide = driver.activeSlide();
    const slideShapes = await driver.loadShapeTree(
      driver.slideShapes(slide),
      'id, width, height, adjustments, tags'
    );
    const idToShape = new Map();
    for (const sh of slideShapes) {
      const id = driver.shapeId(sh);
      if (id != null) idToShape.set(id, sh);
    }
    // 一次 load + sync 拿全部 TagCollection key/value，避免 per-call sync 累积。
    const tagsById = await driver.loadTagsBulk(slideShapes);
    const targets = childIds.map((id) => idToShape.get(id)).filter(Boolean);
    await withWritableShapes(driver, targets.map((sh) => driver.shapeId(sh)), async (fresh) => {
      for (const child of fresh) {
        const id = driver.shapeId(child);
        const target = linkRMode === 'same' ? parentRcm : Math.max(0, parentRcm - paddingCm);
        const r = await writeRadius(driver, child, target, {
          knownLockState: lockStateFromTags(tagsById[id]),
        });
        if (r.ok) applied++;
        else if (!['strict', 'not-roundRect', 'no-size'].includes(r.reason)) failed++;
      }
      await driver.sync();
    });
    console.log(`[syncLayoutChildrenR/driver] done: applied=${applied} failed=${failed} childIds=${JSON.stringify(childIds)}`);
    return { ok: true, applied, failed };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log('[syncLayoutChildrenR/driver] OUTER ERROR:', msg);
    return { ok: false, applied, failed, error: msg };
  }
}

// ---------------- loadLayoutTags driver 版（read + stale 检测） ----------------

/**
 * 解析父 tag value（JSON 字符串）→ 结构化对象
 * 解析失败 / 字段缺失 → 返回 null（caller 应该跳过）
 * @param {string} tagValue
 * @returns {Object|null} { rows, cols, padding, gutter, linkRMode, childIds } | null
 */
function parseLayoutParentTagValue(tagValue) {
  if (typeof tagValue !== 'string' || tagValue.length === 0) return null;
  try {
    const obj = JSON.parse(tagValue);
    if (!obj || !Number.isInteger(obj.rows) || obj.rows < 1 || obj.rows > 5 ||
        !Number.isInteger(obj.cols) || obj.cols < 1 || obj.cols > 5 || !Array.isArray(obj.childIds) ||
        (Number.isFinite(obj.padding) && obj.padding < 0) ||
        (Number.isFinite(obj.gutter) && obj.gutter < 0)) {
      return null;
    }
    return {
      rows: obj.rows,
      cols: obj.cols,
      padding: Number.isFinite(obj.padding) ? obj.padding : 0,
      gutter: Number.isFinite(obj.gutter) ? obj.gutter : 0,
      // 兼容旧版 linkR（boolean），v1.2 改用 linkRMode（'subtract' | 'same' | 'off'）
      linkRMode: ['subtract', 'same', 'off'].includes(obj.linkRMode)
        ? obj.linkRMode
        : (obj.linkR === false ? 'off' : 'subtract'),
      childIds: obj.childIds.filter((x) => typeof x === 'string' && x.length > 0),
    };
  } catch (_) {
    return null;
  }
}

/**
 * 解析父 tag 后过滤出 stale childIds（在 selectedShapeIds 集合里找不到的）
 * 用于 refreshSelection 修 #6：父 tag 里 childIds 可能包含已被删 / 跨 slide 的子
 *
 * @param {Object} parsedTag - parseLayoutParentTagValue 返回的对象
 * @param {Set} selectedShapeIds - 当前 slide 选区里所有 shape id（含父子）
 * @returns {Object} { validChildIds: string[], staleChildIds: string[] }
 */
function detectStaleChildrenInLayout(parsedTag, selectedShapeIds) {
  if (!parsedTag || !Array.isArray(parsedTag.childIds)) {
    return { validChildIds: [], staleChildIds: [] };
  }
  const valid = [];
  const stale = [];
  for (const cid of parsedTag.childIds) {
    if (selectedShapeIds.has(cid)) valid.push(cid);
    else stale.push(cid);
  }
  return { validChildIds: valid, staleChildIds: stale };
}

/**
 * 读 layout tag（driver 版）—— 给 refreshSelection 用
 *
 * 行为：
 *   1. 对每个 shape，集合层读 layoutParent_v1 + layoutChild_v1 tag
 *   2. 解析父 tag → parents[id] = {rows, cols, padding, gutter, linkRMode, childIds}
 *   3. 读子 tag → childOf[id] = parentId
 *   4. 顺便过滤 stale childIds（在整个 slide 上找不到的）→ staleParents[id] = [staleChildId, ...]
 *
 * 注意：
 *   - driver.readTag 在 tag 不存在时返回 null（不 throw），所以 catch 块不会进
 *   - 选区变化时用 getSelectedShapes() 调用，传入 shapes 给这个函数即可
 *   - **stale 检测必须用整 slide 的 shape IDs，不能用选区**（v1.3.7 修 bug：只选父时
 *     子不在选区 → 旧版本误判为 stale → childIds 全被过滤掉 → UI 显示"子 0 个"）
 *
 * @param {Object} driver
 * @param {Array} selectedShapes - shape proxy 列表（已经 load 过 'items/id'）
 * @param {Array} [allSlideShapes] - **整个 slide** 的 shape 列表（已 load 'items/id'）
 *                                   用于 stale 检测；不传则 fallback 到 selectedShapeIds
 *                                   （fallback 仅保留测试兼容，生产 dialog.js 必传）
 * @returns {Promise<{
 *     ok: boolean,
 *     parents: Object,    // { shapeId: {rows, cols, padding, gutter, linkRMode, childIds} }
 *     childOf: Object,    // { shapeId: parentId }
 *     staleParents: Object,  // { parentShapeId: [staleChildId, ...] } —— 父 tag 里有但 slide 上找不到的子
 *     error?: string
 *   }>}
 */
async function loadLayoutTags(driver, selectedShapes, allSlideShapes) {
  const parents = {};
  const childOf = {};
  const staleParents = {};
  try {
    // 先收 shape ids（用于 stale 检测）
    const selectedShapeIds = new Set();
    const shapesList = [];
    if (selectedShapes && typeof selectedShapes.items !== 'undefined') {
      // 集合对象（Office.js proxy）—— 取 items
      for (const sh of selectedShapes.items) {
        shapesList.push(sh);
        if (sh.id != null) selectedShapeIds.add(sh.id);
      }
    } else if (Array.isArray(selectedShapes)) {
      // 数组
      for (const sh of selectedShapes) {
        shapesList.push(sh);
        if (sh && sh.id != null) selectedShapeIds.add(sh.id);
      }
    } else {
      return { ok: false, parents, childOf, staleParents, error: 'selectedShapes 必须有 items 或为数组' };
    }

    // v1.3.7 修 bug：stale 检测用整 slide 的 shape IDs，不用选区
    // （只选父时，4 个子不在选区 → 旧版误判全部 stale → childIds 变空）
    let slideShapeIds = selectedShapeIds;  // fallback
    if (allSlideShapes) {
      slideShapeIds = new Set();
      const slideList = (allSlideShapes && typeof allSlideShapes.items !== 'undefined')
        ? allSlideShapes.items
        : (Array.isArray(allSlideShapes) ? allSlideShapes : []);
      for (const sh of slideList) {
        if (sh && sh.id != null) slideShapeIds.add(sh.id);
      }
    }

    const tagsById = await driver.loadTagsBulk(shapesList);
    // 读每个 shape 的 layout tag
    for (const sh of shapesList) {
      const sid = sh.id;
      if (sid == null) continue;
      // 父 tag
      const parentVal = getBulkTagValue(tagsById[sid], LAYOUT_PARENT_TAG_KEY);
      if (parentVal) {
        const parsed = parseLayoutParentTagValue(parentVal);
        if (parsed) {
          parents[sid] = parsed;
          // stale 检测：父 tag 里的 childIds 不在整 slide 的 shape ids 里（v1.3.7 之前是选区 → 误判）
          const { validChildIds, staleChildIds } = detectStaleChildrenInLayout(parsed, slideShapeIds);
          if (staleChildIds.length > 0) {
            // 写回 parents[sid].childIds 只保留 valid 的（caller 拿到的是过滤后的）
            parents[sid].childIds = validChildIds;
            staleParents[sid] = staleChildIds;
          }
        }
      }
      // 子 tag
      const childVal = getBulkTagValue(tagsById[sid], LAYOUT_CHILD_TAG_KEY);
      if (typeof childVal === 'string' && childVal.length > 0) {
        childOf[sid] = childVal;
      }
    }
    return { ok: true, parents, childOf, staleParents };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    return { ok: false, parents, childOf, staleParents, error: msg };
  }
}

// ---------------- saveLayoutTags driver 版（写父 + 子 tag） ----------------

/**
 * 写 layout tag（driver 版）—— 给 dialog.js 的 saveLayoutTags 用
 *
 * 行为：
 *   1. 在 slide 里找父 + 子
 *   2. 写父 tag（LAYOUT_PARENT_TAG_KEY）= JSON.stringify({rows, cols, padding, gutter, linkRMode, childIds})
 *   3. 给每个存在的子写 tag（LAYOUT_CHILD_TAG_KEY）= parentId
 *   4. stale childIds 会被自动跳过（不在当前 slide 的子不写 tag）
 *
 * @param {Object} driver
 * @param {Object} slide - slide proxy
 * @param {string} parentId
 * @param {Object} params - { rows, cols, padding, gutter, linkRMode }
 * @param {Array} childIds
 * @returns {Promise<{ok, error?, writtenChildIds?, staleChildIds?}>}
 */
async function saveLayoutTags(driver, slide, parentId, params, childIds, opts) {
  try {
    // 集合层递归 load slide shape tree（id only）
    const slideShapesArr = await driver.loadShapeTree(driver.slideShapes(slide), 'id');

    // 建 id → shape 映射
    const idToShape = new Map();
    for (const sh of slideShapesArr) {
      if (sh.id != null) idToShape.set(sh.id, sh);
    }

    // 找父
    const parentSh = idToShape.get(parentId);
    if (!parentSh) {
      return { ok: false, error: '父形状在当前 slide 找不到', writtenChildIds: [], staleChildIds: [] };
    }

    // 过滤掉 stale childIds
    const validChildIds = [];
    const staleChildIds = [];
    for (const cid of childIds) {
      if (idToShape.has(cid)) validChildIds.push(cid);
      else staleChildIds.push(cid);
    }

    // 写父 tag
    const payload = JSON.stringify({
      rows: params.rows,
      cols: params.cols,
      padding: Number.isFinite(params.padding) ? params.padding : 0,
      gutter: Number.isFinite(params.gutter) ? params.gutter : 0,
      linkRMode: ['subtract', 'same', 'off'].includes(params.linkRMode) ? params.linkRMode : 'same',
      childIds: validChildIds,
    });
    await withWritableShapes(driver, [parentId, ...validChildIds], async (fresh) => {
      for (const shape of fresh) {
        const id = driver.shapeId(shape);
        if (id === parentId) driver.addTag(shape, LAYOUT_PARENT_TAG_KEY, payload);
        else driver.addTag(shape, LAYOUT_CHILD_TAG_KEY, parentId);
      }
      await driver.sync();
    }, opts);

    return { ok: true, writtenChildIds: validChildIds, staleChildIds };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    return { ok: false, error: msg, writtenChildIds: [], staleChildIds: [] };
  }
}

// ---------------- pickupFromSelection / applyPickedToSelection driver 版 ----------------

/**
 * 读选区里第一个圆角矩形的 R 角 + strict 状态（driver 版）
 *
 * 跟 dialog.js v1.0/v1.1 pickupFromSelection 行为一致（v1.3.6 抽到 radius-core）
 *
 * 行为：
 *   1. 遍历 selectedShapes
 *   2. 找第一个几何类型已确认为roundRect的形状
 *   3. get(0) 存变量 → sync → 读 value → 算 cm = value * minSideCm
 *   4. 读 strict tag（如果有）
 *   5. 返回 { id, name, cm, sourceStrict }，没找到圆角矩形返回 null
 *
 * @param {Object} driver
 * @param {Array} selectedShapes - shape proxy 列表（已 load 'items/id, items/name, items/width, items/height, items/adjustments, items/tags'）
 * @returns {Promise<{id, name, cm, sourceStrict} | null>}
 */
async function pickupFromSelection(driver, selectedShapes) {
  try {
    const shapesList = [];
    if (selectedShapes && typeof selectedShapes.items !== 'undefined') {
      for (const sh of selectedShapes.items) shapesList.push(sh);
    } else if (Array.isArray(selectedShapes)) {
      for (const sh of selectedShapes) shapesList.push(sh);
    } else {
      return null;
    }

    for (const sh of shapesList) {
      try {
        if (!driver.isRoundRect(sh)) continue;
        // Mac LTSC 模式：get(0) 存变量 → sync → 读 value
        const v = await driver.readAdjFraction(sh);
        if (!Number.isFinite(v)) continue;
        const size = driver.size(sh);
        const minSideCm = Math.min(size.width, size.height) / PT_PER_CM;
        const cm = v * minSideCm;
        // 读 strict
        const sourceStrict = (await driver.readTag(sh, LOCK_STRICT_TAG_KEY)) === '1';
        return {
          id: driver.shapeId(sh),
          name: driver.shapeName(sh) || '(未命名)',
          cm,
          sourceStrict,
        };
      } catch (e) {
        console.log('[pickupFromSelection] shape read failed:', e.message || e);
        throw e;
      }
    }
    return null;
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log('[pickupFromSelection] EXCEPTION:', msg);
    return null;
  }
}

/**
 * 把 pickup 出来的 R 角应用到选区里所有 roundRect（driver 版）
 *
 * 跟 dialog.js v1.0/v1.1 applyPipetteToSelection 行为一致（v1.3.6 抽到 radius-core）
 * 修 #1 bug：之前 dialog.js 调的是旧版 writeRadiusToShape（直接用 ctxShape API，不走 driver），
 *            在 Mac LTSC 某些场景会刷不进去。改用 radius-core.writeRadius（driver 版 + 走 setAdjFraction 路径）后稳定。
 *
 * 已开启防误触的目标始终阻断整个操作，必须用户先手动关闭。
 * syncStrict 只允许在写入成功后，把源的开启状态应用到未保护目标；
 * 源未开启防误触时，不删除目标保护tag，不通过改tag绕过writeRadius。
 *
 * 行为：
 *   1. 选区里有 strict 形状 → 全部拒绝
 *   2. 对未保护roundRect调writeRadius（自动处理clamp及lock同步）
 *   3. syncStrict=true且source.strict=true → 保存实际固定值并开启防误触
 *
 * @param {Object} driver
 * @param {Array} selectedShapes - shape proxy 列表（已 load 'items/id, items/width, items/height, items/adjustments, items/tags'）
 * @param {Object} source - { cm, sourceStrict } —— pickupFromSelection 的结果
 * @param {Object} [opts] - { syncStrict: boolean } 是否复制源的防误触开启状态
 * @returns {Promise<{ok, applied, failed, strictSynced, strictAdded, strictRemoved, error?, rejectReason?}>}
 *   - strictAdded: 加 strict tag 的目标数
 *   - strictRemoved: 删 strict tag 的目标数
 *   - strictSynced: 上面两个的总和（= 实际修改 strict 状态的目标数）
 */
async function applyPickedToSelection(driver, selectedShapes, source, opts) {
  opts = opts || {};
  const syncStrict = !!opts.syncStrict;
  let applied = 0, failed = 0, strictAdded = 0, strictRemoved = 0;
  try {
    if (!source || !Number.isFinite(source.cm) || source.cm < 0) throw new Error('source.cm 不合法');
    const shapes = Array.isArray(selectedShapes) ? selectedShapes : selectedShapes && selectedShapes.items;
    if (!Array.isArray(shapes)) throw new Error('selectedShapes 格式不合法');
    const targets = shapes.filter((sh) => driver.isRoundRect(sh));
    const tagsById = await driver.loadTagsBulk(targets);
    if (targets.some((sh) => lockStateFromTags(tagsById[driver.shapeId(sh)]).isStrict)) {
      return { ok: false, applied, failed, strictAdded, strictRemoved, strictSynced: 0,
        rejectReason: 'strict', error: '选区里有形状启用了防误触，样式刷不生效' };
    }
    await withWritableShapes(driver, targets.map((sh) => driver.shapeId(sh)), async (freshTargets) => {
      if (opts.isCurrent && !opts.isCurrent()) throw new Error('stale-selection');
      for (const sh of freshTargets) {
        const state = lockStateFromTags(tagsById[driver.shapeId(sh)]);
        const r = await writeRadius(driver, sh, source.cm, { knownLockState: state });
        if (!r.ok) {
          if (r.reason !== 'not-roundRect' && r.reason !== 'no-size') failed++;
          continue;
        }
        applied++;
        if (syncStrict && source.sourceStrict) {
          // The target may be smaller than the source. Fix the clamped, actual
          // radius, including zero, before enabling protection.
          const fixed = await writeLockState(driver, sh, { lockedCm: r.newCm, isStrict: true });
          if (!fixed.ok) throw new Error(fixed.error);
          strictAdded++;
        }
      }
      await driver.sync();
    }, opts);
    return { ok: true, applied, failed, strictAdded, strictRemoved, strictSynced: strictAdded + strictRemoved };
  } catch (e) {
    const error = e && e.message ? e.message : String(e);
    console.log('[applyPickedToSelection] EXCEPTION:', error);
    return { ok: false, applied, failed, strictAdded, strictRemoved, strictSynced: strictAdded + strictRemoved, error };
  }
}
//     全部由 driver 集成测试覆盖。）

// ---------------- 导出（Node.js + browser 都支持） ----------------

if (typeof module !== 'undefined' && module.exports) {
  // Node.js
  module.exports = {
    PT_PER_CM,
    ADJ_SCALE,
    LOCK_TAG_KEY,
    LAYOUT_PARENT_TAG_KEY,
    LOCK_STRICT_TAG_KEY,
    LAYOUT_CHILD_TAG_KEY,
    computeLayout,
    computeAutoPadding,
    computeGridCoupledRowsCols,
    computeGridFactors,
    snapToNearestGridFactor,
    valueToCm,
    cmToValue,
    computeLinkedSubR,
    cmToAdj,
    clampRadius,
    computeFinalRadius,
    shouldRejectWriteRadius,
    shouldRejectOnApply,
    shouldRejectLayoutApply,
    syncFixedValueIfLocked,
    writeRadius,
    lockStateFromTags,
    decideLockMonitorUpdate,
    reapplySelectionLocks,
    applyRadiusToSelection,
    withWritableShapes,
    readLockState,
    writeLockState,
    reapplyLock,
    applyLayout,
    syncLayoutChildrenR,
    detectLayoutParentChanges,
    parseLayoutParentTagValue,
    detectStaleChildrenInLayout,
    loadLayoutTags,
    saveLayoutTags,
    pickupFromSelection,
    applyPickedToSelection,
    pushHistory,
    detectLayoutParentSizeChanges,
  };
}
if (typeof window !== 'undefined') {
  // Browser / task pane
  window.RadiusCore = {
    PT_PER_CM,
    ADJ_SCALE,
    LOCK_TAG_KEY,
    LAYOUT_PARENT_TAG_KEY,
    LOCK_STRICT_TAG_KEY,
    LAYOUT_CHILD_TAG_KEY,
    computeLayout,
    computeAutoPadding,
    computeGridCoupledRowsCols,
    computeGridFactors,
    snapToNearestGridFactor,
    valueToCm,
    cmToValue,
    computeLinkedSubR,
    cmToAdj,
    clampRadius,
    computeFinalRadius,
    pushHistory,
    shouldRejectWriteRadius,
    shouldRejectOnApply,
    shouldRejectLayoutApply,
    syncFixedValueIfLocked,
    writeRadius,
    lockStateFromTags,
    decideLockMonitorUpdate,
    reapplySelectionLocks,
    applyRadiusToSelection,
    withWritableShapes,
    readLockState,
    writeLockState,
    reapplyLock,
    applyLayout,
    syncLayoutChildrenR,
    detectLayoutParentChanges,
    parseLayoutParentTagValue,
    detectStaleChildrenInLayout,
    loadLayoutTags,
    saveLayoutTags,
    pickupFromSelection,
    applyPickedToSelection,
    detectLayoutParentSizeChanges,
  };
}
