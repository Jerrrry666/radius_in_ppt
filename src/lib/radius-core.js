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

// ---------------- 布局 math ----------------

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
    return {
      subW: 0, subH: 0, positions: [], feasible: false,
      reason: '边距/间距太大，挤不下 ' + cols + ' 列 × ' + rows + ' 行',
    };
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
    // 1. 读 lock + strict（通过 driver，async）
    const lockVal = await driver.readTag(shape, LOCK_TAG_KEY);
    let isLocked = false;
    let lockedCm = 0;
    if (lockVal) {
      const cm = parseFloat(lockVal);
      if (Number.isFinite(cm) && cm > 0) {
        isLocked = true;
        lockedCm = cm;
      }
    }
    const strictVal = await driver.readTag(shape, LOCK_STRICT_TAG_KEY);
    const isStrict = strictVal === '1';

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
      try { driver.addTag(shape, LAYOUT_CHILD_TAG_KEY, layoutParentId); } catch (_) {}
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
    if (Number.isFinite(cm) && cm > 0) lockedCm = cm;
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
        try { driver.deleteTag(shape, LOCK_TAG_KEY); } catch (_) {}
      } else {
        try { driver.addTag(shape, LOCK_TAG_KEY, String(state.lockedCm)); } catch (_) {}
      }
    }
    if (state.isStrict === true) {
      try { driver.addTag(shape, LOCK_STRICT_TAG_KEY, '1'); } catch (_) {}
    } else if (state.isStrict === false) {
      try { driver.deleteTag(shape, LOCK_STRICT_TAG_KEY); } catch (_) {}
    }
    return { ok: true };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
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
 * @param {Object} [opts] - { writeParentTag, syncR }
 * @returns {Promise<{ok, applied, failed, warn, strictOverridden, lockedCount, error?}>}
 */
async function applyLayout(driver, parentId, params, childIds, opts) {
  opts = opts || {};
  const writeParentTag = opts.writeParentTag !== false;
  const syncR = opts.syncR !== false;
  const linkRMode = params.linkRMode || 'subtract';
  const expectedCount = params.rows * params.cols;

  let applied = 0;
  let failed = 0;
  let strictOverridden = 0;
  let lockedCount = 0;
  let warn = '';
  let lockedChildCm = [];

  try {
    // 1. 当前 slide + 集合层 load（Mac LTSC 必加，per-shape load .count 永远 = 0）
    const slide = driver.activeSlide();
    driver.load(slide, 'shapes/items/id, shapes/items/left, shapes/items/top, shapes/items/width, shapes/items/height, shapes/items/adjustments, shapes/items/tags');
    await driver.sync();

    // 2. 建 id → shape 映射
    const idToShape = new Map();
    const slideShapes = driver.slideShapes(slide);
    for (const sh of slideShapes.items) {
      const id = driver.shapeId(sh);
      if (id != null) idToShape.set(id, sh);
    }

    // 3. 找父
    const parentSh = idToShape.get(parentId);
    if (!parentSh) {
      warn = '父形状在当前 slide 找不到（可能选了别的页）';
      console.log('[applyLayout/driver] WARN parent not found in current slide');
      return { ok: false, applied, failed, warn };
    }

    // 4. 父 R 角（v1.0 per-shape get(0) + sync + 读）
    let parentRcm = 0;
    try {
      if (driver.isRoundRect(parentSh)) {
        const adjResult = parentSh.adjustments.get(0);
        await driver.sync();
        try { parentRcm = adjResult.value * Math.min(driver.size(parentSh).width, driver.size(parentSh).height) / PT_PER_CM; } catch (_) {}
      }
    } catch (_) { /* 父不是 roundRect 时算 0 */ }

    // 5. 父 box + 算 layout
    const parentBox = driver.box(parentSh);
    console.log('[applyLayout/driver] parent box:', JSON.stringify(parentBox), 'Rcm=', parentRcm);

    const layout = computeLayout(parentBox, params.rows, params.cols, params.padding, params.gutter);
    if (!layout.feasible) {
      warn = layout.reason;
      return { ok: false, applied, failed, warn };
    }

    // 6. 收集存在的子（过滤掉 stale / 不在当前 slide 的）
    const validChildIds = [];
    const childShapes = [];
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

    // 7. 写每个子的位置 + 尺寸 + R 角 + child tag
    for (let k = 0; k < childShapes.length; k++) {
      const csh = childShapes[k];
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
      // 写 R 角（走 writeRadius，第二道防线实时查 strict + 同步 lock fixed value）
      if (syncR && linkRMode !== 'off') {
        const subRcm = linkRMode === 'same' ? parentRcm : Math.max(0, parentRcm - params.padding);
        console.log(`[applyLayout/driver] R link #${k}: parentRcm=${parentRcm}, mode=${linkRMode}, padding=${params.padding}, target subRcm=${subRcm}`);
        const r = await writeRadius(driver, csh, subRcm, { layoutParentId: parentId });
        if (r.ok) {
          console.log(`[applyLayout/driver] R link #${k}: written subRcm=${r.newCm}, wasLocked=${r.wasLocked}`);
          if (r.wasLocked) {
            lockedCount++;
            lockedChildCm.push({ id: driver.shapeId(csh), newCm: r.newCm });
          }
          if (r.wasStrict) strictOverridden++;
        } else {
          console.log(`[applyLayout/driver] R link #${k}: skipped, reason=${r.reason}${r.error ? ' error=' + r.error : ''}`);
        }
      }
      // 写 child tag
      try { driver.addTag(csh, LAYOUT_CHILD_TAG_KEY, parentId); } catch (_) {}
      applied++;
    }
    console.log('[applyLayout/driver] applied=', applied, 'failed=', failed);

    // 8. 写父 tag（**用 validChildIds 过滤后的版本**，stale childIds 自动清理）
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
      }
    }
    await driver.sync();
    console.log('[applyLayout/driver] sync done, lockedChildCm count=', lockedChildCm.length);
    return { ok: true, applied, failed, warn, strictOverridden, lockedCount };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log('[applyLayout/driver] OUTER ERROR:', msg);
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
  if (linkRMode === 'off' || !parentRcm) return { ok: true, applied: 0, failed: 0 };
  let applied = 0;
  let failed = 0;
  try {
    const slide = driver.activeSlide();
    driver.load(slide, 'shapes/items/id, shapes/items/width, shapes/items/height, shapes/items/adjustments, shapes/items/tags');
    await driver.sync();
    const idToShape = new Map();
    for (const sh of driver.slideShapes(slide).items) {
      const id = driver.shapeId(sh);
      if (id != null) idToShape.set(id, sh);
    }
    for (const childId of childIds) {
      const csh = idToShape.get(childId);
      if (!csh) {
        console.log(`[syncLayoutChildrenR/driver] skip missing child id=${childId} (stale)`);
        continue;  // skip stale（不在当前 slide / 已删）
      }
      const subRcm = linkRMode === 'same' ? parentRcm : Math.max(0, parentRcm - paddingCm);
      console.log(`[syncLayoutChildrenR/driver] R link child=${childId} parentRcm=${parentRcm} mode=${linkRMode} padding=${paddingCm} target subRcm=${subRcm}`);
      const r = await writeRadius(driver, csh, subRcm, {});
      if (r.ok) {
        applied++;
      } else if (r.reason === 'strict') {
        console.log('[syncLayoutChildrenR/driver] skip strict child', childId);
      } else if (r.reason !== 'not-roundRect' && r.reason !== 'no-size') {
        failed++;
      }
    }
    await driver.sync();
    console.log(`[syncLayoutChildrenR/driver] done: applied=${applied} failed=${failed} childIds=${JSON.stringify(childIds)}`);
    return { ok: true, applied, failed };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log('[syncLayoutChildrenR/driver] OUTER ERROR:', msg);
    return { ok: false, applied, failed, error: msg };
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
    readLockState,
    writeLockState,
    reapplyLock,
    applyLayout,
    syncLayoutChildrenR,
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
    readLockState,
    writeLockState,
    reapplyLock,
    applyLayout,
    syncLayoutChildrenR,
  };
}
