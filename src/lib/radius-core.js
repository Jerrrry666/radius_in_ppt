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
 * v1.2.7：自适应 padding 纯函数
 *
 * 问题：same 模式（R_sub = R_父）下，当 R_父 过大时子 R 角会被 clamp 到子短边一半，
 *       几何上不等宽了（破等宽）。
 * 修法：当 R_父 > min(W, H) / 2 - d_init 时，**自动减小 padding** 让子短边 >= 2*R_父。
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

// ---------------- v1.2.8: iOS 风格连续曲率 (squircle) 数学参考 ----------------

/**
 * v1.2.8：iOS 风格连续曲率角的几何参数（参考 Figma "Desperately Seeking Squircles"
 *  + MartinRGB 反编译的 rounded-corners.js）
 *
 * 背景：
 *   - 普通圆角矩形（PowerPoint 原生）是单段 90° 圆弧 + 直线段，曲率在角与边的衔接处
 *     突然从 1/R 跳到 0（G1 only）。视觉上"硬"一点。
 *   - iOS / Figma 的连续曲率（squircle）每个角由「1 段 1/4 圆弧 + 2 段贝塞尔曲线」组成，
 *     曲率从 1/R 平滑过渡到 0（G2 continuity）。视觉上"软"一点。
 *   - cornerSmoothing 越大，角越像 squircle；0 = 完全普通圆角，1 = 接近纯 squircle。
 *
 * 重要限制（v1.2.8 范围内）：
 *   - PowerPoint 圆角矩形只支持单段圆弧（adjustment value 0-1），**无法**写入多段贝塞尔
 *   - 本函数**只算参考参数**，不写 PPT；UI 用来展示"如果用 Figma 算法画的角，理论长啥样"
 *   - 真正的 squircle 需要嵌入 SVG 路径 / 自定义形状（v1.2.8 不做）
 *
 * 公式（Figma 文档 figure 12.2 / MartinRGB 实现）：
 *   p = (1 + cornerSmoothing) * cornerRadius           // 角的总占地
 *   arcMeasure = 90° * (1 - cornerSmoothing)           // 圆弧度数
 *   arcSectionLength = sin(arcMeasure/2) * R * sqrt(2)  // 圆弧段长
 *   alpha = (90° - arcMeasure) / 2
 *   p3ToP4 = R * tan(alpha / 2)                         // 圆弧两端控制点距离
 *   beta = 45° * cornerSmoothing
 *   c = p3ToP4 * cos(beta)
 *   d = c * tan(beta)
 *   b = (p - arcSectionLength - c - d) / 3
 *   a = 2b
 *
 * @param {number} rCm - 当前 R 角（cm）
 * @param {number} smoothing - 0~1，0 = 普通圆角，1 = 最大平滑（接近纯 squircle）
 *   - iOS 7 app 图标 ≈ 0.6
 *   - iOS 13+ app 图标 ≈ 0.6~0.7
 *   - SwiftUI .continuous ≈ 0.6（Apple default）
 * @returns {Object} 几何参数（单位都是 cm 或度）
 *   - pCm: 角的总占地（cm）
 *   - arcMeasureDeg: 圆弧度数（0~90）
 *   - arcLengthCm: 圆弧段长（cm）
 *   - betaDeg: 45° × smoothing
 *   - cCm, dCm, aCm, bCm: 贝塞尔控制点距离（cm）
 *   - figmaEquivalent: string "iOS 7" / "高 squircle" / "低平滑" 给 UI 展示用
 */
function computeSquircleHint(rCm, smoothing) {
  const R = Number.isFinite(rCm) ? Math.max(0, rCm) : 0;
  const s = Number.isFinite(smoothing) ? Math.max(0, Math.min(1, smoothing)) : 0;
  const pCm = (1 + s) * R;
  const arcMeasureDeg = 90 * (1 - s);
  const arcLengthCm = Math.sin(toRadians(arcMeasureDeg / 2)) * R * Math.sqrt(2);
  const alphaDeg = (90 - arcMeasureDeg) / 2;
  const p3ToP4Cm = R * Math.tan(toRadians(alphaDeg / 2));
  const betaDeg = 45 * s;
  const cCm = p3ToP4Cm * Math.cos(toRadians(betaDeg));
  const dCm = cCm * Math.tan(toRadians(betaDeg));
  const bCm = (pCm - arcLengthCm - cCm - dCm) / 3;
  const aCm = 2 * bCm;
  // UI 标签：iOS 7 默认 0.6 → 标 "iOS 7+ 风格"
  let figmaEquivalent = '普通圆角 (G1)';
  if (s >= 0.55 && s <= 0.7) figmaEquivalent = 'iOS 7+ 风格 (G2)';
  else if (s > 0.7) figmaEquivalent = '高 squircle (G2+)';
  else if (s > 0) figmaEquivalent = '低平滑 (G1+)';
  return {
    pCm,
    arcMeasureDeg,
    arcLengthCm,
    betaDeg,
    cCm,
    dCm,
    aCm,
    bCm,
    figmaEquivalent,
  };
}

/** iOS 7 squircle 默认 smoothing（Apple 在 iOS 7 引入的连续曲率参考值） */
const IOS7_DEFAULT_SMOOTHING = 0.6;

function toRadians(degrees) {
  return (degrees * Math.PI) / 180;
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
    //    修法：caller 用 driver.readTagsBulk 一次拿全部 → 传 knownLockState → 跳过 per-shape readTag
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
        if (Number.isFinite(cm) && cm > 0) {
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
  // v1.2.6：默认 linkRMode 从 'subtract' 改成 'same'
  // 原因：subtract 公式 R_sub = R_父 - d 几何上**不等距**（45° 方向距离 = d - 0.414d ≈ 0.586d），
  //       user 报"子 R 角看着不美观，角部比边窄"。same 公式 R_sub = R_父 是真正的等距。
  //       老 layout（linkRMode 已存 'subtract'）不受影响（直接读 tag 拿到 'subtract'）
  const linkRMode = params.linkRMode || 'same';
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

    // v1.2.7：autoPadding — 父 R 角过大时自动减小 padding，保证 same 模式子 R 角不 clamp
    const parentWcm = parentBox.width / PT_PER_CM;
    const parentHcm = parentBox.height / PT_PER_CM;
    const ap = computeAutoPadding(parentWcm, parentHcm, parentRcm, params.padding);
    const effectivePadding = ap.effectivePaddingCm;
    if (ap.clamped) {
      console.log(`[applyLayout/driver] autoPadding: R=${parentRcm.toFixed(3)}cm > d_max=${ap.dMaxCm.toFixed(3)}cm, d ${params.padding}→${effectivePadding.toFixed(3)}cm`);
    }

    const layout = computeLayout(parentBox, params.rows, params.cols, effectivePadding, params.gutter);
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
    // v1.3.6 修 #6：先 readTagsBulk 拿全部 tag（避开 per-call readTag + sync 在 for 循环内累积）
    const tagsById = driver.readTagsBulk(slideShapes.items);
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
        // 从 readTagsBulk 拿的 tag 构造 knownLockState，传给 writeRadius（跳过 per-call readTag + sync）
        const cTags = tagsById[driver.shapeId(csh)] || {};
        const lockRaw = cTags[LOCK_TAG_KEY];
        let isLocked = false;
        let lockedCm = 0;
        if (lockRaw != null) {
          const cm = parseFloat(lockRaw);
          if (Number.isFinite(cm) && cm > 0) {
            isLocked = true;
            lockedCm = cm;
          }
        }
        const isStrict = cTags[LOCK_STRICT_TAG_KEY] === '1';
        const r = await writeRadius(driver, csh, subRcm, {
          layoutParentId: parentId,
          knownLockState: { isLocked, lockedCm, isStrict },
        });
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
    // v1.3.6 修 #6 子 bug：4 个子只写 2 个（用户实测：调整父 R 角后上面 2 个子变了下面 2 个没变）
    // 根因：writeRadius 内部 readTag 调 ctx.sync()，4 次 readTag + 4 次 setAdjFraction 在同一个 PowerPoint.run
    //       Mac LTSC 上 per-shape sync 累积（v1.2.6 同样坑），后几个 shape 的 setAdjFraction 失败/丢失
    // 修法：sibling applyLayout 模式 —— 一次 load + sync 拿全部 tag（用 readTagsBulk 避开 per-call sync），
    //       写所有 setAdjFraction，final sync 一次
    const slide = driver.activeSlide();
    driver.load(slide, 'shapes/items/id, shapes/items/width, shapes/items/height, shapes/items/adjustments, shapes/items/tags');
    await driver.sync();
    const slideShapes = driver.slideShapes(slide).items;
    const idToShape = new Map();
    for (const sh of slideShapes) {
      const id = driver.shapeId(sh);
      if (id != null) idToShape.set(id, sh);
    }
    // 一次拿全部 tag（不调 ctx.sync()，避免 per-call sync 累积）
    const tagsById = driver.readTagsBulk(slideShapes);
    for (const childId of childIds) {
      const csh = idToShape.get(childId);
      if (!csh) {
        console.log(`[syncLayoutChildrenR/driver] skip missing child id=${childId} (stale)`);
        continue;  // skip stale（不在当前 slide / 已删）
      }
      const subRcm = linkRMode === 'same' ? parentRcm : Math.max(0, parentRcm - paddingCm);
      console.log(`[syncLayoutChildrenR/driver] R link child=${childId} parentRcm=${parentRcm} mode=${linkRMode} padding=${paddingCm} target subRcm=${subRcm}`);
      // 从 readTagsBulk 拿的 tag 构造 knownLockState，传给 writeRadius（跳过 per-call readTag + sync）
      const childTags = tagsById[childId] || {};
      const lockRaw = childTags[LOCK_TAG_KEY];
      let isLocked = false;
      let lockedCm = 0;
      if (lockRaw != null) {
        const cm = parseFloat(lockRaw);
        if (Number.isFinite(cm) && cm > 0) {
          isLocked = true;
          lockedCm = cm;
        }
      }
      const isStrict = childTags[LOCK_STRICT_TAG_KEY] === '1';
      const r = await writeRadius(driver, csh, subRcm, {
        knownLockState: { isLocked, lockedCm, isStrict },
      });
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
    if (!obj || !Number.isFinite(obj.rows) || !Number.isFinite(obj.cols) || !Array.isArray(obj.childIds)) {
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

    // 读每个 shape 的 layout tag
    for (const sh of shapesList) {
      const sid = sh.id;
      if (sid == null) continue;
      // 父 tag
      const parentVal = await driver.readTag(sh, LAYOUT_PARENT_TAG_KEY);
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
      const childVal = await driver.readTag(sh, LAYOUT_CHILD_TAG_KEY);
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
async function saveLayoutTags(driver, slide, parentId, params, childIds) {
  try {
    // 集合层 load slide shapes（id only）
    driver.load(slide, 'shapes/items/id');
    await driver.sync();

    // 建 id → shape 映射
    const idToShape = new Map();
    const slideShapesArr = driver.slideShapes(slide).items;
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
    driver.addTag(parentSh, LAYOUT_PARENT_TAG_KEY, payload);

    // 写子 tag（只写 valid 的）
    for (const cid of validChildIds) {
      const csh = idToShape.get(cid);
      try { driver.addTag(csh, LAYOUT_CHILD_TAG_KEY, parentId); } catch (_) {}
    }
    await driver.sync();

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
 *   2. 找第一个 adjustments.count > 0 的形状
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
        const adjResult = sh.adjustments.get(0);
        await driver.sync();
        let v = null;
        try { v = adjResult.value; } catch (_) { continue; }
        if (!Number.isFinite(v)) continue;
        const size = driver.size(sh);
        const minSideCm = Math.min(size.width, size.height) / PT_PER_CM;
        const cm = v * minSideCm;
        // 读 strict
        let sourceStrict = false;
        try {
          const strictVal = await driver.readTag(sh, LOCK_STRICT_TAG_KEY);
          if (strictVal === '1') sourceStrict = true;
        } catch (_) {}
        return {
          id: driver.shapeId(sh),
          name: sh.name || '(未命名)',
          cm,
          sourceStrict,
        };
      } catch (_) {
        // 单 shape 失败不 throw 整个（user 报 #1 期间可能 race）
        continue;
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
 * 行为：
 *   1. 拦截：选区里有 strict 形状 → 全部拒绝（用 driver.readTag 实时查）
 *   2. 写 R 角：对每个 roundRect 调 writeRadius（自动处理 clamp + lock 同步）
 *   3. （可选）刷 strict 状态：sourceStrict = true → 给所有目标加 strict tag
 *
 * @param {Object} driver
 * @param {Array} selectedShapes - shape proxy 列表（已 load 'items/id, items/width, items/height, items/adjustments, items/tags'）
 * @param {Object} source - { cm, sourceStrict } —— pickupFromSelection 的结果
 * @param {Object} [opts] - { syncStrict: boolean }  是否刷入 strict 状态
 * @returns {Promise<{ok, applied, failed, strictSynced, error?, rejectReason?}>}
 */
async function applyPickedToSelection(driver, selectedShapes, source, opts) {
  opts = opts || {};
  const syncStrict = !!opts.syncStrict;
  let applied = 0;
  let failed = 0;
  let strictSynced = 0;
  try {
    if (!source || !Number.isFinite(source.cm)) {
      return { ok: false, applied, failed, strictSynced, error: 'source.cm 不合法' };
    }
    const shapesList = [];
    if (selectedShapes && typeof selectedShapes.items !== 'undefined') {
      for (const sh of selectedShapes.items) shapesList.push(sh);
    } else if (Array.isArray(selectedShapes)) {
      for (const sh of selectedShapes) shapesList.push(sh);
    } else {
      return { ok: false, applied, failed, strictSynced, error: 'selectedShapes 格式不合法' };
    }

    // 步骤 0：拦截（实时读 strict 标签，不依赖内存）
    for (const sh of shapesList) {
      try {
        if (!driver.isRoundRect(sh)) continue;
        const strictVal = await driver.readTag(sh, LOCK_STRICT_TAG_KEY);
        if (strictVal === '1') {
          return {
            ok: false,
            applied,
            failed,
            strictSynced,
            rejectReason: 'strict',
            error: '选区里有形状启用了防误触，样式刷不生效',
          };
        }
      } catch (_) { /* 读 strict 失败不拦截（defensive） */ }
    }

    // 步骤 1：刷 R 角
    for (const sh of shapesList) {
      try {
        if (!driver.isRoundRect(sh)) continue;
        const r = await writeRadius(driver, sh, source.cm, {});
        if (r.ok) {
          applied++;
        } else if (r.reason === 'strict') {
          // 步骤 0 已经查过，但 writeRadius 是第二道防线
          failed++;
        } else if (r.reason === 'not-roundRect' || r.reason === 'no-size') {
          // 不是 roundRect / 0 尺寸 → 跳过（不计入 failed）
        } else {
          failed++;
        }
      } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        console.log('[applyPickedToSelection] shape fail:', msg);
        failed++;
      }
    }

    // 步骤 2：刷 strict 状态（可选）
    if (syncStrict && source.sourceStrict) {
      for (const sh of shapesList) {
        try {
          if (!driver.isRoundRect(sh)) continue;
          driver.addTag(sh, LOCK_STRICT_TAG_KEY, '1');
          strictSynced++;
        } catch (_) {}
      }
    }
    await driver.sync();
    return { ok: true, applied, failed, strictSynced };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    return { ok: false, applied, failed, strictSynced, error: msg };
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
    detectLayoutParentChanges,
    parseLayoutParentTagValue,
    detectStaleChildrenInLayout,
    loadLayoutTags,
    saveLayoutTags,
    pickupFromSelection,
    applyPickedToSelection,
    pushHistory,
    computeSquircleHint,
    IOS7_DEFAULT_SMOOTHING,
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
    computeSquircleHint,
    IOS7_DEFAULT_SMOOTHING,
  };
}
