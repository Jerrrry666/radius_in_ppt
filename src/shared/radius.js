/*
 * radius.js
 * 圆角矩形 R 角绝对值（厘米）/ 相对值（adjustments[0]）换算 + 读写
 *
 * 关键事实（PowerPoint 圆角矩形）：
 *   - shape.adjustments[0] 是 R 角相对值，范围 [0, 0.5]，是 R 角 ÷ 形状短边
 *   - 形状 width / height 单位是 EMU
 *   - 1 厘米 = 360000 EMU
 *   - 所以 绝对值(cm) = adjustments[0] * min(width, height) / 360000
 *
 * "锁定"语义：
 *   - R 角绝对值（厘米）固定；用户改变形状大小时，重新计算 adjustments[0] 比例
 *   - 用 localStorage 存 { shapeId: { radiusCm, locked } }
 */

const CM_PER_EMU = 1 / 360000;          // EMU -> 厘米
const EMU_PER_CM = 360000;              // 厘米 -> EMU
const ADJUSTMENT_MAX = 0.5;             // adjustments[0] 上限
const STORAGE_KEY = 'radius_in_ppt_locks_v1';

// Office.ShapeType.roundRect 枚举值
const ROUND_RECT_TYPE = 5;

/**
 * 读 localStorage 里的锁定表。
 * @returns {Object<string, {radiusCm:number, locked:boolean}>}
 */
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
  } catch (_) {
    // 忽略：私密模式可能写不进去
  }
}

/**
 * 从 localStorage 拿到指定 shapeId 的锁定信息（不一定存在）。
 */
function getLockEntry(shapeId) {
  return loadLocks()[shapeId] || null;
}

function setLockEntry(shapeId, entry) {
  const locks = loadLocks();
  if (entry === null) {
    delete locks[shapeId];
  } else {
    locks[shapeId] = entry;
  }
  saveLocks(locks);
}

/**
 * 把一个 PowerPoint 形状收集成统一格式。
 * 读取 width/height/adjustments/type/id 之后调用本函数。
 */
function describeShape(shape) {
  const w = shape.width || 0;
  const h = shape.height || 0;
  const shortSide = Math.min(w, h);
  const ratio = (shape.adjustments && shape.adjustments[0]) || 0;
  const radiusCm = shortSide > 0 ? ratio * shortSide * CM_PER_EMU : 0;
  return {
    id: shape.id,
    name: shape.name,
    type: shape.type,
    isRoundedRect: shape.type === ROUND_RECT_TYPE,
    width: w,
    height: h,
    shortSide,
    ratio,
    radiusCm,
  };
}

/**
 * 读取当前选区中所有圆角矩形的信息。
 * 非圆角矩形也会返回（isRoundedRect=false），方便调用方提示用户。
 *
 * @returns {Promise<{all: Array, roundedRects: Array, any: boolean, noneRounded: boolean}>}
 */
function getSelectionInfo() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Shape,
      { asyncContext: { resolve, reject } },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          result.asyncContext.reject(result.error);
          return;
        }
        const shapes = (result.value && result.value.shapes) || [];
        const all = shapes.map(describeShape);
        const roundedRects = all.filter((s) => s.isRoundedRect);
        resolve({
          all,
          roundedRects,
          any: all.length > 0,
          noneRounded: all.length > 0 && roundedRects.length === 0,
        });
      }
    );
  });
}

/**
 * 把厘米值转成 adjustments 比例（按当前短边）。
 * 会 clamp 到 [0, 0.5]，超过短边一半的值视为"满圆角"。
 */
function cmToRatio(radiusCm, shortSide) {
  if (shortSide <= 0 || !Number.isFinite(radiusCm)) return 0;
  const ratio = (radiusCm * EMU_PER_CM) / shortSide;
  return Math.max(0, Math.min(ADJUSTMENT_MAX, ratio));
}

/**
 * 给当前选区中的所有圆角矩形设置 R 角绝对值。
 * @param {number} radiusCm  单位：厘米
 * @returns {Promise<{updated:number, skipped:number}>}
 */
function applyRadiusToSelection(radiusCm) {
  return new Promise((resolve, reject) => {
    const ctx = { resolve, reject, radiusCm, updated: 0, skipped: 0 };
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Shape,
      { asyncContext: ctx },
      (result) => {
        const { asyncContext } = result;
        if (result.status === Office.AsyncResultStatus.Failed) {
          asyncContext.reject(result.error);
          return;
        }
        const shapes = (result.value && result.value.shapes) || [];
        const queue = [];
        for (const s of shapes) {
          if (s.type !== ROUND_RECT_TYPE) {
            asyncContext.skipped++;
            continue;
          }
          // 先读取 width/height（adjustments 已在 getSelectedDataAsync 里一并返回）
          if (typeof s.width !== 'number' || typeof s.height !== 'number') {
            asyncContext.skipped++;
            continue;
          }
          const shortSide = Math.min(s.width, s.height);
          const ratio = cmToRatio(asyncContext.radiusCm, shortSide);
          s.adjustments[0] = ratio;
          queue.push(s);
        }

        if (queue.length === 0) {
          asyncContext.resolve({ updated: 0, skipped: asyncContext.skipped });
          return;
        }

        // 同步写回
        const refs = queue.map((s) => s);
        refs.forEach((s) => {
          // 标记 adjustments 改变，触发 sync 时回写
          // setSelectedDataAsync 需要完整 shape 对象数组
        });

        // 用 setSelectedDataAsync 整体回写
        const writeValue = { shapes: queue };
        Office.context.document.setSelectedDataAsync(
          writeValue,
          { coercionType: Office.CoercionType.Shape },
          (writeResult) => {
            if (writeResult.status === Office.AsyncResultStatus.Failed) {
              asyncContext.reject(writeResult.error);
              return;
            }
            asyncContext.updated = queue.length;
            asyncContext.resolve({ updated: asyncContext.updated, skipped: asyncContext.skipped });
          }
        );
      }
    );
  });
}

/**
 * 把当前选区中所有圆角矩形标记为锁定 / 解锁。
 * 锁定：把当前 R 角绝对值存进 localStorage。
 * 解锁：从 localStorage 移除。
 *
 * @param {boolean} locked
 * @returns {Promise<{updated:number, skipped:number, locked:boolean}>}
 */
function setLockOnSelection(locked) {
  return getSelectionInfo().then((info) => {
    let updated = 0;
    let skipped = 0;
    for (const s of info.roundedRects) {
      if (locked) {
        setLockEntry(s.id, { radiusCm: s.radiusCm, locked: true });
      } else {
        setLockEntry(s.id, null);
      }
      updated++;
    }
    skipped = info.all.length - updated;
    return { updated, skipped, locked };
  });
}

/**
 * 重新应用锁定（用于"立即应用"或"重新应用锁定"按钮）。
 * 遍历所有锁定的 shapeId（不一定在选区里），把 R 角绝对值重新写入。
 * 由于 Office.js 不暴露"按 id 选中"，我们通过当前选区匹配 + 全文档遍历实现。
 *
 * 这里采用"重新应用选区里那些有锁定标记的形状"的策略。
 * 如果你想"用户改完大小就自动锁定"，可以再绑定 SelectionChanged 时自动调用本函数。
 */
function reapplyLocksToSelection() {
  return getSelectionInfo().then((info) => {
    const locks = loadLocks();
    let reapplied = 0;
    const queue = [];
    for (const s of info.roundedRects) {
      const entry = locks[s.id];
      if (entry && entry.locked && Number.isFinite(entry.radiusCm)) {
        const ratio = cmToRatio(entry.radiusCm, s.shortSide);
        s.adjustments[0] = ratio;
        queue.push(s);
        reapplied++;
      }
    }
    if (queue.length === 0) {
      return Promise.resolve({ reapplied: 0 });
    }
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        { shapes: queue },
        { coercionType: Office.CoercionType.Shape },
        (res) => {
          if (res.status === Office.AsyncResultStatus.Failed) reject(res.error);
          else resolve({ reapplied });
        }
      );
    });
  });
}

/**
 * 监听选区变化。每次选区变化时，把"当前选区里所有有锁定标记的形状"重新应用 R 角。
 * 配合"锁定 R 角绝对值"语义：用户改大小后切换选区再切回来，R 角自动恢复。
 */
function installSelectionChangedAutoReapply() {
  if (!Office.context.document.addHandlerAsync) return;
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => {
      // 异步重应用，错误吞掉避免打扰用户
      reapplyLocksToSelection().catch(() => {});
    }
  );
}

// 暴露到全局（dialog 页面使用）
window.RadiusCore = {
  ROUND_RECT_TYPE,
  ADJUSTMENT_MAX,
  CM_PER_EMU,
  EMU_PER_CM,
  loadLocks,
  saveLocks,
  getLockEntry,
  setLockEntry,
  getSelectionInfo,
  applyRadiusToSelection,
  setLockOnSelection,
  reapplyLocksToSelection,
  installSelectionChangedAutoReapply,
  cmToRatio,
};
