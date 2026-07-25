/*
 * dialog.js — R 角调整 v1.2（task pane，纯 Office.js）
 *
 * UI 形式：PowerPoint 侧边栏（task pane），ribbon 上点按钮展开。
 *
 * v1.0 功能：
 *   1. 打开 task pane → getSelectedShapes() → 显示选中的圆角矩形
 *   2. 用户输入 R 角（cm 或 %）→ 「应用 R 角」→ adjustments.set(0, newVal)
 *   3. 「使用数值固定 R 角」→ 写 shape.tags（OOXML <p:tagLst>），跟 .pptx 文件走
 *   4. history 槽位 = 本次 session 内用户主动应用过的 R 角（纯内存）
 *
 * v1.1 新增：
 *   5. 预设库（5 槽位，纯内存，session 内）— 把当前输入框的值存为预设，点预设即应用
 *   6. R 角样式刷（idle / sourcing / brushing 状态机）— 吸 1 个形状的 R 角，连刷其他形状
 *
 * v1.2 新增：
 *   7. 布局模式（rows × cols 网格 + 边距/间距滑块 + R 角联动）
 *      - 选中 1 大 + N×M 小圆角矩形 → 「建布局」 → 子按公式分布
 *      - 滑块拖动：行/列/边距/间距 → 实时重算 + 应用
 *      - R 角联动：子 R = max(0, 父 R − 边距)
 *      - 父子状态用 shape.tags 双向挂载（layoutParent_v1 / layoutChild_v1）
 *
 * Mac LTSC (Office 2021, build 16.111) 实测要点：
 *   - `customProperties` / `customXmlParts` 在 task pane 都不可用 → 锁用 shape.tags
 *   - `adjustments.get(0)` 返回 ClientResult 代理，直接 .value 读（不要 .load）
 *   - `shape.adjustments.get(0).value` 是 0~1 比例（不是 OOXML 0~50000）
 *   - Office.js PowerPoint 没有 shape change 事件 → lock 自动重应用靠 setInterval 轮询
 *
 * 监听 PPT 选区变化：DocumentSelectionChanged →
 *   - idle 状态 → refreshSelection（刷新 selectedShapes 内存 + 渲染）
 *   - sourcing 状态 → 从选区吸取 R 角（pickupFromSelection）
 *   - brushing 状态 → 把 R 角应用到选区里所有 roundRect（applyPipetteToSelection）
 * 多页 PPT：getSelectedShapes() 只返回当前页选中的形状，切页后选区变化 → 自动刷新
 */

(function () {
  const $ = (id) => document.getElementById(id);
  const PT_PER_CM = 28.3464567;        // 1 cm = 28.3464567 pt
  // Mac LTSC: adjustments.get(0).value 是 0~1 比例（占短边），不是 OOXML 0~50000
  const ADJ_SCALE = 1;
  const MAX_HISTORY = 5;
  const LOCK_TAG_KEY = 'radiusLock_v1';
  const LOCK_STRICT_TAG_KEY = 'radiusLockStrict_v1'; // 防误触开关：value "1" = 开启
  // v1.2: 布局模式 tag key（双向挂：父挂 layoutParent_v1，子挂 layoutChild_v1）
  const LAYOUT_PARENT_TAG_KEY = 'layoutParent_v1';
  const LAYOUT_CHILD_TAG_KEY = 'layoutChild_v1';
  // 布局 R 角联动 hook 时的小阈值（避免无意义写）
  const LAYOUT_LAYOUT_RT_DEBOUNCE_MS = 50;

  // 本次 session 内用户主动应用过的 R 角值（纯内存）
  let userHistory = [];

  // 当前选中的形状（refreshSelection 填充）
  // v1.2 扩展字段：layoutRole ('parent' | 'child' | null), layoutParentId
  let selectedShapes = [];

  // v1.2: 当前激活的 layout（当且仅当选中形状里有 layout 父时存在）
  // { parentId, parentName, childIds: [string], params: {rows,cols,padding,gutter,linkR} }
  let currentLayout = null;

  // 当前输入单位：'cm' | '%'
  let currentUnit = 'cm';

  // lock monitor 状态：选区里有 locked 形状时启动，10ms 轮询
  // v1.1 行为：通过 width / adj 变化识别两种拖动
  //  - 拖尺寸手柄（width/height 变） → 立刻反算回固定值
  //  - 拖 R 角黄色滑块（adj 变 + width 不变）→ 视作主动改值：
  //      · 仅「使用数值固定 R 角」（非 strict）→ 更新固定值到当前 adj
  //      · 「防误触」（strict）→ 反算回去
  const LOCK_POLL_MS = 10;
  const IDLE_POLL_MS = 50;          // 未锁定时只读不写，频率慢一点
  const LOCK_STABLE_THRESHOLD = 4;
  const ADJ_EPSILON = 0.0001;       // adj 比较的容差
  const SIZE_EPSILON = 0.001;       // pt（≈ 0.00035 cm）的容差，用于 width / height 变化检测
  let lockMonitor = {
    timer: null,
    lastWidth: {},     // shapeId -> 上次读到的 width（pt）
    lastHeight: {},    // shapeId -> 上次读到的 height（pt）
    lastAdj: {},       // shapeId -> 上次读到的 adj
    stableCount: {},   // shapeId -> adj 连续稳定次数
    lastCm: {},        // shapeId -> 上次读到的 currentCm（cm）—— 仅 layout 父用
    lastSizeCm: {},    // shapeId -> 上次读到的 { widthCm, heightCm } —— 仅 layout 父用（v1.2.9 size 联动）
    parentRDirty: false,  // 是否有 layout 父 R 角或 size 变化需要联动（v1.2.9 合并）
    parentRSyncTimer: null,  // 节流 timer（避免 10ms tick 频繁触发新 run）
  };
  const PARENT_R_SYNC_DEBOUNCE_MS = 200;  // 父 R 角变化 → 同步子的节流窗口

  // ---------------- 单位换算 ----------------

  function getRefShapeMinSideCm() {
    // % 模式的 100% 参考：用第一个 roundRect 的 minSide
    for (const s of selectedShapes) {
      if (s.minSideCm > 0) return s.minSideCm;
    }
    return 0;
  }

  function valueToCm(val, unit) {
    if (unit === '%') {
      const minSideCm = getRefShapeMinSideCm();
      return (val / 100) * minSideCm;
    }
    return val;
  }

  function cmToValue(cm, unit) {
    if (unit === '%') {
      const minSideCm = getRefShapeMinSideCm();
      if (minSideCm <= 0) return 0;
      return (cm / minSideCm) * 100;
    }
    return cm;
  }

  // ---------------- shape.tags 锁（Mac LTSC 唯一可用的持久化方案） ----------------

  // 同时返回 locks（id -> cm）和 strict（id -> true/false）
  // v1.2.2 driver + radius-core 迁移：lock/strict 走新分层
  function loadLocksViaTags() {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const driver = window.PptDriver.createDriver(ctx);
          const sel = driver.selectedShapes();
          driver.load(sel, 'items/id, items/tags');
          await driver.sync();
          const locks = {};   // id -> cm（使用数值固定 R 角）
          const strict = {};  // id -> true（防误触开关）
          for (const sh of sel.items) {
            const state = await window.RadiusCore.readLockState(driver, sh);
            const id = driver.shapeId(sh);
            if (state.lockedCm != null) locks[id] = state.lockedCm;
            if (state.isStrict) strict[id] = true;
          }
          resolve({ ok: true, locks, strict });
        } catch (e) {
          resolve({ ok: false, error: e });
        }
      });
    });
  }

  // strictMap: { id: true } 表示该 id 开启防误触；省略或 false 表示关闭
  function saveLocksViaTags(locks, strictMap) {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const driver = window.PptDriver.createDriver(ctx);
          const sel = driver.selectedShapes();
          driver.load(sel, 'items/id');
          await driver.sync();
          for (const sh of sel.items) {
            const id = driver.shapeId(sh);
            const state = {};
            if (id in locks) state.lockedCm = locks[id];  // number 写 / null 删 / undefined 跳过
            if (id in (strictMap || {})) state.isStrict = !!strictMap[id];
            const r = await window.RadiusCore.writeLockState(driver, sh, state);
            if (!r.ok) console.log('[saveLocks/driver] id=' + id + ' fail: ' + r.error);
          }
          await driver.sync();
          resolve({ ok: true });
        } catch (e) {
          resolve({ ok: false, error: e });
        }
      });
    });
  }

  // 单独更新某个 shape 的 lock tag（onApply 写完 PowerPoint 后调用）
  // cm 语义：number = 写 / null = 删 / undefined = 不动
  // isStrict 语义：true = 开 / false = 关 / null/undefined = 不动
  // v1.2.2 driver + radius-core 迁移
  function updateLockTagForShape(shapeId, cm, isStrict) {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const driver = window.PptDriver.createDriver(ctx);
          const sel = driver.selectedShapes();
          driver.load(sel, 'items/id');
          await driver.sync();
          for (const sh of sel.items) {
            if (driver.shapeId(sh) !== shapeId) continue;
            const state = {};
            if (cm !== undefined) state.lockedCm = cm;       // number 写 / null 删 / undefined 跳过
            if (isStrict === true) state.isStrict = true;
            else if (isStrict === false) state.isStrict = false;
            const r = await window.RadiusCore.writeLockState(driver, sh, state);
            if (!r.ok) console.log('[updateLockTag/driver] id=' + shapeId + ' fail: ' + r.error);
            break;
          }
          await driver.sync();
          resolve({ ok: true });
        } catch (e) {
          resolve({ ok: false, error: e });
        }
      });
    });
  }

  // ---------------- v1.2: 统一写 R 角函数 ----------------
  // v1.3.6：删 writeRadiusToShape（v1.2.2 之前的兼容版，已被 radius-core.writeRadius 取代）
  //   之前 dialog.js 的 applyPipetteToSelection 还调这个老函数（不走 driver），
  //   走的是直接 ctxShape.adjustments.set + ctxShape.tags.getItem，
  //   跟 radius-core.writeRadius（driver 版 + 走 setAdjFraction）行为重复。
  //   v1.3.6 applyPipetteToSelection 改调 radius-core.applyPickedToSelection 后，
  //   writeRadiusToShape 无人调用 → 删了。
  // 修 #1：吸取后无法刷入任何形状的根因（之一）—— dialog.js 用 writeRadiusToShape 不走 driver，
  //        改走 radius-core（driver 版）后稳定。

  // ---------------- v1.2: layout 计算 + apply pipeline ----------------

  // 纯函数：给定父 box + rows/cols/padding/gutter，算出子形状的尺寸 + 位置
  // parent: { left, top, width, height } (pt)
  // 返回：{ subW, subH, positions: [{left, top, w, h, idx}], feasible: bool, reason: string }
  //   positions 按 row-major 排：i*cols + j
  //   feasible = false 表示 padding/gutter 太大，子尺寸 ≤ 0
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

  // 把当前 params 应用到子形状：写位置 + 尺寸 + R 角（如果 linkR）
  // parentId: 父 shape id
  // params: { rows, cols, padding, gutter, linkR }
  // childIds: [string]
  // opts: { writeParentTag: bool, syncR: bool }
  //   - writeParentTag: 写 layoutParent_v1 tag（首次创建时 true；只调参数时 false）
  //   - syncR: 是否同步 R 角（链接时 true，单纯调位置时 false）
  // 返回：{ ok, applied, failed, warn }
  // v1.2.9 迁移：applyLayoutToChildren 第一道防线（strict 检查）保留在 dialog.js（依赖 selectedShapes 状态），
  // PowerPoint.run 部分（写位置/尺寸/R 角/父 tag + 过滤 stale childIds）全部走 radius-core.applyLayout
  async function applyLayoutToChildren(parentId, params, childIds, opts) {
    opts = opts || {};
    const writeParentTag = opts.writeParentTag !== false;
    const syncR = opts.syncR !== false;
    // v1.3.6：v1.2 step 调试 log（verified 后无用，删除）—— 改在 radius-core 内保留必要的诊断 log

    // 第一道防线：进 PowerPoint.run 之前，检查选区里任何子有防误触 → 拒绝整个 apply
    // （位置/尺寸也不写，避免半成品状态；防误触永远最高优先级）
    const childIdsForStrict = childIds.slice(0, params.rows * params.cols);
    const strictInSelection = selectedShapes.filter((s) =>
      s.layoutRole !== 'parent' && childIdsForStrict.indexOf(s.id) >= 0 && s.strictLocked
    );
    if (strictInSelection.length > 0) {
      const names = strictInSelection.map((s) => s.name || '(未命名)').slice(0, 3).join('、');
      const more = strictInSelection.length > 3 ? ` 等 ${strictInSelection.length} 个` : '';
      const warn = `🔒 ${strictInSelection.length} 个子启用了防误触（${names}${more}），请先手动关闭防误触后再建布局`;
      return { ok: false, applied: 0, failed: 0, warn, strictShapes: strictInSelection.length };
    }

    // PowerPoint.run + 业务逻辑全部走 radius-core.applyLayout（driver 模式）
    try {
      const result = await PowerPoint.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        return await window.RadiusCore.applyLayout(driver, parentId, params, childIds, {
          writeParentTag,
          syncR,
        });
      });
      return result;
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      return { ok: false, applied: 0, failed: 0, warn: '', error: msg };
    }
  }

  // 只同步 layout 子形状的 R 角（父 R 角被改时调用）
  // parentId: 父 id
  // paddingCm: 边距（cm）
  // linkRMode: 'subtract' | 'same' | 'off'
  // parentRcm: 父当前 R 角（cm）
  // 只在当前 slide 操作（不跨页）
  // 用统一函数 writeRadiusToChildren 自动处理 strict/lock 同步
  // v1.3.2 迁移：PowerPoint.run 部分全部走 radius-core.syncLayoutChildrenR
  async function syncLayoutChildrenR(parentId, childIds, paddingCm, linkRMode, parentRcm) {
    try {
      return await PowerPoint.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        return await window.RadiusCore.syncLayoutChildrenR(driver, parentId, childIds, paddingCm, linkRMode, parentRcm);
      });
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      return { ok: false, applied: 0, failed: 0, error: msg };
    }
  }

  // 检测选区里是否有 layout 父 → 同步其子 R 角（onApply / applyPipette 末尾调用）
  // v1.2.7：改成调 applyLayoutToChildren —— 父 R 角变化时**重算 layout 几何**（autoPadding 触发时子位置/尺寸也变）
  // v1.3.6：v1.2 step 3 调试 log 清理
  async function syncLayoutChildrenRIfNeeded() {
    if (selectedShapes.length === 0) return;
    for (const s of selectedShapes) {
      if (s.layoutRole === 'parent' && s.layoutParams && s.layoutChildIds) {
        // v1.2.6：默认从 'subtract' 改成 'same'（等距公式 R_sub = R_父，v1.0/v1.2 的 subtract 公式 R_sub = R_父 - d 几何上**不等距**）
        // 老 layout（tag 里存了 'subtract'）直接拿到旧值不受影响
        const linkRMode = s.layoutParams.linkRMode || 'same';
        if (linkRMode === 'off') continue;
        const expected = s.layoutParams.rows * s.layoutParams.cols;
        const childIds = s.layoutChildIds.slice(0, expected);
        // v1.2.7：调 applyLayoutToChildren 而不是 syncLayoutChildrenR
        // 原因：autoPadding 需要重新算 layout 几何（子位置/尺寸），syncLayoutChildrenR 只写 R 角
        // writeParentTag=false（父 tag 不重写——tag 里存的是 d_init，effectivePadding 是动态算的）
        // syncR=true（重写子 R 角）
        const params = { ...s.layoutParams, linkRMode };
        await applyLayoutToChildren(s.id, params, childIds, { writeParentTag: false, syncR: true });
      }
    }
  }

  // ---------------- v1.2: layout UI 渲染 + 交互 ----------------

  // 节流：滑块拖动时只在最后一次输入后 apply
  let layoutApplyTimer = null;
  let layoutApplyPending = false;

  // 方案 A：用户手动指定父/子。{ parentId, childIds[] }
  // 选区变化时自动重置（renderLayoutSetupList 检测 roundRect 列表变化）
  let layoutSetupChoices = { parentId: null, childIds: [] };
  let layoutSetupListSignature = '';  // 上次渲染的列表签名（用 roundRect ids 拼接），变了就重置 choices

  function scheduleLayoutApply() {
    layoutApplyPending = true;
    if (layoutApplyTimer) clearTimeout(layoutApplyTimer);
    layoutApplyTimer = setTimeout(() => {
      layoutApplyTimer = null;
      if (layoutApplyPending && currentLayout) {
        layoutApplyPending = false;
        applyLayoutFromUI({ writeParentTag: true });
      }
    }, LAYOUT_LAYOUT_RT_DEBOUNCE_MS);
  }

  // 立刻 apply（用户改 rows/cols 触发，因为会改变 childIds 数量）
  async function applyLayoutFromUI(opts) {
    if (!currentLayout) return;
    const parentId = currentLayout.parentId;
    const params = currentLayout.params;
    const childIds = currentLayout.childIds;
    stopLockMonitor();
    const r = await applyLayoutToChildren(parentId, params, childIds, opts || { writeParentTag: true, syncR: true });
    if (!r.ok) {
      showToast('布局应用失败：' + (r.error || '未知错误'));
    } else if (r.warn) {
      if (r.warn.indexOf('子形状不足') >= 0) {
        const actualRows = Math.max(1, Math.floor(childIds.length / params.cols));
        currentLayout.params.rows = actualRows;
        renderLayoutPanel();
        showToast(r.warn + ' — 自动缩减到 ' + actualRows + ' 行');
        return applyLayoutFromUI({ writeParentTag: true, syncR: true });
      }
      showToast(r.warn);
    } else {
      const rHint = params.linkRMode && params.linkRMode !== 'off' ? '（含 R 角联动）' : '';
      const strictHint = r.strictOverridden > 0
        ? `（${r.strictOverridden} 个原本是防误触的，已强制更新）`
        : '';
      showToast(`✅ 布局已应用 ${r.applied} 个子形状${r.failed ? `，${r.failed} 个失败` : ''}${rHint}${strictHint}`);
    }
    await refreshSelection();
    if (selectedShapes.length > 0) startLockMonitor();
  }

  // 渲染 layout 面板：根据选区状态切到 empty / active / child-info
  function renderLayoutPanel() {
    const empty = $('layout-empty');
    const active = $('layout-active');
    const childInfo = $('layout-child-info');
    const hint = $('layout-hint');
    if (!empty || !active || !childInfo) return;

    empty.style.display = 'none';
    active.style.display = 'none';
    childInfo.style.display = 'none';

    // 没选 / 选区无 roundRect
    if (selectedShapes.length === 0) {
      hint.textContent = '在 PPT 里选 1+ 个圆角矩形';
      empty.style.display = 'flex';
      $('layout-setup-btn').disabled = true;
      return;
    }
    // 选区里有 layout 父 → 显示 active 面板
    const parentShape = selectedShapes.find((s) => s.layoutRole === 'parent');
    if (parentShape && currentLayout) {
      hint.textContent = '已激活布局';
      active.style.display = 'flex';
      $('layout-parent-name').textContent = currentLayout.parentName;
      $('layout-children-count').textContent = `${currentLayout.childIds.length} 个（${currentLayout.params.rows}×${currentLayout.params.cols}）`;
      // 滑块 + 数字输入填值
      const rowsR = $('layout-rows');
      const rowsN = $('layout-rows-num');
      const colsReadout = $('layout-cols-readout');
      const coupledHint = $('layout-coupled-hint');
      const rowsDatalist = $('layout-rows-ticks');
      const padR = $('layout-padding');
      const padN = $('layout-padding-num');
      const gutR = $('layout-gutter');
      const gutN = $('layout-gutter-num');
      const warn = $('layout-warn');
      // v1.2.13：行的可取值 = N 的所有因子（不是 [1, N] 连续范围）
      // 例：N=4 → [1, 2, 4]；N=6 → [1, 2, 3, 6]；N=12 → [1, 2, 3, 4, 6, 12]
      const N = currentLayout.childIds.length;
      const factors = N > 0 ? window.RadiusCore.computeGridFactors(N) : [1];
      const maxFactor = factors[factors.length - 1];
      if (N > 0) {
        rowsR.max = String(maxFactor);
        rowsN.max = String(maxFactor);
        // 更新 datalist 的 tick（slider 上显示刻度提示用户）
        if (rowsDatalist) {
          rowsDatalist.innerHTML = factors.map((f) => `<option value="${f}"></option>`).join('');
        }
        // N=1 时行=1, 列=1 唯一；slider 禁用
        if (N === 1) {
          rowsR.disabled = true;
          rowsN.disabled = true;
          if (coupledHint) coupledHint.textContent = '（只有 1 个子）';
        } else {
          rowsR.disabled = false;
          rowsN.disabled = false;
          if (coupledHint) coupledHint.textContent = `（自动 = ${N} ÷ 行，N=${N} 的因子: ${factors.join('/')}）`;
        }
      }
      // 同步：params 里的 rows/cols 应该等于 N 的因子（如果 tag 损坏 → 默认 row=N, col=1）
      if (N > 0 && currentLayout.params.rows * currentLayout.params.cols !== N) {
        currentLayout.params.rows = N;
        currentLayout.params.cols = 1;
      }
      rowsR.value = String(currentLayout.params.rows);
      rowsN.value = String(currentLayout.params.rows);
      if (colsReadout) colsReadout.textContent = String(currentLayout.params.cols);
      padR.value = String(currentLayout.params.padding);
      padN.value = currentLayout.params.padding.toFixed(2);
      gutR.value = String(currentLayout.params.gutter);
      gutN.value = currentLayout.params.gutter.toFixed(2);
      // 选对应 R 角联动模式（v1.2.6 默认 'same' 等距，v1.0/v1.2 默认 'subtract' 不等距——见 radius-core 注释）
      const linkRMode = currentLayout.params.linkRMode || 'same';
      document.querySelectorAll('input[name="layout-link-r-mode"]').forEach((r) => {
        r.checked = r.value === linkRMode;
      });
      // 警告文本
      const minSubW = (() => {
        const p = parentShape;
        if (!p || p.width <= 0 || p.height <= 0) return 0;
        const totalW = p.width - 2 * currentLayout.params.padding * PT_PER_CM - (currentLayout.params.cols - 1) * currentLayout.params.gutter * PT_PER_CM;
        const totalH = p.height - 2 * currentLayout.params.padding * PT_PER_CM - (currentLayout.params.rows - 1) * currentLayout.params.gutter * PT_PER_CM;
        if (totalW <= 0 || totalH <= 0) return 0;
        return Math.min(totalW / currentLayout.params.cols, totalH / currentLayout.params.rows) / PT_PER_CM;
      })();
      const childCount = currentLayout.childIds.length;
      const expected = currentLayout.params.rows * currentLayout.params.cols;
      if (minSubW <= 0) {
        warn.textContent = '⚠️ 边距/间距太大，挤不下';
      } else if (childCount < expected) {
        // v1.2.11：理论上 N=rows*cols（联动保证），但旧的 layout tag 可能不对
        warn.textContent = `⚠️ 子形状不足（需要 ${expected}，找到 ${childCount}）`;
      } else {
        warn.textContent = '';
      }
      updateLayoutPreview();
      return;
    }
    // 选区里只有子（无父）
    const childShape = selectedShapes.find((s) => s.layoutRole === 'child');
    if (childShape) {
      hint.textContent = '当前形状是布局子项';
      childInfo.style.display = 'flex';
      const parentInSel = selectedShapes.find((s) => s.id === childShape.layoutParentId);
      $('layout-child-parent').textContent = parentInSel ? (parentInSel.name || '(未命名)') : '（已不在选区）';
      return;
    }
    // 选区里没 layout：显示 setup + 方案 A 列表
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    renderLayoutSetupList(roundShapes);
    const rows = parseInt($('layout-setup-rows').value, 10);
    const cols = parseInt($('layout-setup-cols').value, 10);
    const need = (Number.isFinite(rows) ? rows : 1) * (Number.isFinite(cols) ? cols : 1);
    const rowsOk = Number.isFinite(rows) && rows >= 1 && rows <= 5;
    const colsOk = Number.isFinite(cols) && cols >= 1 && cols <= 5;
    // 防误触检查：选区里有任何 roundRect 是 strictLocked → 整个 setup 拒绝（"进入组合时"判断）
    const strictCount = roundShapes.filter((s) => s.strictLocked).length;
    const canBuild = !!layoutSetupChoices.parentId
      && layoutSetupChoices.childIds.length >= need
      && rowsOk && colsOk
      && strictCount === 0;
    $('layout-setup-btn').disabled = !canBuild;
    if (roundShapes.length === 0) {
      hint.textContent = '在 PPT 里选 1+ 个圆角矩形';
    } else if (strictCount > 0) {
      // 防误触：选区里有 N 个子启用了防误触 → 拒绝整个 setup
      hint.textContent = `🔒 ${strictCount} 个启用了防误触，请先关闭后再建布局`;
      $('layout-setup-btn').disabled = true;  // 再次确认按钮禁用
    } else if (!canBuild) {
      const missing = !layoutSetupChoices.parentId
        ? '请指定一个父'
        : layoutSetupChoices.childIds.length < need
          ? `子不足（需要 ${need}，选了 ${layoutSetupChoices.childIds.length}）`
          : '检查行/列范围';
      hint.textContent = missing;
    } else {
      hint.textContent = `✅ 可以建 ${rows}×${cols} 布局`;
    }
    empty.style.display = 'flex';
  }

  // 滑块 / 数字输入同步 + 触发 apply
  function bindLayoutRangeAndNum(rangeId, numId, paramKey, isInt) {
    const r = $(rangeId);
    const n = $(numId);
    if (!r || !n) return;
    r.addEventListener('input', () => {
      n.value = r.value;
      if (!currentLayout) return;
      currentLayout.params[paramKey] = isInt ? parseInt(r.value, 10) : parseFloat(r.value);
      updateLayoutPreview(); // 拖动时立即更新子尺寸预览（不等 apply）
      scheduleLayoutApply();
    });
    n.addEventListener('input', () => {
      let v = isInt ? parseInt(n.value, 10) : parseFloat(n.value);
      if (!Number.isFinite(v)) return;
      v = Math.max(parseFloat(r.min), Math.min(parseFloat(r.max), v));
      r.value = String(v);
      if (!currentLayout) return;
      currentLayout.params[paramKey] = v;
      updateLayoutPreview();
      scheduleLayoutApply();
    });
    n.addEventListener('change', () => {
      if (currentLayout && layoutApplyTimer) {
        clearTimeout(layoutApplyTimer);
        layoutApplyTimer = null;
        layoutApplyPending = false;
        applyLayoutFromUI({ writeParentTag: true, syncR: true });
      }
    });
  }

  /**
   * v1.2.11：行/列互斥联动绑定（单滑块 UI 版）
   *
   * 设计原则：
   *   - 行 × 列 = 子数 N（layout 创建后 N 固定，4 个子就是 4）
   *   - 改 rows → cols = max(1, ceil(N / rows))，列 readout 同步
   *   - 列不再有独立控件（只有 readout 文本显示），避免用户以为列能改
   *   - slider max = N（不是硬编码 5）；N=1 时 slider 禁用（1×1 唯一）
   *   - 子图形的行/列方向 = rows/cols，N 不会随 rows 变（不动 N）
   *
   * 例：N=4，改 rows=1 → cols=4 → 1×4（4 个子全用上）
   *     N=4，改 rows=3 → cols=2 → 3×2（4 个子全用上）
   *     N=4，改 rows=5 → clamp 到 4 → cols=1 → 4×1
   *
   * 父图形的 size 不变（这是用户明确要求）。computeLayout 已经在算 subW/subH 时按
   * 当前父 size + rows/cols 算，新 rows/cols 触发后子位置/尺寸自动跟上。
   */
  function bindLayoutGridRangeAndNum(rowsRangeId, rowsNumId, colsReadoutId) {
    const rowsR = $(rowsRangeId);
    const rowsN = $(rowsNumId);
    const colsReadout = $(colsReadoutId);
    if (!rowsR || !rowsN) return;
    if (!colsReadout) {
      // readout 缺失就 silent skip（理论上不应发生）
      console.log('[layout-grid] WARN colsReadout 元素缺失，列显示将不更新');
    }

    // 改 rows → 自动算 cols
    // v1.2.13：snap 到 N 的最近因子（行滑块是离散 list，不是连续 range）
    function handleRowsChange(rawVal) {
      if (!currentLayout) return;
      const N = currentLayout.childIds.length;
      if (!Number.isFinite(N) || N <= 0) return;
      // 先 snap 到最近因子
      const factors = window.RadiusCore.computeGridFactors(N);
      const snapped = window.RadiusCore.snapToNearestGridFactor(rawVal, factors);
      // 再算联动
      const coupled = window.RadiusCore.computeGridCoupledRowsCols('rows', snapped, N);
      // 写 params
      currentLayout.params.rows = coupled.rows;
      currentLayout.params.cols = coupled.cols;
      // 同步 UI
      rowsR.value = String(coupled.rows);
      rowsN.value = String(coupled.rows);
      if (colsReadout) colsReadout.textContent = String(coupled.cols);
      updateLayoutPreview();
      scheduleLayoutApply();
    }

    // 监听 rows（slider + number input 双向）
    rowsR.addEventListener('input', () => handleRowsChange(rowsR.value));
    rowsN.addEventListener('input', () => handleRowsChange(rowsN.value));
    // 失去焦点时如果有 pending timer → 立即 flush（不等节流）
    rowsN.addEventListener('change', () => {
      if (currentLayout && layoutApplyTimer) {
        clearTimeout(layoutApplyTimer);
        layoutApplyTimer = null;
        layoutApplyPending = false;
        applyLayoutFromUI({ writeParentTag: true, syncR: true });
      }
    });
  }

  // 渲染方案 A 列表：选区里所有 roundRect，每行一个「父/子」radio
  // 选区变化（roundRect 列表变了）→ 重置 layoutSetupChoices（默认第一个为父）
  function renderLayoutSetupList(roundShapes) {
    const list = $('layout-setup-list');
    if (!list) return;
    const sig = roundShapes.map((s) => s.id).join('|');
    if (sig !== layoutSetupListSignature) {
      // 选区变了：重置（默认第一个为父）
      layoutSetupListSignature = sig;
      if (roundShapes.length > 0) {
        layoutSetupChoices.parentId = roundShapes[0].id;
        layoutSetupChoices.childIds = roundShapes.slice(1).map((s) => s.id);
      } else {
        layoutSetupChoices.parentId = null;
        layoutSetupChoices.childIds = [];
      }
    }
    list.innerHTML = '';
    if (roundShapes.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'layout-setup-empty';
      empty.textContent = '在 PPT 里选 1+ 个圆角矩形';
      list.appendChild(empty);
      return;
    }
    for (const sh of roundShapes) {
      const row = document.createElement('div');
      row.className = 'layout-setup-row';
      // strict 形状：红框 + 提示
      if (sh.strictLocked) {
        row.classList.add('is-strict');
        row.title = '🔒 此形状启用了防误触，请先在「防误触」区域关闭';
      }
      const isParent = layoutSetupChoices.parentId === sh.id;
      const isChild = layoutSetupChoices.childIds.includes(sh.id);
      const nameSpan = document.createElement('span');
      nameSpan.className = 'layout-setup-name';
      nameSpan.textContent = sh.name || '(未命名)';
      // strict 标记
      if (sh.strictLocked) {
        const lockBadge = document.createElement('span');
        lockBadge.className = 'layout-setup-strict-badge';
        lockBadge.textContent = '🔒';
        nameSpan.appendChild(document.createTextNode(' '));
        nameSpan.appendChild(lockBadge);
      }
      const parentLabel = document.createElement('label');
      parentLabel.className = 'layout-setup-radio' + (isParent ? ' is-parent' : '');
      const parentRadio = document.createElement('input');
      parentRadio.type = 'radio';
      parentRadio.name = 'layout-role-' + sh.id;
      parentRadio.dataset.shapeId = sh.id;
      parentRadio.value = 'parent';
      parentRadio.checked = isParent;
      const parentSpan = document.createElement('span');
      parentSpan.textContent = '父';
      parentLabel.appendChild(parentRadio);
      parentLabel.appendChild(parentSpan);
      parentRadio.addEventListener('change', () => {
        if (!parentRadio.checked) return;
        layoutSetupChoices.parentId = sh.id;
        // 旧父 → 子
        layoutSetupChoices.childIds = layoutSetupChoices.childIds.filter((id) => id !== sh.id);
        // 当前 shape 之前如果是子，从 childIds 移除
        renderLayoutSetupList(roundShapes);
        renderLayoutPanel();
      });
      const childLabel = document.createElement('label');
      childLabel.className = 'layout-setup-radio' + (isChild ? ' is-child' : '');
      const childRadio = document.createElement('input');
      childRadio.type = 'radio';
      childRadio.name = 'layout-role-' + sh.id;
      childRadio.dataset.shapeId = sh.id;
      childRadio.value = 'child';
      childRadio.checked = isChild;
      const childSpan = document.createElement('span');
      childSpan.textContent = '子';
      childLabel.appendChild(childRadio);
      childLabel.appendChild(childSpan);
      childRadio.addEventListener('change', () => {
        if (!childRadio.checked) return;
        // 如果当前是父 → 切到子：把 parentId 置空
        if (layoutSetupChoices.parentId === sh.id) {
          layoutSetupChoices.parentId = null;
        }
        if (!layoutSetupChoices.childIds.includes(sh.id)) {
          layoutSetupChoices.childIds.push(sh.id);
        }
        renderLayoutSetupList(roundShapes);
        renderLayoutPanel();
      });
      row.appendChild(nameSpan);
      row.appendChild(parentLabel);
      row.appendChild(childLabel);
      list.appendChild(row);
    }
  }

  // 更新 active 面板里的「子尺寸 X.XX × Y.YY cm」预览
  function updateLayoutPreview() {
    const el = $('layout-preview-size');
    if (!el) return;
    if (!currentLayout) { el.textContent = '—'; return; }
    const parent = selectedShapes.find((s) => s.id === currentLayout.parentId);
    if (!parent || !parent.width || !parent.height) { el.textContent = '—'; return; }
    const r = computeLayout(
      { left: parent.left || 0, top: parent.top || 0, width: parent.width, height: parent.height },
      currentLayout.params.rows,
      currentLayout.params.cols,
      currentLayout.params.padding,
      currentLayout.params.gutter
    );
    if (!r.feasible) {
      el.textContent = '⚠️ 挤不下';
      return;
    }
    el.textContent = `${(r.subW / PT_PER_CM).toFixed(2)} × ${(r.subH / PT_PER_CM).toFixed(2)} cm`;
  }

  // 从选区建立布局：使用 layoutSetupChoices（用户手动指定的父/子）
  async function onLayoutSetup() {
    const rows = parseInt($('layout-setup-rows').value, 10);
    const cols = parseInt($('layout-setup-cols').value, 10);
    if (!Number.isFinite(rows) || rows < 1 || rows > 5) {
      showToast('行数范围 1~5');
      return;
    }
    if (!Number.isFinite(cols) || cols < 1 || cols > 5) {
      showToast('列数范围 1~5');
      return;
    }
    const need = rows * cols;
    if (!layoutSetupChoices.parentId) {
      showToast('请在列表里指定一个父');
      return;
    }
    if (layoutSetupChoices.childIds.length < need) {
      showToast(`子不足（需要 ${need}，选了 ${layoutSetupChoices.childIds.length}）`);
      return;
    }
    const parentId = layoutSetupChoices.parentId;
    const childIds = layoutSetupChoices.childIds.slice(0, need);
    // v1.2.6：layout 新建时默认 linkRMode = 'same'（等距 R_sub = R_父）—— 之前 'subtract' 几何上不等距
    const params = { rows, cols, padding: 0.5, gutter: 0.3, linkRMode: 'same' };
    stopLockMonitor();
    const r = await applyLayoutToChildren(parentId, params, childIds, { writeParentTag: true, syncR: true });
    if (!r.ok) {
      showToast('建布局失败：' + (r.error || r.warn || '未知错误'));
    } else if (r.applied === 0) {
      showToast('⚠️ 没写成功任何子形状：' + (r.warn || '未知问题，看 console'));
    } else {
      showToast(`🎯 已建立 ${rows}×${cols} 布局（${r.applied} 个子形状${r.failed ? '，' + r.failed + ' 个失败' : ''}）`);
    }
    await refreshSelection();
    if (selectedShapes.length > 0) startLockMonitor();
  }

  // 脱离布局：删父 + 子的 layout tag
  async function onLayoutDetach() {
    if (!currentLayout) return;
    const parentId = currentLayout.parentId;
    const childIds = currentLayout.childIds;
    stopLockMonitor();
    const r = await deleteLayoutTags(parentId, childIds);
    if (!r.ok) {
      showToast('脱离失败：' + (r.error || '未知错误'));
    } else {
      showToast('已脱离布局（子形状的位置/尺寸/R 角保留）');
    }
    await refreshSelection();
    if (selectedShapes.length > 0) startLockMonitor();
  }

  // 子形状脱离（只删自己的 child tag + 从父的 childIds 移除）
  // 只在当前 slide 操作（不跨页）
  async function onLayoutChildDetach() {
    const childShape = selectedShapes.find((s) => s.layoutRole === 'child');
    if (!childShape) return;
    const parentId = childShape.layoutParentId;
    stopLockMonitor();
    try {
      await PowerPoint.run(async (ctx) => {
        const activeSlide = ctx.presentation.getSelectedSlides().getItemAt(0);
        activeSlide.load('shapes/items/id');
        await ctx.sync();
        const idToShape = new Map();
        for (const sh of activeSlide.shapes.items) {
          idToShape.set(sh.id, sh);
        }
        const csh = idToShape.get(childShape.id);
        if (csh) {
          try { csh.tags.delete(LAYOUT_CHILD_TAG_KEY); } catch (_) {}
        }
        const parentShape = idToShape.get(parentId);
        if (parentShape) {
          try {
            const t = parentShape.tags.getItem(LAYOUT_PARENT_TAG_KEY);
            t.load('value');
            await ctx.sync();
            const obj = JSON.parse(t.value);
            obj.childIds = (obj.childIds || []).filter((x) => x !== childShape.id);
            parentShape.tags.add(LAYOUT_PARENT_TAG_KEY, JSON.stringify(obj));
          } catch (_) {}
        }
        await ctx.sync();
      });
      showToast('已脱离此布局');
    } catch (e) {
      showToast('脱离失败：' + (e.message || e));
    }
    await refreshSelection();
    if (selectedShapes.length > 0) startLockMonitor();
  }

  // ---------------- v1.2: 调试日志面板（task pane 底部，默认折叠） ----------------

  const DEBUG_LOG_MAX = 200;
  function addDebugLog(level, ...args) {
    const body = $('debug-log-body');
    if (!body) return;
    const line = document.createElement('div');
    line.className = 'log-line log-' + (level || 'info');
    const ts = new Date().toLocaleTimeString('zh-CN', { hour12: false });
    const msg = args.map((a) => {
      if (typeof a === 'string') return a;
      try { return JSON.stringify(a); } catch (_) { return String(a); }
    }).join(' ');
    line.textContent = `[${ts}] ${msg}`;
    body.appendChild(line);
    // 限制行数
    while (body.children.length > DEBUG_LOG_MAX) {
      body.removeChild(body.firstChild);
    }
    // 自动滚到底
    body.scrollTop = body.scrollHeight;
  }
  // 替换 console.log/warn/error（保留原始 console 用于 Safari Inspector）
  const _origLog = console.log;
  const _origWarn = console.warn;
  const _origError = console.error;
  console.log = function () { addDebugLog('info', ...arguments); _origLog.apply(console, arguments); };
  console.warn = function () { addDebugLog('warn', ...arguments); _origWarn.apply(console, arguments); };
  console.error = function () { addDebugLog('error', ...arguments); _origError.apply(console, arguments); };

  // 复制按钮：把整个日志复制到剪贴板（不触发 details toggle）
  function copyDebugLog() {
    const body = $('debug-log-body');
    if (!body) return;
    const lines = [];
    for (const child of body.children) {
      lines.push(child.textContent);
    }
    const text = lines.join('\n');
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => {
        showToast(`📋 已复制 ${lines.length} 行日志`);
      }).catch((err) => {
        // fallback
        fallbackCopy(text);
        showToast(`📋 已复制 ${lines.length} 行（fallback）`);
      });
    } else {
      fallbackCopy(text);
      showToast(`📋 已复制 ${lines.length} 行（fallback）`);
    }
  }
  function fallbackCopy(text) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); } catch (_) {}
    document.body.removeChild(ta);
  }
  function clearDebugLog() {
    const body = $('debug-log-body');
    if (!body) return;
    body.innerHTML = '';
    showToast('🗑️ 调试日志已清空');
  }

  // ============================================================
  // Driver 烟囱测试（v1.2.3）
  // 跑遍所有 15 个 driver 方法 + 在真实 PPT 里读/写一遍
  // 每个方法的输入/输出都打 log，复制整段发我就能看到每个模块状态
  // ============================================================
  async function runDriverSmokeTest() {
    console.log('[smoke] ===== Driver 烟囱测试开始 =====');
    console.log('[smoke] 时间: ' + new Date().toLocaleString());
    console.log('[smoke] driver 方法数: 16（' +
      'load, sync, selectedShapes, activeSlide, slideShapes, ' +
      'shapeId, size, box, isRoundRect, adjFraction, loadAdjValue, ' +
      'setBox, setAdjFraction, addTag, deleteTag, readTag)');
    try {
      await PowerPoint.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        const results = { pass: 0, fail: 0, tests: [] };
        function record(name, ok, detail) {
          if (ok) results.pass++;
          else results.fail++;
          results.tests.push({ name, ok, detail });
          console.log(`[smoke] ${ok ? '✅' : '❌'} ${name}: ${detail}`);
        }

        // 1. Collection accessors
        console.log('[smoke] --- 1. Collection accessors ---');
        let sel, slide, slideShapes;
        try {
          sel = driver.selectedShapes();
          record('selectedShapes()', !!sel, 'returned proxy (no throw)');
        } catch (e) {
          record('selectedShapes()', false, 'threw: ' + e.message);
          return;
        }
        try {
          slide = driver.activeSlide();
          record('activeSlide()', !!slide, 'returned proxy');
        } catch (e) {
          record('activeSlide()', false, 'threw: ' + e.message);
          return;
        }
        try {
          slideShapes = driver.slideShapes(slide);
          record('slideShapes(slide)', !!slideShapes, 'returned proxy');
        } catch (e) {
          record('slideShapes(slide)', false, 'threw: ' + e.message);
          return;
        }

        // 2. Load + sync
        console.log('[smoke] --- 2. Load + sync ---');
        try {
          driver.load(sel, 'items/id, items/width, items/height, items/left, items/top, items/adjustments, items/tags');
          record('load(sel, fields)', true, '7 fields queued');
        } catch (e) {
          record('load(sel, fields)', false, 'threw: ' + e.message);
          return;
        }
        try {
          await driver.sync();
          record('sync()', true, 'no throw');
        } catch (e) {
          record('sync()', false, 'threw: ' + e.message);
          return;
        }

        // v1.2.5：per-shape 显式 load adjustments value（Mac LTSC 必加）
        let adjValueLoadCount = 0;
        for (const sh of sel.items) {
          if (driver.isRoundRect(sh)) {
            driver.loadAdjValue(sh);
            adjValueLoadCount++;
          }
        }
        try {
          await driver.sync();
          if (adjValueLoadCount > 0) {
            console.log('[smoke]   loadAdjValue called for ' + adjValueLoadCount + ' roundRect shapes');
          }
        } catch (e) {
          console.log('[smoke] ⚠️ loadAdjValue + sync failed: ' + e.message);
        }

        // 3. Check selection
        console.log('[smoke] --- 3. 选区检查 ---');
        const itemCount = sel.items ? sel.items.length : 0;
        if (itemCount === 0) {
          console.log('[smoke] ⚠️ 选区为空 — 没法测后续方法。请先在 PPT 里选 1-2 个圆角矩形再点此按钮。');
          return;
        }
        console.log('[smoke] 选区里有 ' + itemCount + ' 个 shape');

        // 4. Read methods (per shape)
        console.log('[smoke] --- 4. 读方法（per shape）---');
        const shapes = [];
        for (let i = 0; i < sel.items.length; i++) {
          const sh = sel.items[i];
          const shId = (() => { try { return driver.shapeId(sh); } catch (e) { return 'ERROR: ' + e.message; } })();
          const isRR = (() => { try { return driver.isRoundRect(sh); } catch (e) { return 'ERROR: ' + e.message; } })();
          const adjF = (() => { try { return driver.adjFraction(sh); } catch (e) { return 'ERROR: ' + e.message; } })();
          const sz = (() => { try { return driver.size(sh); } catch (e) { return 'ERROR: ' + e.message; } })();
          const bx = (() => { try { return driver.box(sh); } catch (e) { return 'ERROR: ' + e.message; } })();
          console.log(`[smoke]   shape[${i}] id=${shId} isRoundRect=${isRR} adjFraction=${adjF} size=${JSON.stringify(sz)} box=${JSON.stringify(bx)}`);
          shapes.push({ sh, shId, isRR, adjF, sz, bx });
        }
        record('shapeId()', shapes.every((s) => typeof s.shId === 'string'),
          'all shapes returned string id (' + shapes.map((s) => s.shId).join(',') + ')');
        record('isRoundRect()', shapes.every((s) => typeof s.isRR === 'boolean'),
          'all shapes returned boolean');
        record('adjFraction()', shapes.every((s) => typeof s.adjF === 'number'),
          'all shapes returned number (0~1)');
        record('size()', shapes.every((s) => s.sz && typeof s.sz.width === 'number'),
          'all shapes returned {width, height}');
        record('box()', shapes.every((s) => s.bx && typeof s.bx.left === 'number'),
          'all shapes returned {left, top, width, height}');

        // 5. Tag operations（找一个 roundRect shape 来测）
        console.log('[smoke] --- 5. Tag 操作（addTag/readTag/deleteTag）---');
        const testShape = shapes.find((s) => s.isRR && s.shId !== 'ERROR');
        if (!testShape) {
          console.log('[smoke] ⚠️ 选区里没有圆角矩形 — 跳过 tag 写测试');
        } else {
          const TEST_KEY = 'driver_smoke_test_v1';
          const TEST_VAL = 'hello_' + Date.now();
          try {
            driver.addTag(testShape.sh, TEST_KEY, TEST_VAL);
            await driver.sync();
            const readBack = await driver.readTag(testShape.sh, TEST_KEY);
            record('addTag + sync + readTag', readBack === TEST_VAL,
              `wrote "${TEST_VAL}", read back "${readBack}"`);
            driver.deleteTag(testShape.sh, TEST_KEY);
            await driver.sync();
            const afterDel = await driver.readTag(testShape.sh, TEST_KEY);
            record('deleteTag + readTag', afterDel == null,
              `after delete, readTag returned ${JSON.stringify(afterDel)}`);
          } catch (e) {
            record('tag operations', false, 'threw: ' + e.message);
          }
        }

        // 6. setAdjFraction（找一个 roundRect shape 来测）
        console.log('[smoke] --- 6. setAdjFraction（写 R 角再读回）---');
        if (testShape) {
          try {
            const origAdj = testShape.adjF;
            const testAdj = origAdj > 0.5 ? 0.05 : 0.5;  // flip between small and big
            // v1.3.1 模式：set + sync → get(0) + read（v1.0 模式适配 set+read）
            // （v1.2.9 的 `load + sync + get` 还是不 work——load 是给没 set 过的 shape 准备 value 用的，
            //   set 之后 value 已经在 PPT 上了，只需要 fresh get(0) 把新值拉到 proxy 即可）
            // 关键：get(0) 必须在 sync 之后（不存变量），因为 set 的 sync 会 invalidate 旧 proxy
            let readBack = null;
            let readException = null;
            try {
              driver.setAdjFraction(testShape.sh, testAdj);
              await driver.sync();
              readBack = testShape.sh.adjustments.get(0).value;  // ← fresh get(0) AFTER sync
            } catch (e) {
              readException = e.message || String(e);
            }
            if (readBack != null) {
              record('setAdjFraction + sync + adjFraction', Math.abs(readBack - testAdj) < 0.001,
                `wrote ${testAdj}, read back ${readBack.toFixed(4)} (orig was ${origAdj})`);
            } else {
              // 同 run 读失败，跨 run 兜底（v1.2.9 之前测的「同 run 不可靠」就靠这个）
              console.log('[smoke] ⚠️ same-run read failed:', readException, '— trying cross-run read');
              let crossRunRead = 0;
              let crossRunErr = null;
              try {
                await PowerPoint.run(async (ctx) => {
                  const d2 = window.PptDriver.createDriver(ctx);
                  const sel2 = d2.selectedShapes();
                  d2.load(sel2, 'items/id, items/adjustments');
                  await d2.sync();
                  for (const sh of sel2.items) {
                    if (d2.shapeId(sh) === testShape.shId) {
                      const adjResult = sh.adjustments.get(0);
                      await d2.sync();
                      crossRunRead = adjResult.value;
                      break;
                    }
                  }
                });
                record('setAdjFraction + sync + adjFraction', Math.abs(crossRunRead - testAdj) < 0.001,
                  `wrote ${testAdj}, cross-run read back ${crossRunRead.toFixed(4)} (orig was ${origAdj}) — same-run 不可靠，跨 run 兜底成功`);
              } catch (e2) {
                crossRunErr = e2.message || String(e2);
                record('setAdjFraction + sync + adjFraction', false,
                  `set ok 但读失败：same-run: ${readException} | cross-run: ${crossRunErr}`);
              }
            }
            // Restore original
            try {
              driver.setAdjFraction(testShape.sh, origAdj);
              await driver.sync();
            } catch (_) {}
            // 跨 run 读一下 orig 验证
            let restoredVal = origAdj;
            try {
              await PowerPoint.run(async (ctx) => {
                const d2 = window.PptDriver.createDriver(ctx);
                const sel2 = d2.selectedShapes();
                d2.load(sel2, 'items/id, items/adjustments');
                await d2.sync();
                for (const sh of sel2.items) {
                  if (d2.shapeId(sh) === testShape.shId) {
                    const adjResult = sh.adjustments.get(0);
                    await d2.sync();
                    restoredVal = adjResult.value;
                    break;
                  }
                }
              });
            } catch (_) {}
            console.log(`[smoke]   restored to orig adj=${restoredVal.toFixed(4)} (was ${origAdj})`);
          } catch (e) {
            console.log('[smoke] ❌ setAdjFraction outer fail:', e.message || e);
            record('setAdjFraction', false, 'outer threw: ' + (e.message || e));
          }
        }

        // 7. setBox（用一个 shape 来测：先读 box，偏移 +10pt，再读回验证）
        console.log('[smoke] --- 7. setBox（写 left 偏移再读回）---');
        if (testShape) {
          try {
            const origBox = testShape.bx;
            const newBox = { left: origBox.left + 10, top: origBox.top, width: origBox.width, height: origBox.height };
            driver.setBox(testShape.sh, newBox);
            await driver.sync();
            const readBack = driver.box(testShape.sh);
            record('setBox + sync + box', Math.abs(readBack.left - newBox.left) < 0.5,
              `wrote left=${newBox.left}, read back ${readBack.left} (orig ${origBox.left})`);
            // Restore
            driver.setBox(testShape.sh, origBox);
            await driver.sync();
            const restored = driver.box(testShape.sh);
            console.log(`[smoke]   restored to orig left=${restored.left} (was ${origBox.left})`);
          } catch (e) {
            record('setBox', false, 'threw: ' + e.message);
          }
        }

        // 总结
        console.log('[smoke] ===== 总结 =====');
        console.log(`[smoke] ✅ pass: ${results.pass}, ❌ fail: ${results.fail}`);
        console.log('[smoke] 详细结果：');
        for (const t of results.tests) {
          console.log(`[smoke]   ${t.ok ? '✅' : '❌'} ${t.name}: ${t.detail}`);
        }
        console.log('[smoke] ===== Driver 烟囱测试结束 =====');
        showToast(`🧪 烟囱测试完成：${results.pass} 通过 / ${results.fail} 失败 — 看 debug 日志`);
      });
    } catch (e) {
      console.log('[smoke] FATAL: ' + (e.message || e));
      showToast('🧪 烟囱测试失败：' + (e.message || e));
    }
  }

  // ---------------- v1.2: layout tag 读写 ----------------

  // 读选区里所有 shape 的 layout tag
  // 返回：parents = { id: {rows, cols, padding, gutter, linkR, childIds} }
  //      childOf = { id: parentId }
  // v1.3.6 迁移：PowerPoint.run 部分走 radius-core.loadLayoutTags
  // thin wrapper：开 PowerPoint.run 拿 selectedShapes proxy → 调 radius-core
  // radius-core 做读 + stale state 检测，dialog.js 只负责 PowerPoint.run 上下文
  //
  // v1.3.7 修 bug：stale 检测必须用整 slide 的 shape IDs（不只是选区）
  // 之前只传 sel：只选父时 4 个子不在选区 → 误判全部 stale → childIds 变空 → UI 显示"子 0 个"
  function loadLayoutTagsViaTags() {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const driver = window.PptDriver.createDriver(ctx);
          const sel = ctx.presentation.getSelectedShapes();
          sel.load('items/id');
          // 拿整 slide 的 shapes —— 用于 stale 检测（v1.3.7 修复）
          const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
          slide.load('shapes/items/id');
          await ctx.sync();
          const r = await window.RadiusCore.loadLayoutTags(driver, sel, slide.shapes);
          resolve(r);
        } catch (e) {
          resolve({ ok: false, error: e });
        }
      });
    });
  }

  // 把 layout 信息写到父 shape 的 tag（包含所有参数 + childIds）
  // 也确保每个子都有 layoutChild_v1 tag 指向父
  // 只在当前 slide 操作（不跨页）
  // v1.3.6 迁移：PowerPoint.run 部分走 radius-core.saveLayoutTags
  function saveLayoutTags(parentId, params, childIds) {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const driver = window.PptDriver.createDriver(ctx);
          const activeSlide = ctx.presentation.getSelectedSlides().getItemAt(0);
          activeSlide.load('shapes/items/id');
          await ctx.sync();
          const r = await window.RadiusCore.saveLayoutTags(
            driver, activeSlide, parentId, params, childIds || []
          );
          resolve(r);
        } catch (e) {
          resolve({ ok: false, error: e });
        }
      });
    });
  }

  // 删 layout 父子 tag（保留形状本身的位置/尺寸/R 角）
  // 只在当前 slide 操作（不跨页）
  function deleteLayoutTags(parentId, childIds) {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const activeSlide = ctx.presentation.getSelectedSlides().getItemAt(0);
          activeSlide.load('shapes/items/id');
          await ctx.sync();
          for (const sh of activeSlide.shapes.items) {
            if (sh.id === parentId) {
              try { sh.tags.delete(LAYOUT_PARENT_TAG_KEY); } catch (_) {}
            }
            if (childIds && childIds.includes(sh.id)) {
              try { sh.tags.delete(LAYOUT_CHILD_TAG_KEY); } catch (_) {}
            }
          }
          await ctx.sync();
          resolve({ ok: true });
        } catch (e) {
          resolve({ ok: false, error: e });
        }
      });
    });
  }

  // ---------------- history（纯内存） ----------------

  function renderHistory(history) {
    const box = $('history-toggle');
    if (!box) return;
    box.innerHTML = '';
    const list = Array.isArray(history) ? history : [];
    // 始终显示 MAX_HISTORY 个槽位：前 N 个是真实记录，后 (MAX_HISTORY-N) 个是 disabled 占位
    for (let i = 0; i < MAX_HISTORY; i++) {
      const h = list[i];
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'history-btn';
      if (h) {
        btn.dataset.value = String(h.value);
        btn.dataset.unit = h.unit;
        const label = h.unit === '%'
          ? `${Number.isInteger(h.value) ? h.value : h.value.toFixed(1)}%`
          : h.value.toFixed(2);
        btn.textContent = label;
        btn.title = h.unit === '%'
          ? `${label}（点击填入输入框）`
          : `${label} cm（点击填入输入框）`;
        btn.addEventListener('click', () => onHistoryChipClick(h.value, h.unit));
      } else {
        btn.disabled = true;
        btn.textContent = '—';
        btn.title = '尚无记录';
      }
      box.appendChild(btn);
    }
  }

  function loadAndRenderHistory() {
    renderHistory(userHistory);
  }

  function pushHistory(value, unit) {
    // v1.3.6 迁移：直接走 radius-core.pushHistory（纯函数 + 不修改入参）
    userHistory = window.RadiusCore.pushHistory(userHistory, value, unit, MAX_HISTORY);
    return userHistory;
  }

  function onHistoryChipClick(value, unit) {
    if (unit !== currentUnit) onUnitChange(unit);
    $('radius-input').value = unit === '%'
      ? (Number.isInteger(value) ? value : value.toFixed(1))
      : value.toFixed(2);
    $('radius-input').focus();
  }

  // ---------------- 选区 + 读选中的 R 角 ----------------

  async function refreshSelection() {
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/name, items/width, items/height, items/adjustments');
        await ctx.sync();
        const shapes = [];
        for (const sh of sel.items) {
          const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
          // Mac LTSC task pane: get(0) 是 ClientResult 代理，先 get 再 sync
          // 让 value 自动填上（不需要显式 load items/value）
          const adjCount = sh.adjustments.count;
          const adjResult = sh.adjustments.get(0);
          await ctx.sync();
          let value = null;
          try { value = adjResult.value; } catch (_) { /* 不是 roundRect */ }
          const isRoundRect = (typeof adjCount === 'number' ? adjCount : 0) > 0;
          let cm = null;
          if (isRoundRect && Number.isFinite(value) && value > 0) {
            cm = value * minSideCm;
          }
          shapes.push({
            id: sh.id,
            name: sh.name,
            width: sh.width,
            height: sh.height,
            minSideCm,
            currentCm: cm,
            isRoundRect,
            locked: false,
            lockedCm: null,
            strictLocked: false,
            // v1.2 layout 字段
            layoutRole: null,        // 'parent' | 'child' | null
            layoutParentId: null,    // 自己作为子时，指向父 id
            layoutParams: null,      // 自己作为父时，{rows, cols, padding, gutter, linkRMode}
            layoutChildIds: null,    // 自己作为父时，[childId...]
          });
        }
        selectedShapes = shapes;
      });
      // 读 lock + strict 后端（shape.tags）
      const tagResult = await loadLocksViaTags();
      if (tagResult.ok) {
        for (const s of selectedShapes) {
          if (tagResult.locks[s.id] != null) {
            s.locked = true;
            s.lockedCm = tagResult.locks[s.id];
          }
          if (tagResult.strict[s.id]) {
            s.strictLocked = true;
          }
        }
      }
      // 读 layout tag
      const layoutResult = await loadLayoutTagsViaTags();
      if (layoutResult.ok) {
        // v1.3.6：检测到 stale childIds 时给用户提示（不自动写回，等下次调 saveLayoutTags 时自然清理）
        if (layoutResult.staleParents && Object.keys(layoutResult.staleParents).length > 0) {
          const totalStale = Object.values(layoutResult.staleParents).reduce((s, arr) => s + arr.length, 0);
          if (totalStale > 0) {
            showToast(`ℹ️ 检测到 ${totalStale} 个 layout 子已被删除/移走，下次应用布局时会自动清理`);
          }
        }
        for (const s of selectedShapes) {
          if (layoutResult.parents[s.id]) {
            s.layoutRole = 'parent';
            const p = layoutResult.parents[s.id];
            s.layoutParams = { rows: p.rows, cols: p.cols, padding: p.padding, gutter: p.gutter, linkRMode: p.linkRMode };
            s.layoutChildIds = p.childIds;
          } else if (layoutResult.childOf[s.id]) {
            s.layoutRole = 'child';
            s.layoutParentId = layoutResult.childOf[s.id];
          }
        }
        // 推导 currentLayout：选区里有 layout 父就激活（取第一个）
        const parentShape = selectedShapes.find((s) => s.layoutRole === 'parent');
        if (parentShape) {
          currentLayout = {
            parentId: parentShape.id,
            parentName: parentShape.name || '(未命名)',
            childIds: parentShape.layoutChildIds.slice(),
            params: { ...parentShape.layoutParams },
          };
        } else {
          currentLayout = null;
        }
      }
      renderUI();
      // monitor：选区非空就启动（实时读 adj + 更新状态卡；locked 的额外反算）
      // 选区空由 monitor 内部 `selectedShapes.length === 0` 自动 stop
      if (selectedShapes.length > 0) {
        startLockMonitor();
      } else {
        stopLockMonitor();
      }
    } catch (err) {
      setStatus('选区', '读失败：' + (err.message || err), 'status-warn');
      showToast('读选区失败: ' + (err.message || err));
    }
  }

  // ---------------- lock monitor：检测拖完松手后自动重应用 ----------------

  function startLockMonitor() {
    if (lockMonitor.timer) return;
    if (selectedShapes.length === 0) return;
    // 选区里有 locked shape 用 10ms 实时反算；只有未锁定的用 50ms 减负
    const interval = selectedShapes.some((s) => s.locked) ? LOCK_POLL_MS : IDLE_POLL_MS;
    lockMonitor.lastWidth = {};
    lockMonitor.lastHeight = {};
    lockMonitor.lastAdj = {};
    lockMonitor.stableCount = {};
    lockMonitor.lastCm = {};
    lockMonitor.lastSizeCm = {};  // v1.2.9
    lockMonitor.parentRDirty = false;
    if (lockMonitor.parentRSyncTimer) {
      clearTimeout(lockMonitor.parentRSyncTimer);
      lockMonitor.parentRSyncTimer = null;
    }
    lockMonitor.timer = setInterval(monitorTick, interval);
  }

  function stopLockMonitor() {
    if (lockMonitor.timer) {
      clearInterval(lockMonitor.timer);
      lockMonitor.timer = null;
    }
    lockMonitor.lastWidth = {};
    lockMonitor.lastHeight = {};
    lockMonitor.lastAdj = {};
    lockMonitor.stableCount = {};
    lockMonitor.lastCm = {};
    lockMonitor.lastSizeCm = {};  // v1.2.9
    lockMonitor.parentRDirty = false;
    if (lockMonitor.parentRSyncTimer) {
      clearTimeout(lockMonitor.parentRSyncTimer);
      lockMonitor.parentRSyncTimer = null;
    }
  }

  // 反算某个 shape 的 adj（被 onApply / 样式刷 / 重置时调用，让 monitor 同步状态）
  function syncLockMonitorForShape(shapeId, adj) {
    lockMonitor.lastAdj[shapeId] = adj;
    lockMonitor.stableCount[shapeId] = 0;
  }

  async function monitorTick() {
    if (selectedShapes.length === 0) {
      stopLockMonitor();
      return;
    }
    let needRefreshUI = false;  // 是否有 shape 的 currentCm 变了，需要重画 UI
    let recomputedIds = [];     // 拖尺寸被反算的
    let updatedLockIds = [];    // 拖 R 角滑块被"更新固定值"的
    let layoutParentChanges = [];  // v1.3.6：layout 父 R 角变化（→ 同步子）
    try {
      await PowerPoint.run(async (ctx) => {
        // v1.2.2 driver + radius-core：lock monitor 走新分层
        const driver = window.PptDriver.createDriver(ctx);
        // v1.3.7 修 #2 真 bug：monitorTick 改回 v1.0 模式（per-shape get(0) + per-shape sync）
        // 原因：v1.3.6 改成 collection-level load 'items/adjustments/items/value' + 1 sync + new get(0)，
        //       但 Mac LTSC proxy 是 snapshot 风格，sync 后 new get(0) 返回新 proxy 没 value。
        //       用户实测 v1.2.4：拖父 R 角时 4 个子没联动，根因就是 monitorTick 读不到 currentAdj → continue → ss.currentCm 不更新。
        //       refreshSelection 用的就是 v1.0 模式（1210-1213 行），实测 work。monitorTick 改回同样模式。
        // bug #3（per-shape sync 累积）→ catch 兜底（实测 v1.2 时期只偶发，无影响）
        const sel = driver.selectedShapes();
        // 不 load 'items/adjustments/items/value'（v1.3.6 那个 load 没用）—— 改用 v1.0 per-shape 模式
        driver.load(sel, 'items/id, items/width, items/height, items/adjustments, items/tags');
        await driver.sync();
        // readTagsBulk 一次拿全部 tag（避开 readTag per-call sync 累积，bug #2 子 bug 修法保留）
        const tagsById = driver.readTagsBulk(sel.items);
        for (const sh of sel.items) {
          const shId = driver.shapeId(sh);
          try {
            if (!driver.isRoundRect(sh)) continue; // 不是 roundRect
            // v1.0 模式：先 get(0) 存变量 → per-shape sync → 读 value
            const adjResult = sh.adjustments.get(0);
            await driver.sync();
            let currentAdj = null;
            try { currentAdj = adjResult.value; } catch (_) { continue; }
            if (currentAdj == null) continue;
            const size = driver.size(sh);
            const minSideCm = Math.min(size.width, size.height) / PT_PER_CM;
            if (minSideCm <= 0) continue;
            const currentCm = currentAdj * minSideCm;

            // 找对应 selectedShape
            const ss = selectedShapes.find((x) => x.id === shId);
            if (!ss) continue;
            // 1) 所有 roundRect 都更新内存 currentCm（实时显示 R 角）
            //    任何变化都标 dirty（浮点抖动 < 0.01 cm 的显示精度，看不出来）
            const oldCm = ss.currentCm;
            ss.currentCm = currentCm;
            ss.width = size.width;
            ss.height = size.height;
            if (oldCm == null || currentCm !== oldCm) {
              needRefreshUI = true;
            }
            // 2) 只对 locked shape 做反算/更新固定值
            if (!ss.locked) continue;

            const targetCm = Math.min(ss.lockedCm, minSideCm / 2);
            const targetAdj = (targetCm / minSideCm) * ADJ_SCALE;
            const lastW = lockMonitor.lastWidth[shId];
            const lastH = lockMonitor.lastHeight[shId];
            const lastA = lockMonitor.lastAdj[shId];
            // 第一轮（lastA = null）：只记录初始状态，不做反算
            // （否则会触发"idle 兜底"，把用户第一次拖 R 角的值当异常反算回去）
            if (lastA == null) {
              lockMonitor.lastWidth[shId] = size.width;
              lockMonitor.lastHeight[shId] = size.height;
              lockMonitor.lastAdj[shId] = currentAdj;
              lockMonitor.stableCount[shId] = 0;
              continue;
            }
            const wChanged = Math.abs(size.width - lastW) > SIZE_EPSILON;
            const hChanged = Math.abs(size.height - lastH) > SIZE_EPSILON;
            const aChanged = Math.abs(currentAdj - lastA) > ADJ_EPSILON;
            const sizeChanged = wChanged || hChanged;  // 任意一边变了都算"调尺寸"

            if (sizeChanged) {
              // 拖尺寸手柄（任意边 / 角）：立刻反算回固定值
              if (Math.abs(currentAdj - targetAdj) > ADJ_EPSILON) {
                driver.setAdjFraction(sh, targetAdj);
                recomputedIds.push(shId);
              }
              lockMonitor.lastAdj[shId] = targetAdj;
              lockMonitor.stableCount[shId] = 0;
            } else if (aChanged) {
              // 拖 R 角黄色滑块：等稳定后视作主动改值
              lockMonitor.stableCount[shId] = (lockMonitor.stableCount[shId] || 0) + 1;
              if (lockMonitor.stableCount[shId] >= LOCK_STABLE_THRESHOLD) {
                if (ss.strictLocked) {
                  // 防误触：反算回去
                  driver.setAdjFraction(sh, targetAdj);
                  lockMonitor.lastAdj[shId] = targetAdj;
                  recomputedIds.push(shId);
                } else {
                  // 仅使用数值固定 R 角：把当前 adj 提升为新的固定值
                  const newCm = currentAdj * minSideCm;
                  await window.RadiusCore.writeLockState(driver, sh, { lockedCm: newCm });
                  ss.lockedCm = newCm;
                  lockMonitor.lastAdj[shId] = currentAdj;
                  updatedLockIds.push(shId);
                }
                lockMonitor.stableCount[shId] = 0;
              }
            } else {
              // 都没变：idle，检查兜底（adj 跟 target 不一致但 size 和 adj 都没"主动变化"）
              if (Math.abs(currentAdj - targetAdj) > ADJ_EPSILON) {
                // 极端 race 兜底：写回
                driver.setAdjFraction(sh, targetAdj);
                lockMonitor.lastAdj[shId] = targetAdj;
                recomputedIds.push(shId);
              }
              lockMonitor.stableCount[shId] = 0;
            }
            lockMonitor.lastWidth[shId] = size.width;
            lockMonitor.lastHeight[shId] = size.height;
          } catch (eShape) {
            // 单个 shape 出错（防御性兜底）—— 不影响其他 shape，也不影响整个 tick
            // 跳过这个 shape，下个 tick 再看
            continue;
          }
        }
        await driver.sync();
      });
      // 轻量更新 UI（只改文本节点，不重建 DOM）
      if (needRefreshUI) renderCurrentRadius();
      if (recomputedIds.length > 0) {
        showToast(`🔒 使用数值固定 R 角：反算了 ${recomputedIds.length} 个被改的 R 角`);
      } else if (updatedLockIds.length > 0) {
        showToast(`🪄 使用数值固定 R 角已跟随 R 角滑块更新（${updatedLockIds.length} 个）`);
      }

      // v1.3.6 修 #2：layout 父 R 角变化 → 同步子 R 角
      // v1.2.9：同时检测父 size 变化（width/height）→ 同步子位置/尺寸 + R 角
      // 场景：用户在 PPT 里直接拖父的 R 角黄色滑块 / 拖父的边缘 / 角控制柄（不走 task pane 的 onApply），
      //       monitorTick 检测到 currentCm 或 width/height 变化，detectLayoutParentChanges / detectLayoutParentSizeChanges 返回变化的父，
      //       调 syncLayoutChildrenRIfNeeded → applyLayoutToChildren → 重算 layout 几何 + R 角
      // 首次见到（lastCm / lastSizeCm = null）只记录不触发 sync（避免启动时无意义重写子 R 角）
      try {
        let hasRealChange = false;
        // 1) R 角变化
        const rChanges = window.RadiusCore.detectLayoutParentChanges(lockMonitor.lastCm, selectedShapes);
        for (const c of rChanges) {
          if (c.lastCm == null) {
            // 首次见到：只记 lastCm，不触发 sync
            lockMonitor.lastCm[c.parentId] = c.newCm;
            console.log(`[layout-link] first-see parentId=${c.parentId} newCm=${c.newCm.toFixed(3)} (记 lastCm 不 fire)`);
            continue;
          }
          // 真正变了：更新 lastCm + 标 dirty
          lockMonitor.lastCm[c.parentId] = c.newCm;
          hasRealChange = true;
          console.log(`[layout-link] R-fire parentId=${c.parentId} lastCm=${c.lastCm.toFixed(3)} newCm=${c.newCm.toFixed(3)} (→ 200ms 后调 syncLayoutChildrenR)`);
        }
        // 2) v1.2.9：size 变化（widthCm / heightCm）
        const sChanges = window.RadiusCore.detectLayoutParentSizeChanges(lockMonitor.lastSizeCm, selectedShapes);
        for (const c of sChanges) {
          if (c.lastSize == null) {
            // 首次见到：只记 lastSizeCm，不触发 sync
            lockMonitor.lastSizeCm[c.parentId] = c.newSize;
            console.log(`[layout-link] first-see parentId=${c.parentId} newSize=${JSON.stringify({ w: c.newSize.widthCm.toFixed(2), h: c.newSize.heightCm.toFixed(2) })} (记 lastSize 不 fire)`);
            continue;
          }
          // 真正变了：更新 lastSize + 标 dirty
          lockMonitor.lastSizeCm[c.parentId] = c.newSize;
          hasRealChange = true;
          console.log(`[layout-link] SIZE-fire parentId=${c.parentId} lastSize=${JSON.stringify({ w: c.lastSize.widthCm.toFixed(2), h: c.lastSize.heightCm.toFixed(2) })} newSize=${JSON.stringify({ w: c.newSize.widthCm.toFixed(2), h: c.newSize.heightCm.toFixed(2) })} (→ 200ms 后调 syncLayoutChildrenR)`);
        }
        if (hasRealChange) {
          lockMonitor.parentRDirty = true;
          scheduleParentRSync();
        }
      } catch (e) {
        // 不影响主流程，silent skip
      }
    } catch (e) {
      // 整个 tick 失败：silent skip（不弹窗，不影响 PPT）
    }
  }

  // v1.3.6 修 #6：节流调 syncLayoutChildrenRIfNeeded（避免 10ms tick 频繁开新 PowerPoint.run）
  // 200ms 窗口：用户在拖父 R 角滑块时（≈ 60fps）一次拖完最多触发 5 次，但节流后实际只跑 1 次
  function scheduleParentRSync() {
    if (lockMonitor.parentRSyncTimer) return;  // 已经有 pending 的 timer
    lockMonitor.parentRSyncTimer = setTimeout(async () => {
      lockMonitor.parentRSyncTimer = null;
      if (!lockMonitor.parentRDirty) return;
      lockMonitor.parentRDirty = false;
      try {
        const before = selectedShapes.filter((s) => s.layoutRole === 'parent' && s.layoutParams && s.layoutChildIds).length;
        const r = await syncLayoutChildrenRIfNeeded();
        if (before > 0) console.log(`[layout-link] syncLayoutChildrenRIfNeeded done: parents=${before} result=${JSON.stringify(r)}`);
      } catch (e) {
        console.log('[layout-link] syncLayoutChildrenRIfNeeded error:', e && e.message ? e.message : e);
      }
    }, PARENT_R_SYNC_DEBOUNCE_MS);
  }

  // ---------------- UI helpers ----------------

  function setStatus(label, text, cardClass) {
    $('status-text').textContent = text;
    $('status-card').className = 'status-card ' + cardClass;
    const labelEl = document.querySelector('.status-row .status-label');
    if (labelEl) labelEl.textContent = label;
  }

  function showToast(msg) {
    const el = $('toast');
    if (!el) return;
    el.textContent = msg;
    el.className = 'toast toast-show';
    // 动态调位置：debug-log 打开时浮到 220px 避开，关闭时贴底 60px
    // （避免 toast 被 fixed 底部的 debug-log bar 挡住，v1.2.4 反馈）
    const debugLog = $('debug-log');
    el.style.bottom = (debugLog && debugLog.open) ? '220px' : '60px';
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => { el.className = 'toast'; }, 2200);
  }

  function renderUI() {
    // 状态卡
    if (selectedShapes.length === 0) {
      setStatus('选区', '未选中', 'status-warn');
      $('current-radius').textContent = '—';
      $('locked-count').textContent = '—';
    } else {
      const allRound = selectedShapes.every((s) => s.isRoundRect);
      const anyRound = selectedShapes.some((s) => s.isRoundRect);
      const mixed = !allRound && anyRound;
      const ok = allRound;
      setStatus('选区', `${selectedShapes.length} 个${mixed ? '（混合）' : ''}`, ok ? 'status-ok' : 'status-warn');
      // 当前 R 角（圆角矩形的）
      const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
      if (roundShapes.length > 0 && roundShapes[0].currentCm != null) {
        $('current-radius').textContent = `${roundShapes[0].currentCm.toFixed(2)} 厘米`;
      } else {
        $('current-radius').textContent = '—';
      }
      // 锁定数
      const lockedCount = selectedShapes.filter((s) => s.locked).length;
      $('locked-count').textContent = lockedCount > 0 ? `${lockedCount}` : '—';
    }
    // 形状列表
    renderShapeList();
    // 输入框：单位标签 + 输入限制 + apply 按钮可用性
    $('unit-label').textContent = currentUnit === 'cm' ? '厘米' : '百分比';
    const hasRound = selectedShapes.length > 0 && selectedShapes.every((s) => s.isRoundRect);
    const inputVal = parseFloat($('radius-input').value);
    $('apply-btn').disabled = !(hasRound && Number.isFinite(inputVal) && inputVal >= 0);
    $('reapply-btn').disabled = !selectedShapes.some((s) => s.locked);
    // 锁定按钮
    updateLockButton();
    // v1.2: 布局面板
    renderLayoutPanel();
  }

  function renderShapeList() {
    const list = $('shape-list');
    if (!list) return;
    if (selectedShapes.length === 0) {
      list.innerHTML = '<div class="empty-list">在 PPT 里框选形状后会出现在这里</div>';
      return;
    }
    list.innerHTML = '';
    for (const s of selectedShapes) {
      const row = document.createElement('div');
      row.className = 'shape-row' + (s.isRoundRect ? '' : ' shape-row-warn');
      row.dataset.shapeId = s.id; // 给 row 加 shapeId 标识，monitor 可以轻量更新 .shape-r 文本
      let tag = '';
      if (s.locked && s.strictLocked) {
        tag = `<span class="shape-lock shape-lock-strict">🔒 ${s.lockedCm.toFixed(2)}cm 防误触</span>`;
      } else if (s.locked) {
        tag = `<span class="shape-lock">🔒 ${s.lockedCm.toFixed(2)}cm</span>`;
      } else if (!s.isRoundRect) {
        tag = '<span class="shape-warn">非圆角矩形</span>';
      }
      const rText = s.currentCm != null ? `${s.currentCm.toFixed(2)}cm` : '—';
      row.innerHTML = `<span class="shape-name">${s.name || '(未命名)'}</span><span class="shape-r">${rText}</span>${tag}`;
      list.appendChild(row);
    }
  }

  // 轻量更新"当前 R 角"显示：状态卡 #current-radius + 形状列表每行 .shape-r
  // 不会重建 DOM，只改文本节点
  function renderCurrentRadius() {
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    if (roundShapes.length > 0 && roundShapes[0].currentCm != null) {
      $('current-radius').textContent = `${roundShapes[0].currentCm.toFixed(2)} 厘米`;
    } else {
      $('current-radius').textContent = '—';
    }
    // 每个 shape 行的 R 角文本
    for (const s of selectedShapes) {
      const row = document.querySelector(`.shape-row[data-shape-id="${s.id}"]`);
      if (!row) continue;
      const rSpan = row.querySelector('.shape-r');
      if (rSpan) {
        rSpan.textContent = s.currentCm != null ? `${s.currentCm.toFixed(2)}cm` : '—';
      }
    }
  }

  function updateLockButton() {
    const btn = $('lock-btn');
    if (!btn) return;
    if (selectedShapes.length === 0) {
      btn.disabled = true;
      $('lock-icon').textContent = '🔒';
      $('lock-label').textContent = '使用数值固定 R 角';
      $('lock-hint').textContent = '读选中…';
      updateStrictToggle();
      return;
    }
    btn.disabled = false;
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    const allLocked = roundShapes.length > 0 && roundShapes.every((s) => s.locked);
    $('lock-icon').textContent = allLocked ? '🔒' : '🔒';
    $('lock-label').textContent = allLocked ? '关闭使用数值固定 R 角' : '使用数值固定 R 角';
    $('lock-hint').textContent = allLocked
      ? `已使用数值固定 R 角 ${roundShapes.length} 个（PPT 内编辑会被反算回固定值）`
      : `开启后 R 角按厘米值保持，PPT 内编辑会被反算`;
    updateStrictToggle();
  }

  /** 根据当前 selectedShapes 状态更新防误触开关：disabled / 状态 / 文案
   *  防误触现在跟"使用数值固定 R 角"互相独立：开启防误触时如果还没 lock，
   *  会自动用当前 R 角作 fixed value（见 onToggleStrict）。
   *  所以 toggle 的可用性只跟"是否选了 roundRect"挂钩，不再需要先 lock。 */
  function updateStrictToggle() {
    const label = $('strict-toggle');
    const cb = $('strict-checkbox');
    const hintEl = label ? label.querySelector('.strict-hint') : null;
    if (!label || !cb) return;
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    if (selectedShapes.length === 0 || roundShapes.length === 0) {
      // 没选 / 不是 roundRect → toggle 不可用
      label.classList.add('disabled');
      cb.disabled = true;
      cb.checked = false;
      if (hintEl) hintEl.textContent = '开启后任何修改都不改 R 角';
    } else {
      // 任何时候都能开 strict（开启时自动 lock）
      label.classList.remove('disabled');
      cb.disabled = false;
      const allStrict = roundShapes.every((s) => s.strictLocked);
      cb.checked = allStrict;
      if (hintEl) hintEl.textContent = allStrict
        ? '已开启（任何修改都不会改 R 角）'
        : '开启后任何修改都不改 R 角';
    }
  }

  // ---------------- 操作 ----------------

  /** 应用 R 角：所有选中的圆角矩形都改成输入的 cm 值 */
  async function onApply() {
    if (selectedShapes.length === 0) {
      showToast('请先在 PPT 里框选圆角矩形');
      return;
    }
    const raw = parseFloat($('radius-input').value);
    if (!Number.isFinite(raw) || raw < 0) {
      showToast('请输入有效的 R 角值');
      return;
    }
    // v1.1 防误触拦截：选区里有任何 strict 锁定 → 全部拒绝
    const strictLocked = selectedShapes.filter((s) => s.isRoundRect && s.strictLocked);
    if (strictLocked.length > 0) {
      showToast(`🔒 防误触已开启（${strictLocked.length} 个），不能改 R 角。先关掉防误触或解锁。`);
      return;
    }
    // 输入值按当前单位换算成 cm
    const cm = valueToCm(raw, currentUnit);
    let updated = 0;
    let failed = 0;
    let lockedSynced = 0; // 计数：locked 子被同步 fixed value 的数量
    // 写之前停 monitor（避免 race）
    stopLockMonitor();
    try {
      await PowerPoint.run(async (ctx) => {
        // === v1.2.2 driver + radius-core 集成：onApply 走新分层 ===
        const driver = window.PptDriver.createDriver(ctx);
        const sel = driver.selectedShapes();
        driver.load(sel, 'items/id, items/width, items/height, items/adjustments, items/tags');
        await driver.sync();
        for (const sh of sel.items) {
          // 走新分层：业务逻辑在 radius-core.writeRadius，driver 只负责 PPT 读写
          const r = await window.RadiusCore.writeRadius(driver, sh, cm, {});
          if (!r.ok) {
            failed++;
            continue;
          }
          updated++;
          if (r.wasLocked) lockedSynced++;
        }
        await driver.sync();
      });
      if (failed === 0) {
        const displayVal = currentUnit === '%'
          ? `${raw.toFixed(1)}%`
          : `${raw.toFixed(2)} 厘米`;
        const lockHint = lockedSynced > 0
          ? `，${lockedSynced} 个使用数值固定 R 角已同步更新`
          : '';
        showToast(`✅ 已更新 ${updated} 个圆角矩形为 ${displayVal}${lockHint}`);
        if (updated > 0) {
          // 写到内存 + 渲染
          const newHistory = pushHistory(raw, currentUnit);
          renderHistory(newHistory);
        }
      } else {
        showToast(`⚠️ ${updated} 个成功，${failed} 个失败（可能不是圆角矩形）`);
      }
      await refreshSelection();
      // v1.2: 选区里有 layout 父 → 同步子 R 角（联动）
      await syncLayoutChildrenRIfNeeded();
    } catch (err) {
      showToast('应用失败：' + (err.message || err));
    } finally {
      // 写完恢复 monitor（stopLockMonitor 已清空 last 状态，startLockMonitor 从干净开始）
      if (selectedShapes.length > 0) startLockMonitor();
    }
  }

  /** 使用数值固定 R 角 开启/关闭：用 shape.tags 存固定值，跟 .pptx 文件走 */
  async function onToggleLock() {
    if (selectedShapes.length === 0) {
      showToast('请先在 PPT 里框选圆角矩形');
      return;
    }
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    if (roundShapes.length === 0) {
      showToast('选中的形状都不是圆角矩形');
      return;
    }
    const allLocked = roundShapes.every((s) => s.locked);
    const inputVal = parseFloat($('radius-input').value);
    const locks = {};
    const strict = {}; // 关闭时清空所有 strict 标记；开启时保留之前 strict 状态
    let touched = 0;
    for (const s of roundShapes) {
      if (allLocked) {
        // 关闭使用数值固定 R 角：显式 null/false，让 writeLockState 走 delete 路径
        // （v1.3.5 修：原版是空 map → saveLocksViaTags 走 no-op → tag 删不掉）
        locks[s.id] = null;
        strict[s.id] = false;
      } else {
        // 开启使用数值固定 R 角：优先用输入框值，否则用当前 R 角
        const inputCm = Number.isFinite(inputVal) && inputVal > 0
          ? valueToCm(inputVal, currentUnit)
          : s.currentCm;
        if (inputCm > 0) locks[s.id] = inputCm;
        // 之前已经开启过 strict（防误触），再次"开启使用数值固定 R 角"时保留 strict 状态
        if (s.strictLocked) strict[s.id] = true;
      }
      touched++;
    }
    const r = await saveLocksViaTags(locks, strict);
    if (!r.ok) {
      showToast('操作失败：' + (r.error?.message || r.error));
      return;
    }
    showToast(allLocked
      ? `已关闭使用数值固定 R 角（${touched} 个）`
      : `已开启使用数值固定 R 角（${touched} 个）— PPT 内编辑会被反算回固定值`);
    await refreshSelection();
  }

  /** 切换「防误触」开关：把所有选中的 roundRect 的 strict 状态切换为 newValue
   *  防误触现在跟"使用数值固定 R 角"互相独立：
   *  - 开启：自动用当前 R 角作 fixed value（如果还没 lock），写 lock + strict 两个 tag
   *  - 关闭：只删 strict tag（保留 lock tag，user 可以选择保留 fixed value）
   *  - 关闭"使用数值固定 R 角"时会同时清掉 strict（见 onToggleLock，因为反算目标没了） */
  async function onToggleStrict(newValue) {
    if (selectedShapes.length === 0) {
      showToast('请先在 PPT 里框选圆角矩形');
      return;
    }
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    if (roundShapes.length === 0) {
      showToast('选中的形状都不是圆角矩形');
      return;
    }
    if (newValue) {
      // 开启：自动用当前 R 角作 fixed value（如果还没 lock）
      const locks = {}; // id -> cm
      for (const s of roundShapes) {
        if (s.locked && s.lockedCm > 0) {
          // 已 lock：保留原 fixed value
          locks[s.id] = s.lockedCm;
        } else if (s.currentCm != null && s.currentCm > 0) {
          // 没 lock：用当前 R 角作 fixed value
          locks[s.id] = s.currentCm;
        } else {
          // R 角 = 0 / 未知：没法设 fixed value
          showToast(`无法开启防误触：${s.name || '(未命名)'} 当前 R 角未知或为 0`);
          return;
        }
      }
      const strict = {};
      for (const s of roundShapes) strict[s.id] = true;
      const r = await saveLocksViaTags(locks, strict);
      if (!r.ok) {
        showToast('操作失败：' + (r.error?.message || r.error));
        return;
      }
      showToast(`🔒 防误触已开启（${roundShapes.length} 个）— 已自动用当前 R 角作固定值`);
    } else {
      // 关闭：只删 strict tag（不动 lock tag）
      for (const s of roundShapes) {
        await updateLockTagForShape(s.id, undefined, false);
      }
      showToast(`防误触已关闭（${roundShapes.length} 个）— 允许主动调整 R 角${roundShapes.some((s) => s.locked) ? '，固定值保留' : ''}`);
    }
    await refreshSelection();
  }

  /** 重新应用锁定：按当前形状大小反算 adj */
  async function onReapply() {
    const locked = selectedShapes.filter((s) => s.locked);
    if (locked.length === 0) {
      showToast('当前选区没有锁定的圆角矩形');
      return;
    }
    let applied = 0;
    let failed = 0;
    try {
      await PowerPoint.run(async (ctx) => {
        // v1.2.2 driver + radius-core：走 reapplyLock（自动处理 clamp）
        const driver = window.PptDriver.createDriver(ctx);
        const sel = driver.selectedShapes();
        driver.load(sel, 'items/id, items/width, items/height, items/adjustments');
        await driver.sync();
        for (const sh of sel.items) {
          const id = driver.shapeId(sh);
          const target = locked.find((x) => x.id === id);
          if (!target) continue;
          const r = await window.RadiusCore.reapplyLock(driver, sh, target.lockedCm);
          if (r.ok) applied++;
          else failed++;
        }
        await driver.sync();
      });
      showToast(`🔒 重新应用了 ${applied} 个锁定的 R 角${failed > 0 ? `，${failed} 个失败` : ''}`);
      await refreshSelection();
    } catch (err) {
      showToast('操作失败：' + (err.message || err));
    }
  }

  function onUnitChange(newUnit) {
    if (newUnit === currentUnit) return;
    // 把当前输入框值换算到新单位
    const oldVal = parseFloat($('radius-input').value);
    if (Number.isFinite(oldVal) && oldVal >= 0) {
      const cm = valueToCm(oldVal, currentUnit);
      const newVal = cmToValue(cm, newUnit);
      $('radius-input').value = newUnit === '%' ? newVal.toFixed(1) : newVal.toFixed(2);
    }
    currentUnit = newUnit;
    // 更新按钮 active 态
    document.querySelectorAll('.unit-btn').forEach((btn) => {
      const active = btn.dataset.unit === newUnit;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });
    // 更新 step / placeholder
    if (newUnit === '%') {
      $('radius-input').step = '0.1';
      $('radius-input').min = '0';
      $('radius-input').max = '50';
      $('radius-input').placeholder = '10';
    } else {
      $('radius-input').step = '0.01';
      $('radius-input').min = '0';
      $('radius-input').removeAttribute('max');
      $('radius-input').placeholder = '0.30';
    }
    $('unit-label').textContent = newUnit === 'cm' ? '厘米' : '百分比';
    renderUI();
  }

  // ---------------- v1.1 新增：预设库（纯内存，session 内） ----------------

  const MAX_PRESETS = 5;

  // userPresets = [{ id, name, value, unit }]
  // unit 跟随吸取/添加时的 currentUnit；应用时按这个单位换算到 cm
  let userPresets = [];

  function nextPresetId() {
    return 'p_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 7);
  }

  function formatPresetValue(value, unit) {
    if (unit === '%') {
      return `${Number.isInteger(value) ? value : value.toFixed(1)}%`;
    }
    return `${value.toFixed(2)}cm`;
  }

  function renderPresets(presets) {
    const list = $('preset-list');
    if (!list) return;
    list.innerHTML = '';
    const arr = Array.isArray(presets) ? presets : [];
    if (arr.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'preset-empty';
      empty.textContent = '保存常用 R 角，一键应用。点击「+ 保存当前值」开始。';
      list.appendChild(empty);
      return;
    }
    for (const p of arr) {
      const row = document.createElement('div');
      row.className = 'preset-row';
      row.dataset.id = p.id;

      // 名称（可编辑）
      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.className = 'preset-name';
      nameInput.value = p.name;
      nameInput.maxLength = 16;
      nameInput.title = '点击重命名，回车保存';
      nameInput.addEventListener('change', () => {
        const v = nameInput.value.trim();
        if (v) {
          p.name = v;
          showToast(`已重命名为「${v}」`);
        } else {
          nameInput.value = p.name;
        }
      });
      nameInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') nameInput.blur();
      });

      // 数值（可编辑，按 p.unit 解读）
      const valInput = document.createElement('input');
      valInput.type = 'number';
      valInput.className = 'preset-value-input';
      valInput.dataset.unit = p.unit;
      if (p.unit === '%') {
        valInput.step = '0.1';
        valInput.min = '0';
        valInput.max = '50';
      } else {
        valInput.step = '0.01';
        valInput.min = '0';
        valInput.removeAttribute('max');
      }
      valInput.value = p.unit === '%'
        ? (Number.isInteger(p.value) ? String(p.value) : p.value.toFixed(1))
        : p.value.toFixed(2);
      valInput.title = `点击编辑数值（按 ${p.unit === '%' ? '百分比' : '厘米'} 解读），回车保存`;
      valInput.addEventListener('change', () => {
        const v = parseFloat(valInput.value);
        if (!Number.isFinite(v) || v < 0) {
          // 还原成原值
          valInput.value = p.unit === '%'
            ? (Number.isInteger(p.value) ? String(p.value) : p.value.toFixed(1))
            : p.value.toFixed(2);
          showToast('请输入有效的 R 角值');
          return;
        }
        if (p.unit === '%' && v > 50) {
          valInput.value = String(p.value);
          showToast('百分比模式范围 0~50%');
          return;
        }
        p.value = v;
        // 更新应用按钮的 title 提示
        applyBtn.title = `应用 ${p.name} = ${formatPresetValue(p.value, p.unit)} 到当前选区`;
        showToast(`已更新「${p.name}」为 ${formatPresetValue(v, p.unit)}`);
      });
      valInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') valInput.blur();
      });

      // 应用按钮
      const applyBtn = document.createElement('button');
      applyBtn.type = 'button';
      applyBtn.className = 'preset-apply';
      applyBtn.textContent = '应用';
      applyBtn.title = `应用 ${p.name} = ${formatPresetValue(p.value, p.unit)} 到当前选区`;
      applyBtn.addEventListener('click', () => applyPreset(p));

      // 删除
      const delBtn = document.createElement('button');
      delBtn.type = 'button';
      delBtn.className = 'preset-del';
      delBtn.textContent = '×';
      delBtn.title = '删除此预设';
      delBtn.addEventListener('click', () => deletePreset(p.id));

      row.appendChild(nameInput);
      row.appendChild(valInput);
      row.appendChild(applyBtn);
      row.appendChild(delBtn);
      list.appendChild(row);
    }
  }

  function addPresetFromInput() {
    if (userPresets.length >= MAX_PRESETS) {
      showToast(`预设库已满（${MAX_PRESETS} 个），请先删一个再加`);
      return;
    }
    // 优先用当前选中的圆角矩形的 R 角；没有就退回输入框
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect && s.currentCm != null && s.currentCm > 0);
    let raw, unit;
    if (roundShapes.length > 0) {
      const src = roundShapes[0];
      unit = currentUnit;
      raw = cmToValue(src.currentCm, unit);
    } else {
      raw = parseFloat($('radius-input').value);
      unit = currentUnit;
      if (!Number.isFinite(raw) || raw < 0) {
        showToast('请先在 PPT 里选 1 个圆角矩形，或在输入框里输入有效的 R 角值');
        $('radius-input').focus();
        return;
      }
    }
    // 单位校验
    if (unit === '%' && (raw < 0 || raw > 50)) {
      showToast('百分比模式范围 0~50%');
      return;
    }
    // 默认命名：「预设 N」
    const name = `预设 ${userPresets.length + 1}`;
    const p = { id: nextPresetId(), name, value: raw, unit };
    userPresets = [p, ...userPresets].slice(0, MAX_PRESETS);
    renderPresets(userPresets);
    showToast(`已保存为「${name}」(${formatPresetValue(raw, unit)})`);
  }

  function applyPreset(preset) {
    if (selectedShapes.length === 0) {
      showToast('请先在 PPT 里框选圆角矩形');
      return;
    }
    const roundCount = selectedShapes.filter((s) => s.isRoundRect).length;
    if (roundCount === 0) {
      showToast('选中的形状都不是圆角矩形');
      return;
    }
    // 切到预设的单位 + 把值写到输入框 + 触发 onApply
    if (preset.unit !== currentUnit) {
      onUnitChange(preset.unit);
    }
    $('radius-input').value = preset.unit === '%'
      ? (Number.isInteger(preset.value) ? preset.value : preset.value.toFixed(1))
      : preset.value.toFixed(2);
    onApply();
  }

  function deletePreset(id) {
    const before = userPresets.length;
    userPresets = userPresets.filter((p) => p.id !== id);
    if (userPresets.length < before) {
      renderPresets(userPresets);
      showToast('预设已删除');
    }
  }

  // ---------------- v1.1 新增：R 角样式刷（idle / sourcing / brushing 状态机） ----------------

  // pipetteSource: { value, unit, sourceShapeName } | null
  // 存吸取时的具体数值 + 单位；应用时按这个单位换算到 cm
  let pipetteState = 'idle';
  let pipetteSource = null; // { value, unit, sourceShapeName, cm, sourceStrict }
  let pipetteSyncStrict = false; // checkbox：是否同时同步源形状的「防误触」状态

  function setPipetteState(newState) {
    pipetteState = newState;
    const btn = $('pipette-btn');
    const badge = $('pipette-state-badge');
    const hint = $('pipette-hint');
    const icon = $('pipette-icon');
    const label = $('pipette-label');
    btn.dataset.state = newState;
    if (newState === 'idle') {
      badge.className = 'pipette-state-badge idle';
      badge.textContent = '空闲';
      hint.classList.remove('has-source');
      hint.textContent = '点击吸取一个圆角矩形的 R 角，再点其他形状应用';
      icon.textContent = '🖌️';
      label.textContent = '吸取 R 角';
      btn.classList.remove('state-sourcing', 'state-brushing');
    } else if (newState === 'sourcing') {
      badge.className = 'pipette-state-badge sourcing';
      badge.textContent = '吸取中…';
      hint.classList.add('has-source');
      hint.textContent = '在 PPT 里点选 1 个圆角矩形吸取其 R 角';
      icon.textContent = '🎯';
      label.textContent = '取消吸取';
      btn.classList.add('state-sourcing');
      btn.classList.remove('state-brushing');
    } else if (newState === 'brushing') {
      badge.className = 'pipette-state-badge brushing';
      badge.textContent = '刷取中…';
      hint.classList.add('has-source');
      if (pipetteSource) {
        const syncTag = pipetteSyncStrict ? '（含防误触同步）' : '';
        hint.textContent = `源：${pipetteSource.sourceShapeName} · ${formatPresetValue(pipetteSource.value, pipetteSource.unit)}${syncTag} · 选中目标形状自动应用`;
      } else {
        hint.textContent = '选中目标形状自动应用';
      }
      icon.textContent = '🪣';
      label.textContent = '退出刷取';
      btn.classList.add('state-brushing');
      btn.classList.remove('state-sourcing');
    }
  }

  async function onPipetteButtonClick() {
    if (pipetteState === 'idle') {
      // 先 refresh 一下，让 selectedShapes 跟当前选区一致（避免用旧内存）
      await refreshSelection();
      const src = selectedShapes.find((s) => s.isRoundRect);
      if (src) {
        // 当前已经有选中的圆角矩形 → 直接以它为 source
        // （R 角 = 0 也允许吸，apply 时会把目标变成直角矩形）
        setPipetteState('sourcing');
        await pickupFromSelection();
        return;
      }
      if (selectedShapes.length > 0) {
        showToast('当前选中的不是圆角矩形，请先在 PPT 里点 1 个圆角矩形');
      }
      setPipetteState('sourcing');
      showToast('🎯 进入吸取模式 — 在 PPT 里点 1 个圆角矩形');
    } else {
      // 任意非 idle 状态点击按钮都退出
      setPipetteState('idle');
      pipetteSource = null;
      showToast('样式刷已关闭');
    }
  }

  // 从选区第一个 roundRect 吸取（自己读选区，不依赖 selectedShapes 内存）
  // Mac LTSC task pane 必加：get(0) 存到变量 → sync → 读 value
  // （不能 load 之后再新调 get(0).value，那时 value 还没填上，会报"尚未加载"）
  // v1.3.6 迁移：PowerPoint.run 部分走 radius-core.pickupFromSelection
  async function pickupFromSelection() {
    let picked = null;
    try {
      await PowerPoint.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/name, items/width, items/height, items/adjustments, items/tags');
        await ctx.sync();
        const r = await window.RadiusCore.pickupFromSelection(driver, sel);
        if (r) {
          picked = r;
        }
      });
    } catch (e) {
      showToast('吸取失败：' + (e.message || e));
      return;
    }
    if (!picked) {
      showToast('选区里没有圆角矩形');
      return;
    }
    // 把 cm 换算到当前 currentUnit（更直观）
    const value = cmToValue(picked.cm, currentUnit);
    pipetteSource = {
      value,
      unit: currentUnit,
      sourceShapeName: picked.name,
      cm: picked.cm, // 内部统一存 cm
      sourceStrict: picked.sourceStrict,  // 源形状的防误触状态（供「刷防误触状态」选项使用）
      sourceId: picked.id,
    };
    setPipetteState('brushing');
    const strictHint = pipetteSyncStrict && picked.sourceStrict ? '（含防误触）' : '';
    showToast(`🪣 已吸取「${pipetteSource.sourceShapeName}」= ${formatPresetValue(value, currentUnit)} — 选中目标形状自动应用${strictHint}`);
    // 顺便把 selectedShapes 内存刷新一下（让状态卡同步显示源形状）
    refreshSelection();
  }

  // 把 pipetteSource 应用到选区里所有 roundRect
  //
  // v1.3.6 迁移：所有 3 步都走 radius-core.applyPickedToSelection（单 PowerPoint.run 完成）：
  //   步骤 0 — 拦截：选区里有任何目标启用了防误触 → 整个样式刷拒绝（不论是否勾选「刷防误触状态」）
  //   步骤 1 — 第一次"应用"：刷 R 角到所有目标 + 同步 fixed value（已 lock 的目标）
  //   步骤 2 — 第二次"应用"：根据【刷防误触状态】勾选决定是否把源 strict 状态写到目标
  //
  // 关键：radius-core.applyPickedToSelection 内部用 driver.readTagsBulk 一次拿全部 tag，
  //       避开 per-call readTag + sync 在 for 循环内累积（v1.2.6 + v1.3.6 Mac LTSC 坑，
  //       之前 dialog.js 直接用 writeRadiusToShape 时 4 个子只写 2 个就是这个原因）。
  // 修 #1：吸取后无法刷入任何形状 → 改走 radius-core（稳定写 R 角 + 同步 lock + 刷 strict 都在一个 run 内）。
  async function applyPipetteToSelection() {
    if (!pipetteSource) {
      setPipetteState('idle');
      return;
    }
    stopLockMonitor();
    let result = null;
    try {
      await PowerPoint.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/width, items/height, items/adjustments, items/tags');
        await ctx.sync();
        result = await window.RadiusCore.applyPickedToSelection(
          driver, sel,
          { cm: pipetteSource.cm, sourceStrict: pipetteSource.sourceStrict },
          { syncStrict: pipetteSyncStrict }
        );
      });
    } catch (err) {
      showToast('样式刷失败：' + (err.message || err));
      if (selectedShapes.length > 0) startLockMonitor();
      return;
    }

    if (!result || !result.ok) {
      if (result && result.rejectReason === 'strict') {
        showToast('🔒 选区里有形状启用了防误触，样式刷不生效。先关掉目标的防误触或解锁。');
      } else if (result && result.applied === 0) {
        showToast('选区为空，先在 PPT 里选 1 个圆角矩形');
      } else if (result) {
        showToast('样式刷失败：' + (result.error || '未知错误'));
      }
      if (selectedShapes.length > 0) startLockMonitor();
      return;
    }

    // toast
    const lockHint = '';  // radius-core.applyPickedToSelection 已经处理 lock 同步
    const strictHint = result.strictSynced > 0
      ? `，${result.strictSynced} 个防误触状态已同步`
      : '';
    showToast(`🪣 样式刷应用了 ${result.applied} 个圆角矩形${result.failed > 0 ? `，${result.failed} 个失败` : ''}${lockHint}${strictHint}`);
    await refreshSelection();
    // v1.2: layout 父被刷 R 角 → 同步子 R 角
    await syncLayoutChildrenRIfNeeded();
    if (selectedShapes.length > 0) startLockMonitor();
  }

  // DocumentSelectionChanged 分发：idle → refreshSelection；sourcing → pickup；brushing → apply
  function onSelectionChangedForPipette() {
    if (pipetteState === 'sourcing') {
      pickupFromSelection();
    } else if (pipetteState === 'brushing') {
      applyPipetteToSelection();
    }
    // idle 状态由原 refreshSelection 处理
  }

  // ---------------- 事件绑定 ----------------

  function bindEvents() {
    // 调试日志：复制 / 清空按钮（点按钮不触发 details toggle）
    const copyBtn = $('debug-copy-btn');
    if (copyBtn) {
      copyBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        copyDebugLog();
      });
    }
    const clearBtn = $('debug-clear-btn');
    if (clearBtn) {
      clearBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        clearDebugLog();
      });
    }
    const smokeBtn = $('smoke-test-btn');
    if (smokeBtn) {
      smokeBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        runDriverSmokeTest();
      });
    }
    $('apply-btn').addEventListener('click', onApply);
    $('lock-btn').addEventListener('click', onToggleLock);
    $('reapply-btn').addEventListener('click', onReapply);
    $('rescan-btn').addEventListener('click', refreshSelection);
    $('radius-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') onApply();
    });
    $('radius-input').addEventListener('input', () => renderUI());
    document.querySelectorAll('.unit-btn').forEach((btn) => {
      btn.addEventListener('click', () => onUnitChange(btn.dataset.unit));
    });
    // v1.1：预设库 + 样式刷
    $('preset-add-btn').addEventListener('click', addPresetFromInput);
    $('pipette-btn').addEventListener('click', onPipetteButtonClick);
    // 样式刷：勾选「刷防误触状态」→ 同步源形状的防误触状态
    const syncCb = $('pipette-sync-checkbox');
    if (syncCb) {
      syncCb.addEventListener('change', () => {
        pipetteSyncStrict = syncCb.checked;
        // brushing 状态下立刻更新 hint
        if (pipetteState === 'brushing' && pipetteSource) {
          setPipetteState('brushing');
        }
      });
    }
    // 防误触开关
    const strictCb = $('strict-checkbox');
    if (strictCb) {
      strictCb.addEventListener('change', () => {
        if (strictCb.disabled) return;
        onToggleStrict(strictCb.checked);
      });
    }
    // v1.2: 布局模式控件
    // v1.2.11：rows/cols 走互斥联动（新函数，UI 是单滑块 + 列 readout，行 × 列 = 子数 N）
    bindLayoutGridRangeAndNum('layout-rows', 'layout-rows-num', 'layout-cols-readout');
    bindLayoutRangeAndNum('layout-padding', 'layout-padding-num', 'padding', false);
    bindLayoutRangeAndNum('layout-gutter', 'layout-gutter-num', 'gutter', false);
    // setup rows/cols 变化时重算 canBuild
    ['layout-setup-rows', 'layout-setup-cols'].forEach((id) => {
      const el = $(id);
      if (el) el.addEventListener('input', () => renderLayoutPanel());
    });
    $('layout-setup-btn').addEventListener('click', onLayoutSetup);
    $('layout-detach-btn').addEventListener('click', onLayoutDetach);
    $('layout-child-detach-btn').addEventListener('click', onLayoutChildDetach);
    // R 角联动模式 radio 组
    document.querySelectorAll('input[name="layout-link-r-mode"]').forEach((r) => {
      r.addEventListener('change', () => {
        if (!r.checked || !currentLayout) return;
        currentLayout.params.linkRMode = r.value;
        scheduleLayoutApply();
      });
    });
  }

  // ---------------- 初始化 ----------------

  Office.onReady(() => {
    bindEvents();
    renderPresets(userPresets); // 渲染空预设库
    refreshSelection();
    // 选区变化：分发到 pipette（sourcing → pickup；brushing → apply）或 refreshSelection（idle）
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => {
        if (pipetteState === 'idle') {
          refreshSelection();
        } else {
          onSelectionChangedForPipette();
          // pipette 路径也更新一下内存里的 selectedShapes（pickup/apply 完内存可能需要刷新）
          // 注意：sourcing 路径只读不写，brushing 路径写完会调 refreshSelection
          if (pipetteState === 'sourcing') {
            // 吸取成功可能进入 brushing 状态，brushing 路径自己处理刷新
            // sourcing 状态下 pickupFromSelection 不动 selectedShapes 数值
          }
        }
      }
    );
  });
})();
