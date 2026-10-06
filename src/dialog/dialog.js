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
  // applyLayout 的 group 安全事务会触发 ungroup/regroup 选区事件；
  // 事务未完成前禁止事件处理器启动并发 refreshSelection。
  let layoutMutationDepth = 0;
  let activeMutationDriver = null;
  let selectionEpoch = 0;
  let monitorGeneration = 0;
  let monitorInFlight = false;
  let selectionShapeKinds = new Map();
  let selectionSlideId = null;
  let ignoredRestoredSelection = null;
  let selectionEventProbeInFlight = false;

  async function runMutation(callback, epoch) {
    epoch = epoch == null ? selectionEpoch : epoch;
    stopLockMonitor();
    const result = await window.PptDriver.run(async (ctx) => {
      if (epoch !== selectionEpoch) return { ok: false, reason: 'stale-selection' };
      const driver = window.PptDriver.createDriver(ctx);
      activeMutationDriver = driver;
      layoutMutationDepth++;
      try { return await callback(driver, () => epoch === selectionEpoch); }
      finally {
        if (driver.structuralChanged) {
          layoutSelectionIgnoreUntil = Date.now() + 300;
          ignoredRestoredSelection = { slideId: selectionSlideId, shapeIds: driver.restoredSelectionIds };
        }
        layoutMutationDepth--;
        activeMutationDriver = null;
      }
    }, () => epoch === selectionEpoch);
    return result || { ok: false, reason: 'stale-selection' };
  }
  // Mac LTSC 会把 regroup 后的 DocumentSelectionChanged 延迟到 PowerPoint.run
  // 返回之后再派发；给事务后的主动 refreshSelection 留一个很短的保护窗口，
  // 避免同一次内部选区恢复再启动第二个并发刷新。
  let layoutSelectionIgnoreUntil = 0;

  // 当前输入单位：'cm' | '%'
  let currentUnit = 'cm';

  // v1.2.14：边距/间距锁链联动（Photoshop 风格）
  // linkPG=true：间距 slider/num 变灰 + 不可用，间距值 = 边距值（自动同步）
  // linkPG=false：默认状态，间距独立
  // 持久化：只在内存（关 task pane 后失效——跟 history / preset 一致）
  let linkPG = false;

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
    lastAdj: {},
    candidateAdj: {},
    groupLockUpdates: {},       // shapeId -> 上次读到的 adj
    stableCount: {},   // shapeId -> adj 连续稳定次数
    lastCm: {},        // shapeId -> 上次读到的 currentCm（cm）—— 仅 layout 父用
    lastSizeCm: {},    // shapeId -> 上次读到的 { widthCm, heightCm } —— 仅 layout 父用（v1.2.9 size 联动）
    parentRDirty: false,  // 是否有 layout 父 R 角或 size 变化需要联动（v1.2.9 合并）
    parentRSyncGeometry: false,
    parentRSyncTimer: null,  // 节流 timer（避免 10ms tick 频繁触发新 run）
    groupLayoutSyncTimer: null, // group 拖拽停止后，安全解组重排布局
  };
  const PARENT_R_SYNC_DEBOUNCE_MS = 200;  // 父 R 角变化 → 同步子的节流窗口
  const GROUP_LAYOUT_SYNC_DEBOUNCE_MS = 300; // group 尺寸稳定 300ms 后再解组重排

  // ---------------- 单位换算 ----------------

  function getRefShapeMinSideCm() {
    // % 模式的 100% 参考：用第一个 roundRect 的 minSide
    for (const s of selectedShapes) {
      if (s.isRoundRect && s.minSideCm > 0) return s.minSideCm;
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
  async function loadLocksViaTags() {
    try {
      return await window.PptDriver.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        const leaves = await driver.loadShapeTree(driver.selectedShapes(), 'id, tags');
        const tags = await driver.loadTagsBulk(leaves);
        const locks = {}, strict = {};
        for (const sh of leaves) {
          const id = driver.shapeId(sh);
          const state = window.RadiusCore.lockStateFromTags(tags[id]);
          if (state.isLocked) locks[id] = state.lockedCm;
          if (state.isStrict) strict[id] = true;
        }
        return { ok: true, locks, strict };
      });
    } catch (error) { return { ok: false, error }; }
  }

  async function saveLocksViaTags(locks, strictMap) {
    try {
      return await runMutation(async (driver, isCurrent) => {
        const leaves = await driver.loadShapeTree(driver.selectedShapes(), 'id, tags');
        if (!isCurrent()) return { ok: false, reason: 'stale-selection' };
        const ids = leaves.map((sh) => driver.shapeId(sh)).filter((id) => id in locks || id in (strictMap || {}));
        await window.RadiusCore.withWritableShapes(driver, ids, async (fresh) => {
          if (!isCurrent()) return;
          for (const sh of fresh) {
            const id = driver.shapeId(sh);
            const state = {};
            if (id in locks) state.lockedCm = locks[id];
            if (id in (strictMap || {})) state.isStrict = !!strictMap[id];
            const r = await window.RadiusCore.writeLockState(driver, sh, state);
            if (!r.ok) throw new Error(r.error);
          }
          await driver.sync();
        }, { isCurrent });
        return { ok: true };
      });
    } catch (error) { return { ok: false, error }; }
  }

  async function updateLockTagForShape(shapeId, cm, isStrict) {
    const locks = {}, strict = {};
    if (cm !== undefined) locks[shapeId] = cm;
    if (isStrict === true || isStrict === false) strict[shapeId] = isStrict;
    return saveLocksViaTags(locks, strict);
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
    return window.RadiusCore.computeLayout(parent, rows, cols, paddingCm, gutterCm);
  }

  // 把当前 params 应用到子形状：写位置 + 尺寸 + R 角（如果 linkR）
  // parentId: 父 shape id
  // params: { rows, cols, padding, gutter, linkR }
  // childIds: [string]
  // opts: { writeParentTag: bool, syncR: bool }
  //   - writeParentTag: 写 layoutParent_v1 tag（首次创建时 true；只调参数时 false）
  //   - syncR: 是否同步 R 角（链接时 true，单纯调位置时 false）
  //   - writeGeometry: 是否写子位置/尺寸（切换 R 联动模式时 false）
  // 返回：{ ok, applied, failed, warn }
  // v1.2.9 迁移：applyLayoutToChildren 第一道防线（strict 检查）保留在 dialog.js（依赖 selectedShapes 状态），
  // PowerPoint.run 部分（写位置/尺寸/R 角/父 tag + 过滤 stale childIds）全部走 radius-core.applyLayout
  async function applyLayoutToChildren(parentId, params, childIds, opts) {
    opts = opts || {};
    const epoch = opts.epoch == null ? selectionEpoch : opts.epoch;
    if (epoch !== selectionEpoch) return { ok: false, reason: 'stale-selection', applied: 0, failed: 0 };
    const writeParentTag = opts.writeParentTag !== false;
    const syncR = opts.syncR !== false;
    const writeGeometry = opts.writeGeometry !== false;
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

    try {
      return await runMutation((driver, isCurrent) => window.RadiusCore.applyLayout(
        driver, parentId, { ...params }, childIds.slice(),
        { writeParentTag, syncR, writeGeometry, isCurrent }
      ), epoch);
    } catch (e) {
      return { ok: false, applied: 0, failed: 0, error: e.message || String(e) };
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
      return await window.PptDriver.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        return await window.RadiusCore.syncLayoutChildrenR(driver, parentId, childIds, paddingCm, linkRMode, parentRcm);
      });
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      return { ok: false, applied: 0, failed: 0, error: msg };
    }
  }

  // 检测选区里是否有 layout 父 → 同步其子 R 角（onApply / applyPipette 末尾调用）
  // Geometry and R-only synchronization share the safe applyLayout path.
  // v1.3.6：v1.2 step 3 调试 log 清理
  async function syncLayoutChildrenRIfNeeded(opts) {
    if (selectedShapes.length === 0) return { geometry: true, results: [] };
    opts = opts || {};
    const epoch = opts.epoch == null ? selectionEpoch : opts.epoch;
    const geometry = opts.geometry !== false;
    const results = [];
    for (const s of selectedShapes) {
      if (epoch !== selectionEpoch) break;
      if (s.layoutRole === 'parent' && s.layoutParams && s.layoutChildIds) {
        // 保留存量布局的 R 模式；关闭 R 联动仍允许几何联动。
        // 老 layout（tag 里存了 'subtract'）直接拿到旧值不受影响
        const linkRMode = s.layoutParams.linkRMode || 'same';
        if (linkRMode === 'off' && !geometry) continue;
        const expected = s.layoutParams.rows * s.layoutParams.cols;
        const childIds = s.layoutChildIds.slice(0, expected);
        if (geometry) {
          // 普通叶子选区：父 size/R 变化后完整重算 layout 几何 + R。
          const params = { ...s.layoutParams, linkRMode };
          results.push(await applyLayoutToChildren(
            s.id,
            params,
            childIds,
            { writeParentTag: false, syncR: true, epoch }
          ));
        } else {
          // R-only still uses the safe transaction when children are grouped.
          results.push(await applyLayoutToChildren(s.id, { ...s.layoutParams }, childIds,
            { writeParentTag: false, syncR: true, writeGeometry: false, epoch }));
        }
      }
    }
    return { geometry, results };
  }

  // ---------------- v1.2: layout UI 渲染 + 交互 ----------------

  // 节流：滑块拖动时只在最后一次输入后 apply
  let layoutApplyTimer = null;
  let layoutApplyPending = false;
  let layoutApplyPendingOpts = null;
  let layoutRequestSerial = 0;

  // 方案 A：用户手动指定父/子。{ parentId, childIds[] }
  // 选区变化时自动重置（renderLayoutSetupList 检测 roundRect 列表变化）
  let layoutSetupChoices = { parentId: null, childIds: [] };
  let layoutSetupListSignature = '';  // 上次渲染的列表签名（用 roundRect ids 拼接），变了就重置 choices

  function scheduleLayoutApply(opts) {
    const request = ++layoutRequestSerial;
    const incoming = {
      writeParentTag: true,
      syncR: true,
      writeGeometry: !(opts && opts.writeGeometry === false),
    };
    layoutApplyPending = true;
    if (layoutApplyPendingOpts) {
      // 如果同一个 50ms 窗口里既改了几何参数又切换了 R 模式，必须保留几何写入；
      // 单独切 R 模式时则保持 R-only，避免一次无关的全布局反算。
      layoutApplyPendingOpts.writeParentTag =
        layoutApplyPendingOpts.writeParentTag || incoming.writeParentTag;
      layoutApplyPendingOpts.syncR =
        layoutApplyPendingOpts.syncR || incoming.syncR;
      layoutApplyPendingOpts.writeGeometry =
        layoutApplyPendingOpts.writeGeometry || incoming.writeGeometry;
    } else {
      layoutApplyPendingOpts = incoming;
    }
    if (layoutApplyTimer) clearTimeout(layoutApplyTimer);
    layoutApplyTimer = setTimeout(() => {
      layoutApplyTimer = null;
      if (layoutApplyPending && currentLayout) {
        const pendingOpts = layoutApplyPendingOpts || {
          writeParentTag: true,
          syncR: true,
          writeGeometry: true,
        };
        layoutApplyPending = false;
        layoutApplyPendingOpts = null;
        applyLayoutFromUI(pendingOpts, request);
      }
    }, LAYOUT_LAYOUT_RT_DEBOUNCE_MS);
  }

  // 立刻 apply（用户改 rows/cols 触发，因为会改变 childIds 数量）
  async function applyLayoutFromUI(opts, request) {
    if (!currentLayout) return;
    request = request == null ? ++layoutRequestSerial : request;
    const epoch = selectionEpoch;
    const parentId = currentLayout.parentId;
    const params = { ...currentLayout.params };
    const childIds = currentLayout.childIds.slice();
    stopLockMonitor();
    const r = await applyLayoutToChildren(parentId, params, childIds, opts || { writeParentTag: true, syncR: true });
    if (epoch !== selectionEpoch || request !== layoutRequestSerial) return;
    if (!r.ok) {
      showToast(i18n.t('toastLayoutFailedFmt', { error: r.error || r.warn || r.reason || (i18n.getLang() === 'zh' ? '未知错误' : 'unknown error') }));
    } else if (r.warn) {
      if (r.warn.indexOf('子形状不足') >= 0) {
        const actualRows = Math.max(1, Math.floor(childIds.length / params.cols));
        currentLayout.params.rows = actualRows;
        renderLayoutPanel();
        showToast(i18n.t('toastLayoutAutoReducedFmt', { warn: r.warn, rows: actualRows }));
        return applyLayoutFromUI({ writeParentTag: true, syncR: true });
      }
      showToast(r.warn);
    } else {
      const rHint = params.linkRMode && params.linkRMode !== 'off' ? '（含 R 角联动）' : '';
      const strictHint = '';
      showToast(i18n.t('toastLayoutAppliedFmt', { applied: r.applied, failed: r.failed ? i18n.t('failedStrFmt', { count: r.failed }) : '', rHint: rHint, strictHint: strictHint }));
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
      hint.textContent = i18n.t('layoutHintEmpty');
      renderLayoutSetupList([]);
      empty.style.display = 'flex';
      $('layout-setup-btn').disabled = true;
      return;
    }
    // 选区里有 layout 父 → 显示 active 面板
    const parentShape = selectedShapes.find((s) => s.layoutRole === 'parent');
    if (parentShape && currentLayout) {
      hint.textContent = i18n.t('layoutActive');
      active.style.display = 'flex';
      $('layout-parent-name').textContent = currentLayout.parentName;
      $('layout-children-count').textContent = i18n.t('layoutChildrenCountFmt', { count: currentLayout.childIds.length, rows: currentLayout.params.rows, cols: currentLayout.params.cols });
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
      const pgPair = $('layout-pg-pair');
      const pgLinkBtn = $('layout-pg-link');
      const pgLinkIcon = $('layout-pg-link-icon');
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
          if (coupledHint) coupledHint.textContent = i18n.t('layoutCoupledOneChild');
        } else {
          rowsR.disabled = false;
          rowsN.disabled = false;
          if (coupledHint) coupledHint.textContent = i18n.t('layoutCoupledFactorFmt', { N: N, factors: factors.join('/') });
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
      // v1.2.14：锁链联动状态下，间距初始 = 边距（首次开启时同步）
      // v1.2.15 改：直接禁用 gutter 控件（disabled 属性，不只是 CSS pointer-events）
      // 原因：用户报 bug —— 在链接状态修改间距 → 形状被改，间距被锁链改回 → 形状没被改回
      //       根因链不好根除（lockMonitor / 节流 / apply 异步返回等竞态），最稳的做法是**不让用户能改**
      //       disabled 是 input 原生属性，挡 mouse + keyboard + focus + input 事件，比 CSS pointer-events 彻底
      if (linkPG) {
        if (currentLayout.params.gutter !== currentLayout.params.padding) {
          currentLayout.params.gutter = currentLayout.params.padding;
        }
        gutR.value = String(currentLayout.params.padding);
        gutN.value = currentLayout.params.padding.toFixed(2);
        gutR.disabled = true;
        gutN.disabled = true;
      } else {
        gutR.value = String(currentLayout.params.gutter);
        gutN.value = currentLayout.params.gutter.toFixed(2);
        gutR.disabled = false;
        gutN.disabled = false;
      }
      // 锁链 button 状态（只用 emoji + 颜色指示，不写"已联动"/"未联动"文字）
      if (pgLinkBtn) {
        pgLinkBtn.dataset.linked = linkPG ? 'true' : 'false';
        // 图标始终保持为链条；仅通过按钮背景色区分联动状态。
        if (pgLinkIcon) pgLinkIcon.textContent = '🔗';
        pgLinkBtn.title = linkPG
          ? '锁链已激活：间距跟随边距（点解开）'
          : '锁链解开：间距独立（点激活）';
      }
      // pair 容器 data-linked → CSS 控制 gutter 变灰 + 不可用
      if (pgPair) pgPair.dataset.linked = linkPG ? 'true' : 'false';
      // 保留既有默认 same；subtract 是单层嵌套的等距缩进公式。
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
        warn.textContent = i18n.t('layoutWarnTooTight');
      } else if (childCount < expected) {
        // v1.2.11：理论上 N=rows*cols（联动保证），但旧的 layout tag 可能不对
        warn.textContent = i18n.t('layoutWarnNotEnoughFmt', { expected: expected, childCount: childCount });
      } else {
        warn.textContent = '';
      }
      updateLayoutPreview();
      return;
    }
    // 选区里只有子（无父）
    const childShape = selectedShapes.find((s) => s.layoutRole === 'child');
    if (childShape) {
      hint.textContent = i18n.t('layoutHintChildShape');
      childInfo.style.display = 'flex';
      const parentInSel = selectedShapes.find((s) => s.id === childShape.layoutParentId);
      $('layout-child-parent').textContent = parentInSel ? (parentInSel.name || i18n.t('unnamed')) : i18n.t('layoutHintNotInSel');
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
      hint.textContent = i18n.t('layoutHintEmpty');
    } else if (strictCount > 0) {
      // 防误触：选区里有 N 个子启用了防误触 → 拒绝整个 setup
      hint.textContent = i18n.t('layoutHintStrictFmt', { count: strictCount });
      $('layout-setup-btn').disabled = true;  // 再次确认按钮禁用
    } else if (!canBuild) {
      const missing = !layoutSetupChoices.parentId
        ? i18n.t('toastSelectParent')
        : layoutSetupChoices.childIds.length < need
          ? i18n.t('toastChildrenInsufficientFmt', { need: need, chosen: layoutSetupChoices.childIds.length })
          : i18n.t('toastRowsRange') + ' / ' + i18n.t('toastColsRange');
      hint.textContent = missing;
    } else {
      hint.textContent = i18n.t('layoutHintCanBuildFmt', { rows: rows, cols: cols });
    }
    empty.style.display = 'flex';
  }

  // 滑块 / 数字输入同步 + 触发 apply
  function bindLayoutRangeAndNum(rangeId, numId, paramKey, isInt) {
    const r = $(rangeId);
    const n = $(numId);
    if (!r || !n) return;

    // v1.2.14：边距/间距锁链联动辅助函数
    // 改 padding 时，如果 linkPG=true → 把 gutter 也同步到 padding 值
    function syncLinkedPartner() {
      if (!linkPG) return;
      if (paramKey !== 'padding') return;  // 改 gutter 不再触发反向（避免循环）
      const gutR = $('layout-gutter');
      const gutN = $('layout-gutter-num');
      if (gutR) gutR.value = r.value;
      if (gutN) gutN.value = r.value;
      if (currentLayout) {
        currentLayout.params.gutter = parseFloat(r.value);
      }
    }

    r.addEventListener('input', () => {
      n.value = r.value;
      if (!currentLayout) return;
      currentLayout.params[paramKey] = isInt ? parseInt(r.value, 10) : parseFloat(r.value);
      syncLinkedPartner();
      updateLayoutPreview();
      scheduleLayoutApply();
    });
    n.addEventListener('input', () => {
      let v = isInt ? parseInt(n.value, 10) : parseFloat(n.value);
      if (!Number.isFinite(v)) return;
      v = Math.max(parseFloat(r.min), Math.min(parseFloat(r.max), v));
      r.value = String(v);
      if (!currentLayout) return;
      currentLayout.params[paramKey] = v;
      syncLinkedPartner();
      updateLayoutPreview();
      scheduleLayoutApply();
    });
    n.addEventListener('change', () => {
      if (currentLayout && layoutApplyTimer) {
        clearTimeout(layoutApplyTimer);
        layoutApplyTimer = null;
        layoutApplyPending = false;
        layoutApplyPendingOpts = null;
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
        layoutApplyPendingOpts = null;
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
      empty.textContent = i18n.t('layoutHintEmpty');
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
      nameSpan.textContent = sh.name || i18n.t('unnamed');
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
      parentSpan.textContent = i18n.t('layoutLabelParent');
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
      childSpan.textContent = i18n.t('layoutLabelChildren');
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
      el.textContent = i18n.t('layoutPreviewWarn');
      return;
    }
    el.textContent = i18n.t('layoutPreviewSizeFmt', { w: (r.subW / PT_PER_CM).toFixed(2), h: (r.subH / PT_PER_CM).toFixed(2) });
  }

  // 从选区建立布局：使用 layoutSetupChoices（用户手动指定的父/子）
  async function onLayoutSetup() {
    const rows = parseInt($('layout-setup-rows').value, 10);
    const cols = parseInt($('layout-setup-cols').value, 10);
    if (!Number.isFinite(rows) || rows < 1 || rows > 5) {
      showToast(i18n.t('toastRowsRange'));
      return;
    }
    if (!Number.isFinite(cols) || cols < 1 || cols > 5) {
      showToast(i18n.t('toastColsRange'));
      return;
    }
    const need = rows * cols;
    if (!layoutSetupChoices.parentId) {
      showToast(i18n.t('toastSelectParent'));
      return;
    }
    if (layoutSetupChoices.childIds.length < need) {
      showToast(i18n.t('toastChildrenInsufficientFmt', { need: need, chosen: layoutSetupChoices.childIds.length }));
      return;
    }
    const parentId = layoutSetupChoices.parentId;
    const childIds = layoutSetupChoices.childIds.slice(0, need);
    // 新布局保持默认 same（相同 R），旧布局的模式不变。
    const params = { rows, cols, padding: 0.5, gutter: 0.3, linkRMode: 'same' };
    stopLockMonitor();
    const r = await applyLayoutToChildren(parentId, params, childIds, { writeParentTag: true, syncR: true });
    if (!r.ok) {
      showToast(i18n.t('toastBuildLayoutFailedFmt', { error: r.error || r.warn || (i18n.getLang() === 'zh' ? '未知错误' : 'unknown error') }));
    } else if (r.applied === 0) {
      showToast(i18n.t('toastNoChildrenWrittenFmt', { warn: r.warn || (i18n.getLang() === 'zh' ? '未知问题，看 console' : 'unknown issue, see console') }));
    } else {
      showToast(i18n.t('toastLayoutBuiltFmt', { rows: rows, cols: cols, applied: r.applied, failed: r.failed ? i18n.t('failedStrFmt', { count: r.failed }) : '' }));
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
      showToast(i18n.t('toastDetachFailedFmt', { error: r.error || r.warn || r.reason || (i18n.getLang() === 'zh' ? '未知错误' : 'unknown error') }));
    } else {
      showToast(i18n.t('toastDetached'));
    }
    await refreshSelection();
    if (selectedShapes.length > 0) startLockMonitor();
  }

  // 子形状脱离（只删自己的 child tag + 从父的 childIds 移除）
  // 只在当前 slide 操作（不跨页）
  async function onLayoutChildDetach() {
    const child = selectedShapes.find((s) => s.layoutRole === 'child');
    if (!child) return;
    try {
      const result = await runMutation(async (driver, isCurrent) => {
        const leaves = await driver.loadShapeTree(driver.slideShapes(driver.activeSlide()), 'id, tags');
        const parent = leaves.find((sh) => driver.shapeId(sh) === child.layoutParentId);
        const parsed = parent ? window.RadiusCore.parseLayoutParentTagValue(await driver.readTag(parent, LAYOUT_PARENT_TAG_KEY)) : null;
        const ids = [child.id];
        if (parent && parsed) ids.push(child.layoutParentId);
        await window.RadiusCore.withWritableShapes(driver, ids, async (fresh) => {
          for (const sh of fresh) {
            if (driver.shapeId(sh) === child.id) driver.deleteTag(sh, LAYOUT_CHILD_TAG_KEY);
            else driver.addTag(sh, LAYOUT_PARENT_TAG_KEY, JSON.stringify({ ...parsed,
              childIds: parsed.childIds.filter((id) => id !== child.id) }));
          }
          await driver.sync();
        }, { isCurrent });
        return { ok: true };
      });
      if (!result.ok) throw new Error(result.error || result.reason);
      showToast(i18n.t('toastDetachedThis'));
    } catch (error) { showToast(i18n.t('toastDetachFailedFmt', { error: error.message || error })); }
    await refreshSelection();
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
        showToast(i18n.t('toastCopiedFmt', { count: lines.length }));
      }).catch((err) => {
        // fallback
        fallbackCopy(text);
        showToast(i18n.t('toastCopiedFallbackFmt', { count: lines.length }));
      });
    } else {
      fallbackCopy(text);
      showToast(i18n.t('toastCopiedFallbackFmt', { count: lines.length }));
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
    showToast(i18n.t('toastLogCleared'));
  }

  // ============================================================
  // Driver 烟囱测试（v1.2.3）
  // 跑遍所有 15 个 driver 方法 + 在真实 PPT 里读/写一遍
  // 每个方法的输入/输出都打 log，复制整段发我就能看到每个模块状态
  // ============================================================
  let driverSmokeInFlight = false;
  async function runDriverSmokeTest() {
    // The test spans separate host runs for fresh readback and cleanup. A second
    // click must not capture the first test's temporary values as its originals.
    if (driverSmokeInFlight) return;
    driverSmokeInFlight = true;
    const button = $('smoke-test-btn');
    if (button) button.disabled = true;
    try { await runDriverSmokeTestOnce(); }
    finally {
      driverSmokeInFlight = false;
      if (button) button.disabled = false;
    }
  }

  async function runDriverSmokeTestOnce() {
    const epoch = selectionEpoch;
    stopLockMonitor();
    const tests = [];
    const record = (name, ok, detail) => {
      tests.push({ name, ok });
      console.log(`[smoke] ${ok ? '✅' : '❌'} ${name}: ${detail || ''}`);
    };
    let original = null;
    try {
      await window.PptDriver.run(async (ctx) => {
        const d = window.PptDriver.createDriver(ctx);
        const sel = d.selectedShapes();
        record('selectedShapes', !!sel);
        const slide = d.activeSlide();
        d.load(slide, 'id');
        record('activeSlide', !!slide);
        record('slideShapes', !!d.slideShapes(slide));
        const leaves = await d.loadShapeTree(sel, 'id, name, left, top, width, height, adjustments, tags');
        record('loadShapeTree + sync', leaves.length > 0);
        const shape = leaves.find((s) => d.isRoundRect(s));
        if (!shape || d.hasTopLevelGroup(sel)) throw new Error('请选一个未组合的圆角矩形运行烟囱测试');
        const tags = await d.loadTagsBulk([shape]);
        const protection = window.RadiusCore.lockStateFromTags(tags[d.shapeId(shape)]);
        if (protection.isStrict || protection.isLocked) throw new Error('请选一个未锁定的圆角矩形运行烟囱测试');
        const value = await d.readAdjFraction(shape);
        if (epoch !== selectionEpoch) throw new Error('测试选区已改变');
        const tagKey = Object.keys(tags[d.shapeId(shape)]).find((key) => key.toUpperCase() === 'DRIVER_SMOKE_TEST_V1');
        original = { id: d.shapeId(shape), slideId: slide.id, box: d.box(shape), adj: value,
          tagKey: tagKey || 'driver_smoke_test_v1', tagValue: tagKey == null ? null : tags[d.shapeId(shape)][tagKey] };
        record('shapeId', typeof original.id === 'string');
        record('isRoundRect', d.isRoundRect(shape));
        record('readAdjFraction', Number.isFinite(value));
        record('size', d.size(shape).width > 0);
        record('box', Number.isFinite(original.box.left));
        const key = original.tagKey;
        d.addTag(shape, key, 'test');
        await d.sync();
        record('addTag + readTag', await d.readTag(shape, key) === 'test');
        d.deleteTag(shape, key);
        await d.sync();
        record('deleteTag + readTag', await d.readTag(shape, key) === null);
        const bulk = await d.loadTagsBulk([shape]);
        record('loadTagsBulk', !!bulk[original.id]);
        if (epoch !== selectionEpoch) throw new Error('测试选区已改变');
        d.setAdjFraction(shape, value > 0.25 ? 0.1 : 0.4);
        d.setBox(shape, { ...original.box, left: original.box.left + 10 });
        await d.sync();
      });
      await window.PptDriver.run(async (ctx) => {
        if (epoch !== selectionEpoch) throw new Error('测试选区已改变');
        const d = window.PptDriver.createDriver(ctx);
        const identity = await d.loadSelectionIdentity();
        if (identity.slideId !== original.slideId || !identity.shapeIds.includes(original.id)) throw new Error('测试选区已改变');
        const leaves = await d.loadShapeTree(d.selectedShapes(), 'id, left, top, width, height, adjustments');
        const shape = leaves.find((s) => d.shapeId(s) === original.id);
        if (!shape) throw new Error('测试选区已改变');
        record('setAdjFraction readback', Math.abs(await d.readAdjFraction(shape) - (original.adj > 0.25 ? 0.1 : 0.4)) < 0.001);
        record('setBox readback', Math.abs(d.box(shape).left - original.box.left - 10) < 0.5);
      });
    } catch (error) {
      record('host operation', false, error.message || String(error));
    } finally {
      if (original) {
        try {
          await window.PptDriver.run(async (ctx) => {
            const d = window.PptDriver.createDriver(ctx);
            const leaves = await d.loadShapeTree(d.slideShapes(d.slideById(original.slideId)), 'id');
            const shape = leaves.find((s) => d.shapeId(s) === original.id);
            if (!shape) throw new Error('无法找回测试形状');
            d.setAdjFraction(shape, original.adj);
            d.setBox(shape, original.box);
            if (original.tagValue == null) d.deleteTag(shape, original.tagKey);
            else d.addTag(shape, original.tagKey, original.tagValue);
            await d.sync();
          });
        } catch (error) { record('restore', false, error.message || String(error)); }
      }
      console.log(`[smoke] ${tests.filter((t) => t.ok).length}/${tests.length} passed`);
      await refreshSelection();
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
  async function loadLayoutTagsViaTags() {
    try {
      return await window.PptDriver.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        const leaves = await driver.loadShapeTree(driver.selectedShapes(), 'id, tags');
        const all = await driver.loadShapeTree(driver.slideShapes(driver.activeSlide()), 'id');
        return window.RadiusCore.loadLayoutTags(driver, leaves, all);
      });
    } catch (error) { return { ok: false, error }; }
  }

  async function saveLayoutTags(parentId, params, childIds) {
    try {
      return await runMutation((driver, isCurrent) => window.RadiusCore.saveLayoutTags(
        driver, driver.activeSlide(), parentId, { ...params }, (childIds || []).slice(), { isCurrent }
      ));
    } catch (error) { return { ok: false, error }; }
  }

  async function deleteLayoutTags(parentId, childIds) {
    try {
      return await runMutation(async (driver, isCurrent) => {
        const leaves = await driver.loadShapeTree(driver.slideShapes(driver.activeSlide()), 'id');
        if (!isCurrent()) return { ok: false, reason: 'stale-selection' };
        const ids = leaves.map((sh) => driver.shapeId(sh)).filter((id) => id === parentId || (childIds || []).includes(id));
        await window.RadiusCore.withWritableShapes(driver, ids, async (fresh) => {
          for (const sh of fresh) {
            const id = driver.shapeId(sh);
            if (id === parentId) driver.deleteTag(sh, LAYOUT_PARENT_TAG_KEY);
            if ((childIds || []).includes(id)) driver.deleteTag(sh, LAYOUT_CHILD_TAG_KEY);
          }
          await driver.sync();
        }, { isCurrent });
        return { ok: true };
      });
    } catch (error) { return { ok: false, error }; }
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
    const epoch = ++selectionEpoch;
    const isCurrent = () => epoch === selectionEpoch;
    stopLockMonitor();
    if (layoutApplyTimer) clearTimeout(layoutApplyTimer);
    layoutApplyTimer = null;
    layoutApplyPending = false;
    layoutApplyPendingOpts = null;
    try {
      const snapshot = await window.PptDriver.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        const identity = await driver.loadSelectionIdentity();
        const leaves = await driver.loadShapeTree(driver.selectedShapes(), 'id, name, width, height, adjustments, tags');
        if (!isCurrent()) return null;
        const tags = await driver.loadTagsBulk(leaves);
        const shapes = [];
        for (const sh of leaves) {
          if (!isCurrent()) return null;
          const size = driver.size(sh);
          const minSideCm = Math.min(size.width, size.height) / PT_PER_CM;
          const isRoundRect = driver.isRoundRect(sh);
          const value = isRoundRect ? await driver.readAdjFraction(sh) : null;
          const id = driver.shapeId(sh);
          const lock = window.RadiusCore.lockStateFromTags(tags[id]);
          const currentCm = isRoundRect && Number.isFinite(value) && value >= 0 ? value * minSideCm : null;
          shapes.push({ id, name: driver.shapeName(sh), width: size.width, height: size.height,
            minSideCm, currentCm, isRoundRect,
            locked: lock.isLocked || lock.isStrict && isRoundRect && currentCm != null,
            lockedCm: lock.isLocked ? lock.lockedCm : lock.isStrict ? currentCm : null,
            legacyStrictNeedsRepair: lock.isStrict && !lock.isLocked && isRoundRect && currentCm != null,
            strictLocked: lock.isStrict,
            layoutRole: null, layoutParentId: null, layoutParams: null, layoutChildIds: null });
        }
        if (!isCurrent()) return null;
        const all = await driver.loadShapeTree(driver.slideShapes(driver.activeSlide()), 'id');
        const layout = await window.RadiusCore.loadLayoutTags(driver, leaves, all);
        if (!layout.ok) throw new Error(layout.error);
        const kinds = leaves.length ? await driver.loadShapeKinds() : new Map();
        return { shapes, layout, kinds, identity };
      }, isCurrent);
      if (!isCurrent() || !snapshot) return;
      selectedShapes = snapshot.shapes;
      selectionShapeKinds = snapshot.kinds;
      selectionSlideId = snapshot.identity.slideId;
      const layoutResult = snapshot.layout;
      for (const shape of selectedShapes) {
        const parent = layoutResult.parents[shape.id];
        if (parent) {
          shape.layoutRole = 'parent';
          shape.layoutParams = { rows: parent.rows, cols: parent.cols, padding: parent.padding,
            gutter: parent.gutter, linkRMode: parent.linkRMode };
          shape.layoutChildIds = parent.childIds;
        } else if (layoutResult.childOf[shape.id]) {
          shape.layoutRole = 'child';
          shape.layoutParentId = layoutResult.childOf[shape.id];
        }
      }
      const parent = selectedShapes.find((shape) => shape.layoutRole === 'parent');
      currentLayout = parent ? { parentId: parent.id, parentName: parent.name || '(未命名)',
        childIds: parent.layoutChildIds.slice(), params: { ...parent.layoutParams } } : null;
      const staleCount = Object.values(layoutResult.staleParents || {}).reduce((n, ids) => n + ids.length, 0);
      if (staleCount) showToast(i18n.t('toastStaleLayoutsFmt', { count: staleCount }));
      renderUI();
      if (selectedShapes.length) startLockMonitor();
      if (selectedShapes.some((shape) => shape.legacyStrictNeedsRepair)) scheduleGroupLayoutSync({ geometry: false });
      return epoch;
    } catch (error) {
      if (!isCurrent()) return;
      selectedShapes = [];
      currentLayout = null;
      selectionShapeKinds = new Map();
      renderUI();
      console.log('[refreshSelection] EXCEPTION:', error.message || error);
      setStatus('选区', '读失败：' + (error.message || error), 'status-warn');
      showToast(i18n.t('toastReadFailedFmt', { error: error.message || error }));
    }
  }

  // ---------------- lock monitor：检测拖完松手后自动重应用 ----------------

  function startLockMonitor() {
    if (lockMonitor.timer || layoutMutationDepth > 0 || window.PptDriver.isBusy()) return;
    if (selectedShapes.length === 0) return;
    // 选区里有 locked shape 用 10ms 实时反算；只有未锁定的用 50ms 减负
    const interval = selectedShapes.some((s) => s.locked) ? LOCK_POLL_MS : IDLE_POLL_MS;
    lockMonitor.lastWidth = {};
    lockMonitor.lastHeight = {};
    lockMonitor.lastAdj = {};
    lockMonitor.candidateAdj = {};
    lockMonitor.groupLockUpdates = {};
    lockMonitor.stableCount = {};
    lockMonitor.lastCm = {};
    lockMonitor.lastSizeCm = {};  // v1.2.9
    lockMonitor.parentRDirty = false;
    lockMonitor.parentRSyncGeometry = false;
    if (lockMonitor.parentRSyncTimer) {
      clearTimeout(lockMonitor.parentRSyncTimer);
      lockMonitor.parentRSyncTimer = null;
    }
    if (lockMonitor.groupLayoutSyncTimer) {
      clearTimeout(lockMonitor.groupLayoutSyncTimer);
      lockMonitor.groupLayoutSyncTimer = null;
    }
    lockMonitor.timer = setInterval(monitorTick, interval);
  }

  function stopLockMonitor() {
    monitorGeneration++;
    if (lockMonitor.timer) {
      clearInterval(lockMonitor.timer);
      lockMonitor.timer = null;
    }
    lockMonitor.lastWidth = {};
    lockMonitor.lastHeight = {};
    lockMonitor.lastAdj = {};
    lockMonitor.candidateAdj = {};
    lockMonitor.groupLockUpdates = {};
    lockMonitor.stableCount = {};
    lockMonitor.lastCm = {};
    lockMonitor.lastSizeCm = {};  // v1.2.9
    lockMonitor.parentRDirty = false;
    lockMonitor.parentRSyncGeometry = false;
    if (lockMonitor.parentRSyncTimer) {
      clearTimeout(lockMonitor.parentRSyncTimer);
      lockMonitor.parentRSyncTimer = null;
    }
    if (lockMonitor.groupLayoutSyncTimer) {
      clearTimeout(lockMonitor.groupLayoutSyncTimer);
      lockMonitor.groupLayoutSyncTimer = null;
    }
  }

  // 反算某个 shape 的 adj（被 onApply / 样式刷 / 重置时调用，让 monitor 同步状态）
  function syncLockMonitorForShape(shapeId, adj) {
    lockMonitor.lastAdj[shapeId] = adj;
    lockMonitor.stableCount[shapeId] = 0;
  }

  async function monitorTick() {
    if (monitorInFlight || layoutMutationDepth > 0 || window.PptDriver.isBusy()) return;
    if (!selectedShapes.length) { stopLockMonitor(); return; }
    const generation = monitorGeneration;
    const epoch = selectionEpoch;
    const isCurrent = () => generation === monitorGeneration && epoch === selectionEpoch && layoutMutationDepth === 0;
    monitorInFlight = true;
    let needRefreshUI = false;
    let needRefreshProtectionUI = false;
    let selectionRootHasGroup = false;
    let groupedChange = false;
    try {
      await window.PptDriver.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx, { shapeKinds: selectionShapeKinds });
        const sel = driver.selectedShapes();
        const leaves = await driver.loadShapeTree(sel, 'id, width, height, level, adjustments, tags');
        if (!isCurrent()) return;
        selectionRootHasGroup = driver.hasTopLevelGroup(sel);
        let verifiedTags = null;
        for (const sh of leaves) {
          if (!isCurrent()) return;
          if (!driver.isRoundRect(sh)) continue;
          const id = driver.shapeId(sh);
          const currentAdj = await driver.readAdjFraction(sh);
          if (!isCurrent()) return;
          const size = driver.size(sh);
          const minSideCm = Math.min(size.width, size.height) / PT_PER_CM;
          if (!(minSideCm > 0)) continue;
          const shape = selectedShapes.find((s) => s.id === id);
          if (!shape) continue;
          const previous = { width: lockMonitor.lastWidth[id], height: lockMonitor.lastHeight[id],
            adj: lockMonitor.lastAdj[id], candidateAdj: lockMonitor.candidateAdj[id], stableCount: lockMonitor.stableCount[id] };
          const currentCm = currentAdj * minSideCm;
          if (shape.currentCm !== currentCm) needRefreshUI = true;
          Object.assign(shape, { currentCm, width: size.width, height: size.height, minSideCm });
          const sample = { width: size.width, height: size.height, adj: currentAdj };
          let next;
          if (selectionRootHasGroup || driver.shapeLevel(sh) > 0) {
            // Native group transforms must complete without any descendant
            // adjustment, geometry or tag write, including fixed/strict shapes.
            const sizeChanged = previous.adj != null &&
              (Math.abs(size.width - previous.width) > SIZE_EPSILON || Math.abs(size.height - previous.height) > SIZE_EPSILON);
            const adjChanged = previous.adj != null && Math.abs(currentAdj - previous.adj) > ADJ_EPSILON;
            if (sizeChanged || adjChanged) groupedChange = true;
            if (!selectionRootHasGroup && shape.locked && !shape.strictLocked && adjChanged && !sizeChanged) {
              lockMonitor.groupLockUpdates[id] = currentCm;
            }
            next = { ...sample, candidateAdj: currentAdj, stableCount: 0 };
          } else {
            let decision = window.RadiusCore.decideLockMonitorUpdate(shape, sample, previous);
            if (decision.action !== 'none') {
              // Reconfirm subtype and protection tags only when a write is
              // needed. Normal polling does not repeatedly export the slide.
              if (!verifiedTags) {
                selectionShapeKinds = await driver.loadShapeKinds(true);
                verifiedTags = await driver.loadTagsBulk(leaves);
              }
              if (!isCurrent()) return;
              if (!driver.isRoundRect(sh)) { shape.isRoundRect = false; continue; }
              const lock = window.RadiusCore.lockStateFromTags(verifiedTags[id]);
              if (lock.isStrict && !lock.isLocked && shape.legacyStrictNeedsRepair) {
                lock.isLocked = true;
                lock.lockedCm = shape.lockedCm;
              }
              shape.locked = lock.isLocked;
              shape.lockedCm = lock.lockedCm;
              shape.strictLocked = lock.isStrict;
              decision = window.RadiusCore.decideLockMonitorUpdate(shape, sample, previous);
              if (decision.action === 'restore') {
                const result = await window.RadiusCore.reapplyLock(driver, sh, shape.lockedCm);
                if (!result.ok) throw new Error(result.error || result.reason);
                shape.currentCm = result.newCm;
                needRefreshUI = true;
              } else if (decision.action === 'updateLock') {
                const result = await window.RadiusCore.writeLockState(driver, sh, { lockedCm: decision.lockedCm });
                if (!result.ok) throw new Error(result.error);
                shape.lockedCm = decision.lockedCm;
                needRefreshProtectionUI = true;
              }
            }
            next = decision.next;
          }
          lockMonitor.lastWidth[id] = next.width;
          lockMonitor.lastHeight[id] = next.height;
          lockMonitor.lastAdj[id] = next.adj;
          lockMonitor.candidateAdj[id] = next.candidateAdj;
          lockMonitor.stableCount[id] = next.stableCount;
        }
        if (isCurrent()) await driver.sync();
      }, isCurrent);
      if (!isCurrent()) return;
      if (needRefreshUI) renderCurrentRadius();
      if (needRefreshProtectionUI) renderShapeList();
      const rChanges = window.RadiusCore.detectLayoutParentChanges(lockMonitor.lastCm, selectedShapes);
      const sizeChanges = window.RadiusCore.detectLayoutParentSizeChanges(lockMonitor.lastSizeCm, selectedShapes);
      let changed = false, geometry = false;
      for (const change of rChanges) {
        if (change.lastCm != null) changed = true;
        lockMonitor.lastCm[change.parentId] = change.newCm;
      }
      for (const change of sizeChanges) {
        if (change.lastSize != null) { changed = true; geometry = true; }
        lockMonitor.lastSizeCm[change.parentId] = change.newSize;
      }
      if (groupedChange || selectionRootHasGroup && changed) scheduleGroupLayoutSync();
      else if (changed) {
        lockMonitor.parentRDirty = true;
        lockMonitor.parentRSyncGeometry = lockMonitor.parentRSyncGeometry || geometry;
        scheduleParentRSync();
      }
    } catch (error) {
      console.log('[monitorTick] EXCEPTION:', error.message || String(error));
    } finally { monitorInFlight = false; }
  }

  // v1.3.6 修 #6：节流调 syncLayoutChildrenRIfNeeded（避免 10ms tick 频繁开新 PowerPoint.run）
  // 200ms 窗口：用户在拖父 R 角滑块时（≈ 60fps）一次拖完最多触发 5 次，但节流后实际只跑 1 次
  function scheduleParentRSync() {
    if (lockMonitor.parentRSyncTimer) return;  // 已经有 pending 的 timer
    const epoch = selectionEpoch, generation = monitorGeneration;
    lockMonitor.parentRSyncTimer = setTimeout(async () => {
      lockMonitor.parentRSyncTimer = null;
      if (epoch !== selectionEpoch || generation !== monitorGeneration) return;
      if (!lockMonitor.parentRDirty) return;
      lockMonitor.parentRDirty = false;
      const geometry = lockMonitor.parentRSyncGeometry !== false;
      lockMonitor.parentRSyncGeometry = false;
      const request = layoutRequestSerial;
      try {
        const before = selectedShapes.filter((s) => s.layoutRole === 'parent' && s.layoutParams && s.layoutChildIds).length;
        const r = await syncLayoutChildrenRIfNeeded({ geometry, epoch });
        if (before > 0) console.log(`[layout-link] syncLayoutChildrenRIfNeeded done: parents=${before} geometry=${geometry} result=${JSON.stringify(r)}`);
        const failed = r.results.find((result) => !result.ok);
        if (epoch === selectionEpoch && failed) showToast(failed.warn || failed.error || failed.reason);
      } catch (e) {
        console.log('[layout-link] syncLayoutChildrenRIfNeeded error:', e && e.message ? e.message : e);
      } finally {
        if (epoch === selectionEpoch) {
          if (request === layoutRequestSerial) await refreshSelection();
          else startLockMonitor();
        }
      }
    }, PARENT_R_SYNC_DEBOUNCE_MS);
  }

  // group 原生缩放会连同 padding / gutter 一起按比例缩放；UI 中保存的厘米值
  // 因而与画面不再一致。不能直接修改 group.shapes 后代（Mac LTSC 会再次套用
  // transform 导致错位），必须等用户松手后走安全事务：
  // ungroup → 读取新父 box → 完整 applyLayout → regroup。
  function scheduleGroupLayoutSync(opts) {
    if (lockMonitor.groupLayoutSyncTimer) clearTimeout(lockMonitor.groupLayoutSyncTimer);
    const epoch = selectionEpoch, generation = monitorGeneration;
    lockMonitor.groupLayoutSyncTimer = setTimeout(async () => {
      lockMonitor.groupLayoutSyncTimer = null;
      if (epoch !== selectionEpoch || generation !== monitorGeneration || layoutMutationDepth > 0) return;
      const lockedIds = selectedShapes.filter((s) => s.isRoundRect && s.locked).map((s) => s.id);
      const updates = { ...lockMonitor.groupLockUpdates };
      const legacy = Object.fromEntries(selectedShapes.filter((shape) => shape.legacyStrictNeedsRepair)
        .map((shape) => [shape.id, shape.lockedCm]));
      const parents = selectedShapes.filter((s) => s.layoutRole === 'parent');
      const request = layoutRequestSerial;
      stopLockMonitor();
      try {
        if (lockedIds.length) await runMutation((driver, isCurrent) => window.RadiusCore.reapplySelectionLocks(driver, lockedIds, updates, legacy, { isCurrent }), epoch);
        if (epoch !== selectionEpoch) return;
        if (parents.length && !(opts && opts.geometry === false)) {
          const result = await syncLayoutChildrenRIfNeeded({ geometry: true, epoch });
          const failed = result.results.find((r) => !r.ok);
          if (failed) showToast(failed.warn || failed.error || failed.reason);
        }
      } catch (error) {
        console.log('[group settle] EXCEPTION:', error.message || String(error));
      } finally {
        if (epoch === selectionEpoch) {
          if (request === layoutRequestSerial) await refreshSelection();
          else startLockMonitor();
        }
      }
    }, GROUP_LAYOUT_SYNC_DEBOUNCE_MS);
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
      setStatus(i18n.t('statusLabelSelection'), i18n.t('statusShapeCountFmt', { count: 0 }), 'status-warn');
      $('current-radius').textContent = '—';
      $('locked-count').textContent = '—';
    } else {
      const allRound = selectedShapes.every((s) => s.isRoundRect);
      const anyRound = selectedShapes.some((s) => s.isRoundRect);
      const mixed = !allRound && anyRound;
      const ok = allRound;
      setStatus(
        i18n.t('statusLabelSelection'),
        mixed
          ? i18n.t('statusShapeCountMixedFmt', { count: selectedShapes.length })
          : i18n.t('statusShapeCountFmt', { count: selectedShapes.length }),
        ok ? 'status-ok' : 'status-warn'
      );
      // 当前 R 角（圆角矩形的）
      const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
      if (roundShapes.length > 0 && roundShapes[0].currentCm != null) {
        $('current-radius').textContent = i18n.t('statusCurrentRCmFmt', { value: roundShapes[0].currentCm.toFixed(2) });
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
    $('unit-label').textContent = i18n.t(currentUnit === 'cm' ? 'unitCm' : 'unitPercent');
    const hasRound = selectedShapes.some((s) => s.isRoundRect);
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
      list.innerHTML = '<div class="empty-list">' + i18n.t('emptyShapes') + '</div>';
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
        tag = '<span class="shape-warn">' + i18n.t('nonRoundRect') + '</span>';
      }
      const rText = s.currentCm != null ? `${s.currentCm.toFixed(2)}cm` : '—';
      const name = document.createElement('span');
      name.className = 'shape-name';
      name.textContent = s.name || i18n.t('unnamed');
      const radius = document.createElement('span');
      radius.className = 'shape-r';
      radius.textContent = rText;
      row.appendChild(name);
      row.appendChild(radius);
      if (tag) {
        const badge = document.createElement('span');
        badge.innerHTML = tag;
        row.appendChild(badge);
      }
      list.appendChild(row);
    }
  }

  // 轻量更新"当前 R 角"显示：状态卡 #current-radius + 形状列表每行 .shape-r
  // 不会重建 DOM，只改文本节点
  function renderCurrentRadius() {
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    if (roundShapes.length > 0 && roundShapes[0].currentCm != null) {
      $('current-radius').textContent = i18n.t('statusCurrentRCmFmt', { value: roundShapes[0].currentCm.toFixed(2) });
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
      $('lock-label').textContent = i18n.t('lockRadius');
      $('lock-hint').textContent = i18n.t('emptyShapes');
      updateStrictToggle();
      return;
    }
    btn.disabled = false;
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    const allLocked = roundShapes.length > 0 && roundShapes.every((s) => s.locked);
    $('lock-icon').textContent = allLocked ? '🔒' : '🔒';
    $('lock-label').textContent = allLocked ? i18n.t('lockRadiusOff') : i18n.t('lockRadius');
    $('lock-hint').textContent = allLocked
      ? (i18n.getLang() === 'zh' ? `已使用数值固定 R 角 ${roundShapes.length} 个（PPT 内编辑会被反算回固定值）` : `Fix-R-by-value on ${roundShapes.length} shape(s) (in-PPT edits are reversed to fixed value)`)
      : (i18n.getLang() === 'zh' ? '开启后 R 角按厘米值保持，PPT 内编辑会被反算' : 'When on, R is held in cm and in-PPT edits are reversed');
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
      if (hintEl) hintEl.textContent = i18n.t('strictHint');
    } else {
      // 任何时候都能开 strict（开启时自动 lock）
      label.classList.remove('disabled');
      cb.disabled = false;
      const allStrict = roundShapes.every((s) => s.strictLocked);
      cb.checked = allStrict;
      if (hintEl) hintEl.textContent = allStrict
        ? (i18n.getLang() === 'zh' ? '已开启（任何修改都不会改 R 角）' : 'Enabled (all edits rejected)')
        : i18n.t('strictHint');
    }
  }

  // ---------------- 操作 ----------------

  /** 应用 R 角：所有选中的圆角矩形都改成输入的 cm 值 */
  async function onApply() {
    if (selectedShapes.length === 0) {
      showToast(i18n.t('toastSelectRoundRect'));
      return;
    }
    const raw = parseFloat($('radius-input').value);
    if (!Number.isFinite(raw) || raw < 0) {
      showToast(i18n.t('toastInvalidR'));
      return;
    }
    // v1.1 防误触拦截：选区里有任何 strict 锁定 → 全部拒绝
    const strictLocked = selectedShapes.filter((s) => s.isRoundRect && s.strictLocked);
    if (strictLocked.length > 0) {
      showToast(i18n.t('toastStrictBlocksFmt', { count: strictLocked.length }));
      return;
    }
    // 输入值按当前单位换算成 cm
    const cm = valueToCm(raw, currentUnit);
    let updated = 0;
    let failed = 0;
    let lockedSynced = 0; // 计数：locked 子被同步 fixed value 的数量
    const epoch = selectionEpoch;
    let resumeEpoch = epoch;
    stopLockMonitor();
    try {
      const result = await runMutation(async (driver, isCurrent) => {
        const leaves = await driver.loadShapeTree(driver.selectedShapes(), 'id, width, height, adjustments, tags');
        return window.RadiusCore.applyRadiusToSelection(driver, leaves, cm, { isCurrent });
      }, epoch);
      if (epoch !== selectionEpoch || result.reason === 'stale-selection') return;
      if (!result.ok) throw new Error(result.error || result.reason);
      updated = result.applied;
      failed = result.failed;
      lockedSynced = result.lockedSynced;
      if (failed === 0) {
        const displayVal = currentUnit === '%'
          ? `${raw.toFixed(1)}%`
          : `${raw.toFixed(2)} 厘米`;
        const lockHint = lockedSynced > 0
          ? `，${lockedSynced} 个使用数值固定 R 角已同步更新`
          : '';
        showToast(i18n.t('toastRUpdatedFmt', { count: updated, value: displayVal, lockHint: lockHint }));
        if (updated > 0) {
          // 写到内存 + 渲染
          const newHistory = pushHistory(raw, currentUnit);
          renderHistory(newHistory);
        }
      } else {
        showToast(i18n.t('toastPartialSuccessFmt', { updated: updated, failed: failed }));
      }
      resumeEpoch = await refreshSelection();
      if (resumeEpoch !== selectionEpoch) return;
      // v1.2: 选区里有 layout 父 → 同步子 R 角（联动）
      await syncLayoutChildrenRIfNeeded({ geometry: false, epoch: resumeEpoch });
    } catch (err) {
      if (resumeEpoch === selectionEpoch) {
        showToast(i18n.t('toastApplyFailedFmt', { error: err.message || err }));
        resumeEpoch = await refreshSelection();
      }
    } finally {
      // 写完恢复 monitor（stopLockMonitor 已清空 last 状态，startLockMonitor 从干净开始）
      if (resumeEpoch === selectionEpoch && selectedShapes.length > 0) startLockMonitor();
    }
  }

  /** 使用数值固定 R 角 开启/关闭：用 shape.tags 存固定值，跟 .pptx 文件走 */
  async function onToggleLock() {
    if (selectedShapes.length === 0) {
      showToast(i18n.t('toastSelectRoundRect'));
      return;
    }
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    if (roundShapes.length === 0) {
      showToast(i18n.t('toastNotRoundRect'));
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
        const inputCm = Number.isFinite(inputVal) && inputVal >= 0
          ? valueToCm(inputVal, currentUnit)
          : s.currentCm;
        if (Number.isFinite(inputCm) && inputCm >= 0) locks[s.id] = inputCm;
        // 之前已经开启过 strict（防误触），再次"开启使用数值固定 R 角"时保留 strict 状态
        if (s.strictLocked) strict[s.id] = true;
      }
      touched++;
    }
    const r = await saveLocksViaTags(locks, strict);
    if (!r.ok) {
      showToast(i18n.t('toastApplyFailedFmt', { error: r.error?.message || r.error || r.reason }));
      await refreshSelection();
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
      showToast(i18n.t('toastSelectRoundRect'));
      return;
    }
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    if (roundShapes.length === 0) {
      showToast(i18n.t('toastNotRoundRect'));
      return;
    }
    if (newValue) {
      // 开启：自动用当前 R 角作 fixed value（如果还没 lock）
      const locks = {}; // id -> cm
      for (const s of roundShapes) {
        if (s.locked && s.lockedCm >= 0) {
          // 已 lock：保留原 fixed value
          locks[s.id] = s.lockedCm;
        } else if (s.currentCm != null && s.currentCm >= 0) {
          // 没 lock：用当前 R 角作 fixed value
          locks[s.id] = s.currentCm;
        } else {
          // 未知 R 角不能设 fixed value；0 是有效值
          showToast(i18n.t('toastStrictCannotEnableFmt', { name: s.name || i18n.t('unnamed') }));
          return;
        }
      }
      const strict = {};
      for (const s of roundShapes) strict[s.id] = true;
      const r = await saveLocksViaTags(locks, strict);
      if (!r.ok) {
        showToast(i18n.t('toastApplyFailedFmt', { error: r.error?.message || r.error || r.reason }));
        await refreshSelection();
        return;
      }
      showToast(i18n.t('toastStrictEnabledFmt', { count: roundShapes.length }));
    } else {
      // 关闭：只删 strict tag（不动 lock tag）
      const result = await saveLocksViaTags({}, Object.fromEntries(roundShapes.map((s) => [s.id, false])));
      if (!result.ok) {
        showToast(i18n.t('toastApplyFailedFmt', { error: result.error || result.reason }));
        await refreshSelection();
        return;
      }
      showToast(i18n.t('toastStrictDisabledFmt', { count: roundShapes.length, keepHint: roundShapes.some((s) => s.locked) ? i18n.t('keepLockHint') : '' }));
    }
    await refreshSelection();
  }

  /** 重新应用锁定：按当前形状大小反算 adj */
  async function onReapply() {
    const locked = selectedShapes.filter((s) => s.locked);
    if (locked.length === 0) {
      showToast(i18n.t('toastNoLockedRects'));
      return;
    }
    let applied = 0;
    let failed = 0;
    const epoch = selectionEpoch;
    try {
      const result = await runMutation((driver, isCurrent) => window.RadiusCore.reapplySelectionLocks(
        driver, locked.map((shape) => shape.id), null, null, { isCurrent }
      ), epoch);
      if (epoch !== selectionEpoch || result.reason === 'stale-selection') return;
      applied = result.applied;
      failed = result.failed;
      showToast(i18n.t('toastReappliedFmt', { count: applied, failed: failed > 0 ? i18n.t('failedStrFmt', { count: failed }) : '' }));
      await refreshSelection();
    } catch (err) {
      if (epoch === selectionEpoch) {
        showToast(i18n.t('toastApplyFailedFmt', { error: err.message || err }));
        await refreshSelection();
      }
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
    $('unit-label').textContent = i18n.t(newUnit === 'cm' ? 'unitCm' : 'unitPercent');
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
      empty.textContent = i18n.t('emptyPreset');
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
          showToast(i18n.t('toastRenamedFmt', { name: v }));
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
      valInput.title = i18n.t('presetEditTitleFmt', { unit: i18n.t(p.unit === '%' ? 'unitPercent' : 'unitCm') });
      valInput.addEventListener('change', () => {
        const v = parseFloat(valInput.value);
        if (!Number.isFinite(v) || v < 0) {
          // 还原成原值
          valInput.value = p.unit === '%'
            ? (Number.isInteger(p.value) ? String(p.value) : p.value.toFixed(1))
            : p.value.toFixed(2);
          showToast(i18n.t('toastInvalidR'));
          return;
        }
        if (p.unit === '%' && v > 50) {
          valInput.value = String(p.value);
          showToast(i18n.t('toastPercentRange'));
          return;
        }
        p.value = v;
        // 更新应用按钮的 title 提示
        applyBtn.title = `应用 ${p.name} = ${formatPresetValue(p.value, p.unit)} 到当前选区`;
        showToast(i18n.t('toastPresetUpdatedFmt', { name: p.name, value: formatPresetValue(v, p.unit) }));
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
      showToast(i18n.t('toastPresetsFullFmt', { max: MAX_PRESETS }));
      return;
    }
    // 优先用当前选中的圆角矩形的 R 角；没有就退回输入框
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect && s.currentCm != null && s.currentCm >= 0);
    let raw, unit;
    if (roundShapes.length > 0) {
      const src = roundShapes[0];
      unit = currentUnit;
      raw = cmToValue(src.currentCm, unit);
    } else {
      raw = parseFloat($('radius-input').value);
      unit = currentUnit;
      if (!Number.isFinite(raw) || raw < 0) {
        showToast(i18n.t('toastNoRValue'));
        $('radius-input').focus();
        return;
      }
    }
    // 单位校验
    if (unit === '%' && (raw < 0 || raw > 50)) {
      showToast(i18n.t('toastPercentRange'));
      return;
    }
    // 默认命名：「预设 N」
    const name = `预设 ${userPresets.length + 1}`;
    const p = { id: nextPresetId(), name, value: raw, unit };
    userPresets = [p, ...userPresets].slice(0, MAX_PRESETS);
    renderPresets(userPresets);
    showToast(i18n.t('toastPresetSavedFmt', { name: name, value: formatPresetValue(raw, unit) }));
  }

  function applyPreset(preset) {
    if (selectedShapes.length === 0) {
      showToast(i18n.t('toastSelectRoundRect'));
      return;
    }
    const roundCount = selectedShapes.filter((s) => s.isRoundRect).length;
    if (roundCount === 0) {
      showToast(i18n.t('toastNotRoundRect'));
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
      showToast(i18n.t('toastPresetDeleted'));
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
      badge.textContent = i18n.t('pipetteStateIdle');
      hint.classList.remove('has-source');
      hint.textContent = i18n.t('hintPipette');
      icon.textContent = '🖌️';
      label.textContent = i18n.t('pipettePickR');
      btn.classList.remove('state-sourcing', 'state-brushing');
    } else if (newState === 'sourcing') {
      badge.className = 'pipette-state-badge sourcing';
      badge.textContent = i18n.getLang() === 'zh' ? '吸取中…' : 'Sourcing…';
      hint.classList.add('has-source');
      hint.textContent = i18n.t('pipetteHintSourcing');
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
        showToast(i18n.t('toastNotRoundRectPickup'));
      }
      setPipetteState('sourcing');
      showToast(i18n.t('toastEnterPipette'));
    } else {
      // 任意非 idle 状态点击按钮都退出
      setPipetteState('idle');
      pipetteSource = null;
      showToast(i18n.t('toastPipetteClosed'));
    }
  }

  // 从选区第一个 roundRect 吸取（自己读选区，不依赖 selectedShapes 内存）
  // Mac LTSC task pane 必加：get(0) 存到变量 → sync → 读 value
  // （不能 load 之后再新调 get(0).value，那时 value 还没填上，会报"尚未加载"）
  // v1.3.6 迁移：PowerPoint.run 部分走 radius-core.pickupFromSelection
  async function pickupFromSelection() {
    const epoch = selectionEpoch;
    let picked = null;
    try {
      await window.PptDriver.run(async (ctx) => {
        const driver = window.PptDriver.createDriver(ctx);
        const sel = driver.selectedShapes();
        const selLeaves = await driver.loadShapeTree(
          sel,
          'id, name, width, height, adjustments, tags'
        );
        const r = await window.RadiusCore.pickupFromSelection(driver, selLeaves);
        if (r) {
          picked = r;
        }
      });
    } catch (e) {
      showToast(i18n.t('toastPickupFailedFmt', { error: e.message || e }));
      return;
    }
    if (epoch !== selectionEpoch) return;
    if (!picked) {
      showToast(i18n.t('toastNoRoundRectInSelection'));
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
    showToast(i18n.t('toastPickedFmt', { name: pipetteSource.sourceShapeName, value: formatPresetValue(value, currentUnit), strictHint: strictHint }));
    // 顺便把 selectedShapes 内存刷新一下（让状态卡同步显示源形状）
    await refreshSelection();
  }

  // 把 pipetteSource 应用到选区里所有 roundRect
  //
  // v1.3.6 迁移：所有步骤走 radius-core.applyPickedToSelection（单 PowerPoint.run 完成）
  // 已保护目标总是拦截；同步选项只复制源的开启状态，不自动解除目标保护。
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
    if (selectedShapes.some((shape) => shape.isRoundRect && shape.strictLocked)) {
      showToast(i18n.t('toastBrushBlocked'));
      return;
    }
    const epoch = selectionEpoch;
    const source = { cm: pipetteSource.cm, sourceStrict: pipetteSource.sourceStrict };
    const syncStrict = pipetteSyncStrict;
    let result;
    try {
      result = await runMutation(async (driver, isCurrent) => {
        const leaves = await driver.loadShapeTree(driver.selectedShapes(), 'id, width, height, adjustments, tags');
        if (!isCurrent()) return { ok: false, reason: 'stale-selection' };
        return window.RadiusCore.applyPickedToSelection(driver, leaves, source, { syncStrict, isCurrent });
      }, epoch);
      if (epoch !== selectionEpoch || result.reason === 'stale-selection') return;
    } catch (error) {
      showToast(i18n.t('toastBrushFailedFmt', { error: error.message || error }));
      if (epoch === selectionEpoch) await refreshSelection();
      return;
    }

    if (!result || !result.ok) {
      if (result && result.rejectReason === 'strict') {
        showToast(i18n.t('toastBrushBlocked'));
      } else if (result && result.applied === 0) {
        showToast(i18n.t('toastBrushNoSelection'));
      } else if (result) {
        showToast(i18n.t('toastBrushFailedFmt', { error: result.error || (i18n.getLang() === 'zh' ? '未知错误' : 'unknown error') }));
      }
      if (selectedShapes.length > 0) startLockMonitor();
      return;
    }

    // toast
    const lockHint = '';  // radius-core.applyPickedToSelection 已经处理 lock 同步
    const strictHint = result.strictAdded > 0 ? `，${result.strictAdded} 个开启防误触` : '';
    showToast(i18n.t('toastBrushAppliedFmt', { count: result.applied, failed: result.failed > 0 ? i18n.t('failedStrFmt', { count: result.failed }) : '', lockHint: lockHint, strictHint: strictHint }));
    const refreshedEpoch = await refreshSelection();
    if (refreshedEpoch !== selectionEpoch) return;
    // v1.2: layout 父被刷 R 角 → 同步子 R 角
    await syncLayoutChildrenRIfNeeded({ geometry: false, epoch: refreshedEpoch });
    if (refreshedEpoch === selectionEpoch && selectedShapes.length > 0) startLockMonitor();
  }

  // DocumentSelectionChanged 分发：idle → refreshSelection；sourcing → pickup；brushing → apply
  async function onSelectionChangedForPipette() {
    if (pipetteState === 'sourcing') {
      await pickupFromSelection();
    } else if (pipetteState === 'brushing') {
      // Protection in memory must describe the new target, not the source or
      // previous target. The live tag preflight remains the second defense.
      const refreshedEpoch = await refreshSelection();
      if (refreshedEpoch !== selectionEpoch || pipetteState !== 'brushing') return;
      await applyPipetteToSelection();
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
    // v1.2.14：边距/间距 锁链联动按钮（Photoshop 风格，跨两行垂直居中）
    const pgLinkBtn = $('layout-pg-link');
    if (pgLinkBtn) {
      pgLinkBtn.addEventListener('click', () => {
        linkPG = !linkPG;
        renderLayoutPanel();  // 同步 params.gutter = params.padding（linkPG=true 时）+ 禁用 gutter UI
        if (linkPG && currentLayout) {
          // v1.2.15：立即 apply 形状，不等 200ms 节流
          // 原因：用户报 bug —— 锁链激活时如果之前 gutter 已经被改过（跟 padding 不同），
          //   scheduleLayoutApply 200ms 内可能被外层 input 事件覆盖，导致形状不按 padding 写
          // 锁链切换是"明确意图"，不走节流，立即 apply
          // v1.2.15 改：这里只是"初始 snap 写入"——之后 gutter 控件 disabled，用户碰不到，不需要再 sync
          if (layoutApplyTimer) {
            clearTimeout(layoutApplyTimer);
            layoutApplyTimer = null;
            layoutApplyPending = false;
            layoutApplyPendingOpts = null;
          }
          applyLayoutFromUI({ writeParentTag: true, syncR: true });
        }
      });
    }
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
        console.log('[layout-ui] linkRMode change → R-only apply mode=', r.value);
        scheduleLayoutApply({
          writeParentTag: true,
          syncR: true,
          writeGeometry: false,
        });
      });
    });
  }

  async function handleSelectionChanged() {
    if (layoutMutationDepth > 0 && activeMutationDriver && activeMutationDriver.structuralChanged) return;
    if (Date.now() < layoutSelectionIgnoreUntil && ignoredRestoredSelection) {
      // Distinguish a delayed internal regroup event from a real user switch.
      // Stop the old monitor immediately while this serialized read is pending.
      stopLockMonitor();
      if (selectionEventProbeInFlight) return;
      selectionEventProbeInFlight = true;
      try {
        const identity = await window.PptDriver.run((ctx) => window.PptDriver.createDriver(ctx).loadSelectionIdentity());
        const expected = ignoredRestoredSelection;
        if (expected && identity.slideId === expected.slideId &&
            identity.shapeIds.length === expected.shapeIds.length &&
            identity.shapeIds.every((id) => expected.shapeIds.includes(id))) {
          startLockMonitor();
          return;
        }
      } catch (error) { console.log('[selection identity] EXCEPTION:', error.message || error); }
      finally { selectionEventProbeInFlight = false; }
    }
    selectionEpoch++;
    stopLockMonitor();
    if (pipetteState === 'idle') await refreshSelection();
    else await onSelectionChangedForPipette();
  }

  // ---------------- 初始化 ----------------

  window.PptDriver.onReady(() => {
    // Apply i18n to any [data-i18n] / [data-i18n-*] attributes (HTML inline script
    // already did this on DOMContentLoaded, but call again in case Office is slow
    // and dynamic textContent / placeholder updates need to be re-translated).
    if (window.i18n && window.i18n.applyAll) window.i18n.applyAll();
    bindEvents();
    renderPresets(userPresets); // 渲染空预设库
    refreshSelection();
    // 选区变化：分发到 pipette（sourcing → pickup；brushing → apply）或 refreshSelection（idle）
    window.PptDriver.onSelectionChanged(handleSelectionChanged);
  });
})();
