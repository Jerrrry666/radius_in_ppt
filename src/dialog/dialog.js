/*
 * dialog.js — R 角调整 v1.1（task pane，纯 Office.js）
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

  // 本次 session 内用户主动应用过的 R 角值（纯内存）
  let userHistory = [];

  // 当前选中的形状（refreshSelection 填充）
  let selectedShapes = [];

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
  };

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
  function loadLocksViaTags() {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const sel = ctx.presentation.getSelectedShapes();
          sel.load('items/id');
          await ctx.sync();
          const locks = {};   // id -> cm（使用数值固定 R 角）
          const strict = {};  // id -> true（防误触开关）
          for (const sh of sel.items) {
            // 读使用数值固定 R 角
            try {
              const lockTag = sh.tags.getItem(LOCK_TAG_KEY);
              lockTag.load('value');
              await ctx.sync();
              const cm = parseFloat(lockTag.value);
              if (Number.isFinite(cm) && cm > 0) {
                locks[sh.id] = cm;
              }
            } catch (_) { /* lock tag 不存在 */ }
            // 读防误触标记
            try {
              const strictTag = sh.tags.getItem(LOCK_STRICT_TAG_KEY);
              strictTag.load('value');
              await ctx.sync();
              if (strictTag.value === '1') {
                strict[sh.id] = true;
              }
            } catch (_) { /* strict tag 不存在 */ }
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
          const sel = ctx.presentation.getSelectedShapes();
          sel.load('items/id');
          await ctx.sync();
          for (const sh of sel.items) {
            const cm = locks[sh.id];
            const isStrict = !!(strictMap && strictMap[sh.id]);
            try {
              if (cm == null) {
                sh.tags.delete(LOCK_TAG_KEY);
              } else {
                sh.tags.add(LOCK_TAG_KEY, String(cm));
              }
            } catch (_) { /* 单个 shape 写失败不影响其他 */ }
            try {
              if (isStrict) {
                sh.tags.add(LOCK_STRICT_TAG_KEY, '1');
              } else {
                sh.tags.delete(LOCK_STRICT_TAG_KEY);
              }
            } catch (_) {}
          }
          await ctx.sync();
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
  function updateLockTagForShape(shapeId, cm, isStrict) {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          // 用 getSelectedShapes + filter 找目标 shape
          const sel = ctx.presentation.getSelectedShapes();
          sel.load('items/id');
          await ctx.sync();
          for (const sh of sel.items) {
            if (sh.id !== shapeId) continue;
            // cm：number 写 / null 删 / undefined 跳过
            if (cm !== undefined) {
              try {
                if (cm == null) {
                  sh.tags.delete(LOCK_TAG_KEY);
                } else {
                  sh.tags.add(LOCK_TAG_KEY, String(cm));
                }
              } catch (_) {}
            }
            // isStrict：true 开 / false 关 / null/undefined 跳过
            if (isStrict === true) {
              try { sh.tags.add(LOCK_STRICT_TAG_KEY, '1'); } catch (_) {}
            } else if (isStrict === false) {
              try { sh.tags.delete(LOCK_STRICT_TAG_KEY); } catch (_) {}
            }
            break;
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
    // 纯内存：去重 + 移到最前 + 限 5 条
    const filtered = userHistory.filter((h) => !(h.value === value && h.unit === unit));
    filtered.unshift({ value, unit, ts: Date.now() });
    userHistory = filtered.slice(0, MAX_HISTORY);
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
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/width, items/height, items/adjustments');
        await ctx.sync();
        for (const sh of sel.items) {
          const adjCount = sh.adjustments.count;
          if (adjCount === 0) continue; // 不是 roundRect
          const adjResult = sh.adjustments.get(0);
          await ctx.sync();
          let currentAdj = null;
          try { currentAdj = adjResult.value; } catch (_) {}
          if (currentAdj == null) continue;
          const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
          if (minSideCm <= 0) continue;
          const currentCm = currentAdj * minSideCm;

          // 找对应 selectedShape
          const ss = selectedShapes.find((x) => x.id === sh.id);
          if (!ss) continue;
          // 1) 所有 roundRect 都更新内存 currentCm（实时显示 R 角）
          //    任何变化都标 dirty（浮点抖动 < 0.01 cm 的显示精度，看不出来）
          const oldCm = ss.currentCm;
          ss.currentCm = currentCm;
          ss.width = sh.width;
          ss.height = sh.height;
          if (oldCm == null || currentCm !== oldCm) {
            needRefreshUI = true;
          }
          // 2) 只对 locked shape 做反算/更新固定值
          if (!ss.locked) continue;

          const targetCm = Math.min(ss.lockedCm, minSideCm / 2);
          const targetAdj = (targetCm / minSideCm) * ADJ_SCALE;
          const lastW = lockMonitor.lastWidth[sh.id];
          const lastH = lockMonitor.lastHeight[sh.id];
          const lastA = lockMonitor.lastAdj[sh.id];
          // 第一轮（lastA = null）：只记录初始状态，不做反算
          // （否则会触发"idle 兜底"，把用户第一次拖 R 角的值当异常反算回去）
          if (lastA == null) {
            lockMonitor.lastWidth[sh.id] = sh.width;
            lockMonitor.lastHeight[sh.id] = sh.height;
            lockMonitor.lastAdj[sh.id] = currentAdj;
            lockMonitor.stableCount[sh.id] = 0;
            continue;
          }
          const wChanged = Math.abs(sh.width - lastW) > SIZE_EPSILON;
          const hChanged = Math.abs(sh.height - lastH) > SIZE_EPSILON;
          const aChanged = Math.abs(currentAdj - lastA) > ADJ_EPSILON;
          const sizeChanged = wChanged || hChanged;  // 任意一边变了都算"调尺寸"

          if (sizeChanged) {
            // 拖尺寸手柄（任意边 / 角）：立刻反算回固定值
            if (Math.abs(currentAdj - targetAdj) > ADJ_EPSILON) {
              sh.adjustments.set(0, targetAdj);
              recomputedIds.push(sh.id);
            }
            lockMonitor.lastAdj[sh.id] = targetAdj;
            lockMonitor.stableCount[sh.id] = 0;
          } else if (aChanged) {
            // 拖 R 角黄色滑块：等稳定后视作主动改值
            lockMonitor.stableCount[sh.id] = (lockMonitor.stableCount[sh.id] || 0) + 1;
            if (lockMonitor.stableCount[sh.id] >= LOCK_STABLE_THRESHOLD) {
              if (ss.strictLocked) {
                // 防误触：反算回去
                sh.adjustments.set(0, targetAdj);
                lockMonitor.lastAdj[sh.id] = targetAdj;
                recomputedIds.push(sh.id);
              } else {
                // 仅使用数值固定 R 角：把当前 adj 提升为新的固定值
                const newCm = currentAdj * minSideCm;
                await updateLockTagForShape(sh.id, newCm, undefined);
                ss.lockedCm = newCm;
                lockMonitor.lastAdj[sh.id] = currentAdj;
                updatedLockIds.push(sh.id);
              }
              lockMonitor.stableCount[sh.id] = 0;
            }
          } else {
            // 都没变：idle，检查兜底（adj 跟 target 不一致但 size 和 adj 都没"主动变化"）
            if (Math.abs(currentAdj - targetAdj) > ADJ_EPSILON) {
              // 极端 race 兜底：写回
              sh.adjustments.set(0, targetAdj);
              lockMonitor.lastAdj[sh.id] = targetAdj;
              recomputedIds.push(sh.id);
            }
            lockMonitor.stableCount[sh.id] = 0;
          }
          lockMonitor.lastWidth[sh.id] = sh.width;
          lockMonitor.lastHeight[sh.id] = sh.height;
        }
        await ctx.sync();
      });
      // 轻量更新 UI（只改文本节点，不重建 DOM）
      if (needRefreshUI) renderCurrentRadius();
      if (recomputedIds.length > 0) {
        showToast(`🔒 使用数值固定 R 角：反算了 ${recomputedIds.length} 个被改的 R 角`);
      } else if (updatedLockIds.length > 0) {
        showToast(`🪄 使用数值固定 R 角已跟随 R 角滑块更新（${updatedLockIds.length} 个）`);
      }
    } catch (e) {
      console.warn('lock monitor error:', e);
    }
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
    // 防误触拦截：如果选区里有任何 strict 锁定 → 全部拒绝
    const strictLocked = selectedShapes.filter((s) => s.isRoundRect && s.strictLocked);
    if (strictLocked.length > 0) {
      showToast(`🔒 防误触已开启（${strictLocked.length} 个），不能改 R 角。先关掉防误触或解锁。`);
      return;
    }
    // 输入值按当前单位换算成 cm
    const cm = valueToCm(raw, currentUnit);
    let updated = 0;
    let failed = 0;
    // 收集需要更新固定值的目标（已使用数值固定 R 角但未开启防误触的 roundRect）
    const lockedTargets = []; // [{ id, newCm }]
    // 写之前停 monitor（避免 race：新写的 adj 被 monitor 当作"用户拖 R 角"再触发一次）
    stopLockMonitor();
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/width, items/height, items/adjustments');
        await ctx.sync();
        for (const sh of sel.items) {
          const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
          if (minSideCm <= 0) { failed++; continue; }
          const targetCm = Math.min(cm, minSideCm / 2);
          // ADJ_SCALE=1 → newAdj 是 0~0.5 的小数比例，不能 round 到整数
          const newAdj = (targetCm / minSideCm) * ADJ_SCALE;
          if (!Number.isFinite(newAdj)) { failed++; continue; }
          try {
            sh.adjustments.set(0, newAdj);
            updated++;
            // 检查对应 selectedShape 是否已使用数值固定 R 角（且未开启防误触），是的话收集起来稍后更新固定值
            const ss = selectedShapes.find((x) => x.id === sh.id);
            if (ss && ss.locked && !ss.strictLocked) {
              lockedTargets.push({ id: sh.id, newCm: targetCm });
            }
          } catch (_) {
            // 这个形状可能不是 roundRect，set 失败
            failed++;
          }
        }
        await ctx.sync();
      });
      // 对已使用数值固定 R 角的 shape 更新固定值（这样后续 lock monitor 反算用的是新值）
      for (const t of lockedTargets) {
        await updateLockTagForShape(t.id, t.newCm, undefined); // undefined = strict 状态不变
      }
      if (failed === 0) {
        const displayVal = currentUnit === '%'
          ? `${raw.toFixed(1)}%`
          : `${raw.toFixed(2)} 厘米`;
        const lockHint = lockedTargets.length > 0
          ? `，${lockedTargets.length} 个使用数值固定 R 角已同步更新`
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
        // 关闭使用数值固定 R 角：不写 lock tag；strict 标记也要清掉
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
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/width, items/height');
        await ctx.sync();
        for (const sh of sel.items) {
          const target = locked.find((x) => x.id === sh.id);
          if (!target) continue;
          const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
          if (minSideCm <= 0) { failed++; continue; }
          const targetCm = Math.min(target.lockedCm, minSideCm / 2);
          const newAdj = (targetCm / minSideCm) * ADJ_SCALE;
          if (!Number.isFinite(newAdj)) { failed++; continue; }
          try {
            sh.adjustments.set(0, newAdj);
            applied++;
          } catch (_) { failed++; }
        }
        await ctx.sync();
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
  async function pickupFromSelection() {
    let picked = null;
    let sourceStrict = false;
    let sourceId = null;
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/name, items/width, items/height, items/adjustments');
        await ctx.sync();
        if (sel.items.length === 0) return;
        for (const sh of sel.items) {
          const adjCount = sh.adjustments.count;
          if (adjCount > 0) {
            const adjResult = sh.adjustments.get(0); // 先存变量
            await ctx.sync();                        // 再 sync（让 value 填上）
            const v = adjResult.value;               // 用之前的变量读
            if (Number.isFinite(v)) {
              // v = 0 也允许（R 角 0 = 直角矩形，apply 时直接把它变直角）
              const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
              picked = {
                id: sh.id,
                name: sh.name || '(未命名)',
                cm: v * minSideCm,
              };
              sourceId = sh.id;
              // 顺手读源形状的「防误触」标记
              try {
                const strictTag = sh.tags.getItem(LOCK_STRICT_TAG_KEY);
                strictTag.load('value');
                await ctx.sync();
                sourceStrict = strictTag.value === '1';
              } catch (_) { /* 没 strict tag */ }
              break;
            }
          }
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
      sourceStrict,  // 源形状的防误触状态（供「刷防误触状态」选项使用）
      sourceId,
    };
    setPipetteState('brushing');
    const strictHint = pipetteSyncStrict && sourceStrict ? '（含防误触）' : '';
    showToast(`🪣 已吸取「${pipetteSource.sourceShapeName}」= ${formatPresetValue(value, currentUnit)} — 选中目标形状自动应用${strictHint}`);
    // 顺便把 selectedShapes 内存刷新一下（让状态卡同步显示源形状）
    refreshSelection();
  }

  // 把 pipetteSource 应用到选区里所有 roundRect
  //
  // 应用顺序（明确两步）：
  //   步骤 0 — 拦截：选区里有任何目标启用了防误触 → 整个样式刷拒绝（不论是否勾选「刷防误触状态」）
  //   步骤 1 — 第一次"应用"：刷 R 角到所有目标 + 同步 fixed value（已 lock 的目标）
  //   步骤 2 — 第二次"应用"：根据【刷防误触状态】勾选决定是否把源 strict 状态写到目标
  //   步骤 3 — toast + refreshSelection
  //
  // 关键：拦截和写值都用「自己重新读选区（含 strict / lock tag）」，**不依赖 selectedShapes 内存**。
  // 原因：brushing 状态下 DocumentSelectionChanged 不会调 refreshSelection，
  //       selectedShapes 内存会残留之前选中的形状（包括他们的 strict 状态），
  //       导致拦截误命中（user 报：选了未勾防误触的目标却报错"目标启用了防误触"）。
  async function applyPipetteToSelection() {
    if (!pipetteSource) {
      setPipetteState('idle');
      return;
    }

    // ============ 步骤 0：自己读选区（含 strict / lock tag），做最新状态判断 ============
    const liveShapes = []; // [{ id, isRoundRect, isStrict, isLocked, lockedCm, minSideCm }]
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/width, items/height, items/adjustments');
        await ctx.sync();
        for (const sh of sel.items) {
          const adjCount = sh.adjustments.count;
          const isRoundRect = (typeof adjCount === 'number' ? adjCount : 0) > 0;
          let isStrict = false;
          let isLocked = false;
          let lockedCm = null;
          if (isRoundRect) {
            try {
              const strictTag = sh.tags.getItem(LOCK_STRICT_TAG_KEY);
              strictTag.load('value');
              await ctx.sync();
              isStrict = strictTag.value === '1';
            } catch (_) { /* 没 strict tag */ }
            try {
              const lockTag = sh.tags.getItem(LOCK_TAG_KEY);
              lockTag.load('value');
              await ctx.sync();
              const cm = parseFloat(lockTag.value);
              if (Number.isFinite(cm) && cm > 0) {
                isLocked = true;
                lockedCm = cm;
              }
            } catch (_) { /* 没 lock tag */ }
          }
          liveShapes.push({
            id: sh.id,
            isRoundRect,
            isStrict,
            isLocked,
            lockedCm,
            minSideCm: Math.min(sh.width, sh.height) / PT_PER_CM,
          });
        }
      });
    } catch (err) {
      showToast('读选区失败：' + (err.message || err));
      return;
    }

    // 拦截：用 liveShapes 不用 selectedShapes 内存
    const antiStrictTargets = liveShapes.filter((s) => s.isRoundRect && s.isStrict);
    if (antiStrictTargets.length > 0) {
      showToast(`🔒 ${antiStrictTargets.length} 个目标启用了防误触，样式刷不生效。先关掉目标的防误触或解锁。`);
      return;
    }

    stopLockMonitor();

    // ============ 步骤 1：第一次"应用" — 刷 R 角 + 同步 fixed value ============
    let applied = 0;
    let failed = 0;
    const lockedTargets = []; // [{ id, newCm }] — 已 lock 的目标，需要同步 fixed value
    const strictTargets = []; // [{ id, sourceStrict }] — 步骤 2 要刷 strict 的目标（仅在「刷防误触状态」勾选时收集）
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/adjustments');
        await ctx.sync();
        for (const sh of sel.items) {
          const ls = liveShapes.find((x) => x.id === sh.id);
          if (!ls || !ls.isRoundRect) continue;
          if (ls.minSideCm <= 0) { failed++; continue; }
          const targetCm = Math.min(pipetteSource.cm, ls.minSideCm / 2);
          const newAdj = (targetCm / ls.minSideCm) * ADJ_SCALE;
          if (!Number.isFinite(newAdj)) { failed++; continue; }
          try {
            // 写 R 角
            sh.adjustments.set(0, newAdj);
            applied++;
            // 已 lock 的目标：步骤 1 内同步 fixed value（避免 monitor 按旧 lockedCm 反算）
            if (ls.isLocked) {
              lockedTargets.push({ id: sh.id, newCm: targetCm });
            }
            // 勾选了「刷防误触状态」+ 是 roundRect：记录到步骤 2 处理
            if (pipetteSyncStrict) {
              strictTargets.push({ id: sh.id, sourceStrict: pipetteSource.sourceStrict });
            }
          } catch (_) { failed++; }
        }
        await ctx.sync();
      });
      // 同步 fixed value（不动 strict）
      for (const t of lockedTargets) {
        await updateLockTagForShape(t.id, t.newCm, undefined);
      }
    } catch (err) {
      showToast('样式刷 R 值失败：' + (err.message || err));
      if (selectedShapes.length > 0) startLockMonitor();
      return;
    }

    // ============ 步骤 2：第二次"应用" — 根据【刷防误触状态】勾选决定是否刷入防误触状态 ============
    if (pipetteSyncStrict) {
      for (const t of strictTargets) {
        await updateLockTagForShape(t.id, undefined, t.sourceStrict);
      }
    }

    // ============ 步骤 3：toast + refreshSelection（让 task pane UI 跟 PPT 同步） ============
    if (applied === 0 && failed === 0) {
      showToast('选区为空，先在 PPT 里选 1 个圆角矩形');
    } else {
      const lockHint = lockedTargets.length > 0
        ? `，${lockedTargets.length} 个使用数值固定 R 角已同步更新`
        : '';
      const strictHint = strictTargets.length > 0
        ? `，${strictTargets.length} 个防误触状态已同步`
        : '';
      showToast(`🪣 样式刷应用了 ${applied} 个圆角矩形${failed > 0 ? `，${failed} 个失败` : ''}${lockHint}${strictHint}`);
    }
    await refreshSelection();
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
