/*
 * dialog.js — R 角调整 v1.0（task pane，纯 Office.js）
 *
 * UI 形式：PowerPoint 侧边栏（task pane），ribbon 上点按钮展开。
 *
 * 工作流：
 *   1. 打开 task pane → getSelectedShapes() → 显示选中的圆角矩形
 *   2. 用户输入 R 角（cm 或 %）→ 「应用 R 角」→ adjustments.set(0, newVal)
 *   3. 「锁定 R 角」→ 写 shape.tags（OOXML <p:tagLst>），跟 .pptx 文件走
 *   4. history 槽位 = 本次 session 内用户主动应用过的 R 角（纯内存）
 *
 * Mac LTSC (Office 2021, build 16.111) 实测要点：
 *   - `customProperties` / `customXmlParts` 在 task pane 都不可用 → 锁用 shape.tags
 *   - `adjustments.get(0)` 返回 ClientResult 代理，直接 .value 读（不要 .load）
 *   - `shape.adjustments.get(0).value` 是 0~1 比例（不是 OOXML 0~50000）
 *   - Office.js PowerPoint 没有 shape change 事件 → lock 自动重应用靠 setInterval 轮询
 *
 * 监听 PPT 选区变化：DocumentSelectionChanged → 自动 refresh
 * 多页 PPT：getSelectedShapes() 只返回当前页选中的形状，切页后选区变化 → 自动 refresh
 */

(function () {
  const $ = (id) => document.getElementById(id);
  const PT_PER_CM = 28.3464567;        // 1 cm = 28.3464567 pt
  // Mac LTSC: adjustments.get(0).value 是 0~1 比例（占短边），不是 OOXML 0~50000
  const ADJ_SCALE = 1;
  const MAX_HISTORY = 5;
  const LOCK_TAG_KEY = 'radiusLock_v1';

  // 本次 session 内用户主动应用过的 R 角值（纯内存）
  let userHistory = [];

  // 当前选中的形状（refreshSelection 填充）
  let selectedShapes = [];

  // 当前输入单位：'cm' | '%'
  let currentUnit = 'cm';

  // lock monitor 状态：选区里有 locked 形状时启动，10ms 轮询，4 次稳定反算 adj 写回
  const LOCK_POLL_MS = 10;
  const LOCK_STABLE_THRESHOLD = 4;
  let lockMonitor = {
    timer: null,
    lastDims: {},    // shapeId -> 'w|h' 字符串
    stableCount: {}, // shapeId -> 连续稳定次数
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

  function loadLocksViaTags() {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const sel = ctx.presentation.getSelectedShapes();
          sel.load('items/id');
          await ctx.sync();
          const locks = {};
          for (const sh of sel.items) {
            try {
              // Mac LTSC: getItem 返回 ClientResult 代理，直接 .value 读
              const tag = sh.tags.getItem(LOCK_TAG_KEY);
              tag.load('value');
              await ctx.sync();
              const cm = parseFloat(tag.value);
              if (Number.isFinite(cm) && cm > 0) {
                locks[sh.id] = cm;
              }
            } catch (_) {
              // tag 不存在，跳过
            }
          }
          resolve({ ok: true, locks });
        } catch (e) {
          resolve({ ok: false, error: e });
        }
      });
    });
  }

  function saveLocksViaTags(locks) {
    return new Promise((resolve) => {
      PowerPoint.run(async (ctx) => {
        try {
          const sel = ctx.presentation.getSelectedShapes();
          sel.load('items/id');
          await ctx.sync();
          for (const sh of sel.items) {
            const cm = locks[sh.id];
            try {
              if (cm == null) {
                sh.tags.delete(LOCK_TAG_KEY);
              } else {
                sh.tags.add(LOCK_TAG_KEY, String(cm));
              }
            } catch (_) {
              // 单个 shape 写失败不影响其他
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
        sel.load('items/id, items/name, items/width, items/height, items/type, items/adjustments');
        await ctx.sync();
        const shapes = [];
        for (const sh of sel.items) {
          // 关键：task pane 上下文里 adjustments 子项的 value 不会自动跟随
          // shapes.load 一起填，必须显式 load items/value 再 sync
          let cm = null;
          let isRoundRect = false;
          if (sh.adjustments && sh.adjustments.count > 0) {
            isRoundRect = true;
            sh.adjustments.load('items/value');
            await ctx.sync();
            const value = sh.adjustments.get(0).value;
            if (Number.isFinite(value) && value > 0) {
              const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
              cm = value * minSideCm;
            }
          }
          shapes.push({
            id: sh.id,
            name: sh.name,
            type: sh.type,
            width: sh.width,
            height: sh.height,
            minSideCm: Math.min(sh.width, sh.height) / PT_PER_CM,
            currentCm: cm,
            isRoundRect,
            locked: false,
            lockedCm: null,
          });
        }
        selectedShapes = shapes;
      });
      // 读 lock 后端（shape.tags）
      const tagResult = await loadLocksViaTags();
      if (tagResult.ok) {
        for (const s of selectedShapes) {
          if (tagResult.locks[s.id] != null) {
            s.locked = true;
            s.lockedCm = tagResult.locks[s.id];
          }
        }
      }
      renderUI();
      // 锁监控：选区里有 locked 时启动，否则停
      if (selectedShapes.some((s) => s.locked)) {
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
    lockMonitor.lastDims = {};
    lockMonitor.stableCount = {};
    lockMonitor.timer = setInterval(monitorTick, LOCK_POLL_MS);
  }

  function stopLockMonitor() {
    if (lockMonitor.timer) {
      clearInterval(lockMonitor.timer);
      lockMonitor.timer = null;
    }
    lockMonitor.lastDims = {};
    lockMonitor.stableCount = {};
  }

  async function monitorTick() {
    const locked = selectedShapes.filter((s) => s.locked);
    if (locked.length === 0) {
      stopLockMonitor();
      return;
    }
    let appliedIds = [];
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id, items/width, items/height');
        await ctx.sync();
        for (const sh of sel.items) {
          const target = locked.find((x) => x.id === sh.id);
          if (!target) continue;
          const currentKey = `${sh.width.toFixed(4)}|${sh.height.toFixed(4)}`;
          const lastKey = lockMonitor.lastDims[sh.id];
          if (lastKey === currentKey) {
            lockMonitor.stableCount[sh.id] = (lockMonitor.stableCount[sh.id] || 0) + 1;
          } else {
            lockMonitor.stableCount[sh.id] = 0;
            lockMonitor.lastDims[sh.id] = currentKey;
          }
          if (lockMonitor.stableCount[sh.id] >= LOCK_STABLE_THRESHOLD) {
            const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
            if (minSideCm > 0) {
              const targetCm = Math.min(target.lockedCm, minSideCm / 2);
              const newAdj = (targetCm / minSideCm) * ADJ_SCALE;
              if (Number.isFinite(newAdj) && newAdj >= 0) {
                sh.adjustments.set(0, newAdj);
                appliedIds.push(sh.id);
              }
            }
          }
        }
        await ctx.sync();
      });
      if (appliedIds.length > 0) {
        // 重应用后不要 refreshSelection（会引发 selection redraw 闪烁），只更新内存
        for (const s of selectedShapes) {
          if (appliedIds.includes(s.id)) {
            // 重新读 adj 算出 currentCm
            s.currentCm = s.lockedCm; // 简化为锁值
          }
        }
        showToast(`🔒 自动重应用了 ${appliedIds.length} 个锁定的 R 角`);
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
      const tag = s.locked
        ? `<span class="shape-lock">🔒 ${s.lockedCm.toFixed(2)}cm</span>`
        : (s.isRoundRect ? '' : '<span class="shape-warn">非圆角矩形</span>');
      const rText = s.currentCm != null ? `${s.currentCm.toFixed(2)}cm` : '—';
      row.innerHTML = `<span class="shape-name">${s.name || '(未命名)'}</span><span class="shape-r">${rText}</span>${tag}`;
      list.appendChild(row);
    }
  }

  function updateLockButton() {
    const btn = $('lock-btn');
    if (!btn) return;
    if (selectedShapes.length === 0) {
      btn.disabled = true;
      $('lock-icon').textContent = '🔓';
      $('lock-label').textContent = '锁定 R 角';
      $('lock-hint').textContent = '读选中…';
      return;
    }
    btn.disabled = false;
    const roundShapes = selectedShapes.filter((s) => s.isRoundRect);
    const allLocked = roundShapes.length > 0 && roundShapes.every((s) => s.locked);
    $('lock-icon').textContent = allLocked ? '🔒' : '🔓';
    $('lock-label').textContent = allLocked ? '解锁 R 角' : '锁定 R 角';
    $('lock-hint').textContent = allLocked
      ? `已锁定 ${roundShapes.length} 个（跟随形状 tag，跨设备保留）`
      : `锁定后 R 角按厘米值保持；改变形状大小时自动按比例调整`;
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
    // 输入值按当前单位换算成 cm
    const cm = valueToCm(raw, currentUnit);
    let updated = 0;
    let failed = 0;
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
          } catch (_) {
            // 这个形状可能不是 roundRect，set 失败
            failed++;
          }
        }
        await ctx.sync();
      });
      if (failed === 0) {
        const displayVal = currentUnit === '%'
          ? `${raw.toFixed(1)}%`
          : `${raw.toFixed(2)} 厘米`;
        showToast(`✅ 已更新 ${updated} 个圆角矩形为 ${displayVal}`);
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
    }
  }

  /** 锁定 / 解锁 R 角（用 shape.tags，跟 .pptx 文件走） */
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
    let touched = 0;
    for (const s of roundShapes) {
      if (allLocked) {
        // 解锁：不写 tag
      } else {
        // 锁定：优先用输入框值，否则用当前 R 角
        const inputCm = Number.isFinite(inputVal) && inputVal > 0
          ? valueToCm(inputVal, currentUnit)
          : s.currentCm;
        if (inputCm > 0) locks[s.id] = inputCm;
      }
      touched++;
    }
    const r = await saveLocksViaTags(locks);
    if (!r.ok) {
      showToast('操作失败：' + (r.error?.message || r.error));
      return;
    }
    showToast(allLocked ? `已解锁 ${touched} 个（跟随 .pptx 文件）` : `已锁定 ${touched} 个（跟随 .pptx 文件）`);
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
  }

  // ---------------- 初始化 ----------------

  Office.onReady(() => {
    bindEvents();
    refreshSelection();
    // 选区变化：用户在 PPT 里点别的形状、框选、切页 都会触发
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => refreshSelection()
    );
  });
})();
