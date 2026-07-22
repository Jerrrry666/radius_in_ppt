/*
 * dialog.js — R 角调整 Dialog（纯 Office.js 实现）
 *
 * 工作流：
 *   1. 打开 dialog → getSelectedShapes() → 显示选中的圆角矩形
 *   2. 用户输入 R 角 → 「应用 R 角」→ adjustments.set(0, newVal)
 *   3. 锁定信息优先存到 customProperty（key = "lock:{shapeId}"）
 *      如果 customProperties API 不可用（部分 dialog 上下文缺失），降级到 localStorage
 *   4. toast: "已更新 N 个圆角矩形为 X 厘米"
 *
 * 监听 PPT 选区变化：DocumentSelectionChanged 事件 → 自动 refresh
 *
 * 多页 PPT：getSelectedShapes() 只返回当前页选中的形状。
 *          切页后选区变化 → 自动 refresh。
 */

(function () {
  const $ = (id) => document.getElementById(id);
  const PT_PER_CM = 28.3464567;       // 1 cm = 28.3464567 pt
  const ADJ_SCALE = 100000;            // adj value 0~50000 对应 0%~50%
  const LOCK_STORAGE_KEY = 'radius_in_ppt_locks_v2';
  let lockBackend = 'unknown';         // 'customProperty' | 'localStorage' | 'none'

  /** 当前选中的形状（refreshSelection 填充） */
  let selectedShapes = [];  // [{id, name, width, height, currentCm, locked, lockedCm, isRoundRect, adjCount}]

  // ---------------- localStorage 锁定降级 ----------------

  function loadLocksLS() {
    try { return JSON.parse(localStorage.getItem(LOCK_STORAGE_KEY) || '{}'); }
    catch (_) { return {}; }
  }
  function saveLocksLS(locks) {
    try { localStorage.setItem(LOCK_STORAGE_KEY, JSON.stringify(locks)); } catch (_) {}
  }
  function getLockLS(shapeId) {
    return loadLocksLS()[shapeId] || null;
  }
  function setLockLS(shapeId, cm) {
    const m = loadLocksLS();
    if (cm === null) delete m[shapeId];
    else m[shapeId] = cm;
    saveLocksLS(m);
  }

  // ---------------- 初始化 ----------------

  Office.onReady(() => {
    bindEvents();
    refreshSelection();
    // 监听选区变化：用户在 PPT 里点别的形状、框选、切页 都会触发
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => refreshSelection()
    );
  });

  function bindEvents() {
    $('apply-btn').addEventListener('click', onApply);
    $('lock-btn').addEventListener('click', onToggleLock);
    $('reapply-btn').addEventListener('click', onReapply);
    $('rescan-btn').addEventListener('click', refreshSelection);
    $('radius-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') onApply();
    });
  }

  // ---------------- 读选区 ----------------

  async function refreshSelection() {
    setStatus('选区', '读选中…', 'status-empty');
    const debugLines = [];
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id,items/name,items/width,items/height,items/adjustments');
        await ctx.sync();

        // 收集锁定信息：先尝试 customProperty，不可用则降级 localStorage
        const locks = {};
        lockBackend = 'none';
        if (ctx.presentation.customProperties) {
          try {
            const props = ctx.presentation.customProperties;
            props.load('items');
            await ctx.sync();
            for (const cp of props.items || []) {
              cp.load('key, value');
            }
            await ctx.sync();
            for (const cp of props.items || []) {
              const k = cp.key;
              if (k && k.startsWith('lock:')) {
                const cm = parseFloat(cp.value);
                if (Number.isFinite(cm)) locks[k.slice(5)] = cm;
              }
            }
            lockBackend = 'customProperty';
          } catch (e) {
            console.warn('customProperties 不可用，降级到 localStorage:', e);
            lockBackend = 'localStorage';
            Object.assign(locks, loadLocksLS());
          }
        } else {
          lockBackend = 'localStorage';
          Object.assign(locks, loadLocksLS());
        }

        selectedShapes = [];
        for (const sh of sel.items) {
          const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
          // 必须显式 load('count, items/value')，否则 Mac dialog 上下文里 .count 可能是 0、.value 是 undefined
          sh.adjustments.load('count, items/value');
          await ctx.sync();
          const adjCount = sh.adjustments.count;
          const adj = (adjCount > 0 && sh.adjustments.items.length > 0)
            ? sh.adjustments.items[0].value
            : 0;

          // 调试：把所有 raw 值都打出来
          debugLines.push(`Shape id=${sh.id} name="${sh.name}"`);
          debugLines.push(`  width=${sh.width}pt  height=${sh.height}pt  minSide=${minSideCm.toFixed(2)}cm`);
          debugLines.push(`  adjustments.count = ${adjCount} (type=${typeof adjCount})`);
          debugLines.push(`  adjustments.items.length = ${sh.adjustments.items.length}`);
          debugLines.push(`  adjustments.items[0].value = ${adj} (type=${typeof adj})`);
          debugLines.push(`  isRoundRect 判定: ${(typeof adjCount === 'number' && adjCount > 0) ? 'true' : 'false'}`);
          debugLines.push('');

          const isRoundRect = (typeof adjCount === 'number' ? adjCount : 0) > 0;
          const currentCm = (isRoundRect && Number.isFinite(adj))
            ? (adj / ADJ_SCALE) * minSideCm
            : 0;
          const lockedCm = locks[sh.id];
          selectedShapes.push({
            id: sh.id,
            name: sh.name,
            width: sh.width,
            height: sh.height,
            currentAdj: adj,
            currentCm,
            isRoundRect,
            adjCount: adjCount || 0,
            locked: lockedCm != null,
            lockedCm: lockedCm == null ? null : lockedCm,
          });
        }
      });
      renderUI();
      const dbg = $('debug-out');
      if (dbg) dbg.textContent = debugLines.join('\n') || '（无选中）';
    } catch (err) {
      setStatus('选区', '读失败：' + (err.message || err), 'status-warn');
      showToast('读选区失败: ' + (err.message || err));
      const dbg = $('debug-out');
      if (dbg) dbg.textContent = '读失败：' + (err.message || err);
    }
  }

  // ---------------- UI helpers ----------------

  function setStatus(label, text, cardClass) {
    $('status-text').textContent = text;
    $('status-card').className = 'status-card ' + cardClass;
    const labelEl = document.querySelector('.status-row .status-label');
    if (labelEl) labelEl.textContent = label;
  }

  // ---------------- 渲染 ----------------

  function renderUI() {
    // 状态卡
    if (selectedShapes.length === 0) {
      setStatus('选区', '未选中', 'status-warn');
      $('current-radius').textContent = '—';
      $('locked-count').textContent = '—';
    } else {
      setStatus('选区', `${selectedShapes.length} 个`, 'status-ok');
      if (selectedShapes.length === 1) {
        $('current-radius').textContent = `${selectedShapes[0].currentCm.toFixed(2)} 厘米`;
      } else {
        const cms = selectedShapes.map((s) => s.currentCm);
        const allSame = cms.every((c) => Math.abs(c - cms[0]) < 0.005);
        $('current-radius').textContent = allSame
          ? `${cms[0].toFixed(2)} 厘米（多选相同）`
          : `${Math.min(...cms).toFixed(2)} ~ ${Math.max(...cms).toFixed(2)} 厘米`;
      }
      const lockedN = selectedShapes.filter((s) => s.locked).length;
      $('locked-count').textContent = `${lockedN} / ${selectedShapes.length}`;
    }
    // 锁定后端提示
    const lockHint = $('lock-hint');
    if (lockBackend === 'localStorage') {
      lockHint.textContent = (lockHint.textContent || '') + ' · 本机存储';
      lockHint.title = 'customProperty 在本 dialog 不可用，锁定暂存本机 localStorage';
    }

    // 列表
    const list = $('shape-list');
    list.innerHTML = '';
    $('list-count').textContent = `${selectedShapes.length} 个`;
    if (selectedShapes.length === 0) {
      list.innerHTML = '<div class="empty-list">在 PPT 里框选形状后会出现在这里</div>';
    } else {
      for (const s of selectedShapes) {
        const row = document.createElement('div');
        row.className = 'shape-row';
        if (!s.isRoundRect) row.classList.add('not-round');
        const tag = document.createElement('div');
        tag.className = 'shape-name';
        const lockMark = s.locked ? ' 🔒' : '';
        const name = s.name ? s.name : `Shape ${s.id}`;
        tag.textContent = `${name}${lockMark}`;
        const meta = document.createElement('div');
        meta.className = 'shape-meta';
        if (!s.isRoundRect) {
          meta.textContent = '非圆角矩形';
          meta.style.color = '#c50f1f';
        } else if (s.locked) {
          meta.textContent = `${s.currentCm.toFixed(2)}cm · 锁 ${s.lockedCm.toFixed(2)}`;
        } else {
          meta.textContent = `${s.currentCm.toFixed(2)} cm`;
        }
        row.appendChild(tag);
        row.appendChild(meta);
        list.appendChild(row);
      }
    }

    // 锁定按钮状态
    updateLockButton();
  }

  function updateLockButton() {
    const btn = $('lock-btn');
    const applyBtn = $('apply-btn');
    const reapplyBtn = $('reapply-btn');
    const label = $('lock-label');
    const icon = $('lock-icon');
    const hint = $('lock-hint');
    const roundN = selectedShapes.filter((s) => s.isRoundRect).length;
    if (selectedShapes.length === 0) {
      btn.disabled = true;
      applyBtn.disabled = true;
      reapplyBtn.disabled = true;
      btn.classList.remove('is-locked');
      label.textContent = '锁定 R 角';
      icon.textContent = '🔓';
      hint.textContent = '请先在 PPT 里框选圆角矩形';
      return;
    }
    if (roundN === 0) {
      // 全是非圆角
      btn.disabled = true;
      applyBtn.disabled = true;
      reapplyBtn.disabled = true;
      hint.textContent = '选中的都不是圆角矩形';
      return;
    }
    btn.disabled = false;
    applyBtn.disabled = false;
    const lockedN = selectedShapes.filter((s) => s.locked).length;
    reapplyBtn.disabled = lockedN === 0;
    if (lockedN === 0) {
      btn.classList.remove('is-locked');
      label.textContent = '锁定 R 角';
      icon.textContent = '🔓';
      hint.textContent = `当前 0 / ${selectedShapes.length} 已锁定`;
    } else if (lockedN === selectedShapes.length) {
      btn.classList.add('is-locked');
      label.textContent = '解锁 R 角';
      icon.textContent = '🔒';
      hint.textContent = `当前 ${lockedN} / ${selectedShapes.length} 已锁定`;
    } else {
      btn.classList.add('is-locked');
      label.textContent = '全部锁定';
      icon.textContent = '🔒';
      hint.textContent = `当前 ${lockedN} / ${selectedShapes.length} 已锁定`;
    }
  }

  // ---------------- 操作 ----------------

  /** 应用 R 角：所有选中的形状都改成输入的 cm 值 */
  async function onApply() {
    if (selectedShapes.length === 0) {
      showToast('请先在 PPT 里框选圆角矩形');
      return;
    }
    const cm = parseFloat($('radius-input').value);
    if (!Number.isFinite(cm) || cm < 0) {
      showToast('请输入有效的 R 角值');
      return;
    }
    let updated = 0;
    let failed = 0;
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id,items/width,items/height,items/adjustments');
        await ctx.sync();
        for (const sh of sel.items) {
          // sh.width / sh.height 是已 load 的真数字
          const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
          if (minSideCm <= 0) {
            failed++;
            continue;
          }
          const targetCm = Math.min(cm, minSideCm / 2);
          const newAdj = Math.round((targetCm / minSideCm) * ADJ_SCALE);
          if (!Number.isFinite(newAdj)) {
            failed++;
            continue;
          }
          try {
            sh.adjustments.set(0, newAdj);
            updated++;
          } catch (e) {
            // 这个形状可能不是 roundRect，set 失败
            console.warn('set 失败:', sh.id, e);
            failed++;
          }
        }
        await ctx.sync();
      });
      if (failed === 0) {
        showToast(`✅ 已更新 ${updated} 个圆角矩形为 ${cm.toFixed(2)} 厘米`);
      } else {
        showToast(`⚠️ ${updated} 个成功，${failed} 个失败（可能不是圆角矩形）`);
      }
      await refreshSelection();
    } catch (err) {
      showToast('应用失败：' + (err.message || err));
    }
  }

  /** 锁定 / 解锁 R 角（customProperty 优先，localStorage 降级） */
  async function onToggleLock() {
    if (selectedShapes.length === 0) {
      showToast('请先在 PPT 里框选圆角矩形');
      return;
    }
    const allLocked = selectedShapes.every((s) => s.locked);
    const inputCm = parseFloat($('radius-input').value);
    let touched = 0;
    try {
      if (lockBackend === 'localStorage') {
        for (const s of selectedShapes) {
          if (allLocked) {
            setLockLS(s.id, null);
          } else {
            const lockCm = Number.isFinite(inputCm) && inputCm > 0 ? inputCm : s.currentCm;
            setLockLS(s.id, lockCm);
          }
          touched++;
        }
      } else {
        await PowerPoint.run(async (ctx) => {
          const props = ctx.presentation.customProperties;
          for (const s of selectedShapes) {
            const key = `lock:${s.id}`;
            const item = props.getItemOrNullObject(key);
            await ctx.sync();
            if (allLocked) {
              if (!item.isNullObject) {
                item.delete();
                touched++;
              }
            } else {
              const lockCm = Number.isFinite(inputCm) && inputCm > 0
                ? inputCm
                : s.currentCm;
              if (item.isNullObject) {
                props.add(key, String(lockCm));
              } else {
                item.value = String(lockCm);
              }
              touched++;
            }
            await ctx.sync();
          }
        });
      }
      const where = lockBackend === 'localStorage' ? '（本机 localStorage）' : '';
      showToast(allLocked ? `已解锁 ${touched} 个${where}` : `已锁定 ${touched} 个${where}`);
      await refreshSelection();
    } catch (err) {
      showToast('操作失败：' + (err.message || err));
    }
  }

  /** 重新应用锁定：按当前形状大小反算 adj */
  async function onReapply() {
    const locked = selectedShapes.filter((s) => s.locked);
    if (locked.length === 0) {
      showToast('当前选区没有锁定的圆角矩形');
      return;
    }
    let updated = 0;
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id,items/width,items/height,items/adjustments');
        await ctx.sync();
        for (const sh of sel.items) {
          const target = locked.find((x) => x.id === sh.id);
          if (!target) continue;
          const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
          if (minSideCm <= 0) continue;
          const lockCm = Math.min(target.lockedCm, minSideCm / 2);
          const newAdj = Math.round((lockCm / minSideCm) * ADJ_SCALE);
          if (!Number.isFinite(newAdj)) continue;
          try {
            sh.adjustments.set(0, newAdj);
            updated++;
          } catch (e) {
            console.warn('reapply set 失败:', sh.id, e);
          }
        }
        await ctx.sync();
      });
      showToast(`✅ 已重新应用 ${updated} 个锁定`);
      await refreshSelection();
    } catch (err) {
      showToast('操作失败：' + (err.message || err));
    }
  }

  // ---------------- UI helpers ----------------

  let toastTimer = null;
  function showToast(msg) {
    let el = document.querySelector('.toast');
    if (!el) {
      el = document.createElement('div');
      el.className = 'toast';
      document.body.appendChild(el);
    }
    el.textContent = msg;
    el.classList.add('show');
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
  }
})();
