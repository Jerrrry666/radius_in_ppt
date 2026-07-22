/*
 * dialog.js — R 角调整 Dialog（纯 Office.js 实现）
 *
 * 工作流：
 *   1. 打开 dialog → getSelectedShapes() → 显示选中的圆角矩形
 *   2. 用户输入 R 角 → 「应用 R 角」→ adjustments.set(0, newVal)
 *   3. 锁定信息存到 OOXML CustomXmlPart（namespace = LOCK_XML_NS），跟 .pptx 文件走
 *      如果 customXmlParts 在当前平台不可用，降级到 localStorage
 *   4. toast: "已更新 N 个圆角矩形为 X 厘米"
 *
 * 监听 PPT 选区变化：DocumentSelectionChanged 事件 → 自动 refresh
 * 多页 PPT：getSelectedShapes() 只返回当前页选中的形状，切页后选区变化 → 自动 refresh
 *
 * 锁定 R 角 + 自动重应用：
 *   - 选区里有 locked 形状时，启动 setInterval 轮询（10ms）
 *   - 若连续 4 次尺寸无变化（≈40ms 稳定）→ 视为用户松手 → 反算 adj 写回
 *   - 拖拽中尺寸在变 → 跳过 apply，避免和拖拽手感冲突
 *   - 注：Office.js PowerPoint 没有 shape change 事件（详见 AGENTS.md 4 节），
 *         这是当前的 API 限制下的最优解
 */

(function () {
  const $ = (id) => document.getElementById(id);
  const PT_PER_CM = 28.3464567;       // 1 cm = 28.3464567 pt
  // Mac LTSC Office.js: adjustments.get(0).value 返回 0~1 的小数（占短边比例）
  // OOXML 原始是 0~50000（0~50%），但 Office.js 在 Mac dialog 上下文里 normalize 成 0~1 了
  const ADJ_SCALE = 1;
  const LOCK_STORAGE_KEY = 'radius_in_ppt_locks_v2';
  const LOCK_XML_NS = 'https://radius.jerrrry666.com/radius-in-ppt/locks/v1';
  let lockBackend = 'unknown';         // 'customXmlPart' | 'localStorage' | 'none'

  /** 当前选中的形状（refreshSelection 填充） */
  let selectedShapes = [];  // [{id, name, width, height, currentCm, locked, lockedCm, isRoundRect, adjCount}]

  // ---------------- CustomXmlPart 锁存储（跟随 .pptx 文件） ----------------
  // Mac dialog 里 customProperties 不可用，改用 Common API 的 customXmlParts：
  // 写入 .pptx 文件的 customXml 段（OOXML 标准的隐藏 XML 块），换机器/发文件都跟着走
  // XML 格式：
  //   <locks xmlns="...">
  //     <lock shapeId="123" cm="1.28"/>
  //   </locks>

  function buildLocksXml(locks) {
    const entries = Object.entries(locks)
      .filter(([_, cm]) => cm != null && Number.isFinite(cm) && cm > 0)
      .map(([id, cm]) => `    <lock shapeId="${escapeXml(id)}" cm="${cm.toFixed(4)}"/>`)
      .join('\n');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<locks xmlns="${LOCK_XML_NS}">\n${entries}\n</locks>`;
  }

  function parseLocksXml(xml) {
    const locks = {};
    if (!xml) return locks;
    try {
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const nodes = doc.getElementsByTagNameNS(LOCK_XML_NS, 'lock');
      for (const n of nodes) {
        const id = n.getAttribute('shapeId');
        const cm = parseFloat(n.getAttribute('cm'));
        if (id && Number.isFinite(cm)) locks[id] = cm;
      }
    } catch (e) {
      console.warn('parseLocksXml failed:', e);
    }
    return locks;
  }

  function escapeXml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;'
    })[c]);
  }

  function loadLocksCustomXml() {
    return new Promise((resolve, reject) => {
      if (!Office.context.document || !Office.context.document.customXmlParts) {
        reject(new Error('customXmlParts not available'));
        return;
      }
      Office.context.document.customXmlParts.getByNamespaceAsync(
        LOCK_XML_NS,
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(result.error?.message || 'getByNamespace failed'));
            return;
          }
          if (!result.value || result.value.length === 0) {
            resolve({});
            return;
          }
          result.value[0].getXmlAsync((xmlResult) => {
            if (xmlResult.status !== Office.AsyncResultStatus.Succeeded) {
              reject(new Error(xmlResult.error?.message || 'getXml failed'));
              return;
            }
            resolve(parseLocksXml(xmlResult.value));
          });
        }
      );
    });
  }

  function saveLocksCustomXml(locks) {
    return new Promise((resolve, reject) => {
      if (!Office.context.document || !Office.context.document.customXmlParts) {
        reject(new Error('customXmlParts not available'));
        return;
      }
      const xml = buildLocksXml(locks);
      Office.context.document.customXmlParts.getByNamespaceAsync(
        LOCK_XML_NS,
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(result.error?.message || 'getByNamespace failed'));
            return;
          }
          if (result.value && result.value.length > 0) {
            result.value[0].setXmlAsync(xml, (setResult) => {
              if (setResult.status === Office.AsyncResultStatus.Succeeded) resolve();
              else reject(new Error(setResult.error?.message || 'setXml failed'));
            });
          } else {
            Office.context.document.customXmlParts.addAsync(xml, (addResult) => {
              if (addResult.status === Office.AsyncResultStatus.Succeeded) resolve();
              else reject(new Error(addResult.error?.message || 'addAsync failed'));
            });
          }
        }
      );
    });
  }

  // 统一接口：先试 CustomXmlPart（跟文件走），失败回退 localStorage（机机本地）
  async function loadLocks() {
    try {
      const locks = await loadLocksCustomXml();
      return { locks, backend: 'customXmlPart' };
    } catch (e) {
      console.warn('CustomXmlPart 不可用，降级到 localStorage:', e);
      return { locks: loadLocksLS(), backend: 'localStorage' };
    }
  }

  async function saveLocks(locks, preferBackend) {
    // preferBackend: 写入时优先用这个 backend（跟读保持一致，避免读写分离）
    if (preferBackend === 'localStorage') {
      saveLocksLS(locks);
      return 'localStorage';
    }
    try {
      await saveLocksCustomXml(locks);
      return 'customXmlPart';
    } catch (e) {
      console.warn('CustomXmlPart 写失败，降级到 localStorage:', e);
      saveLocksLS(locks);
      return 'localStorage';
    }
  }

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
      // 先从 CustomXmlPart（或 localStorage fallback）读锁，跟随 .pptx 文件
      const { locks, backend: lb } = await loadLocks();
      lockBackend = lb;

      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load('items/id,items/name,items/width,items/height,items/adjustments');
        await ctx.sync();

        selectedShapes = [];
        for (const sh of sel.items) {
          const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
          // 恢复原版：get(0) 在 Mac LTSC dialog 里返回的是 ClientResult 代理，.value 直接拿
          const adjCount = sh.adjustments.count;
          const adjResult = sh.adjustments.get(0);
          await ctx.sync();
          const adj = adjResult.value;

          const isRoundRect = (typeof adjCount === 'number' ? adjCount : 0) > 0;
          let currentCm = (isRoundRect && Number.isFinite(adj))
            ? (adj / ADJ_SCALE) * minSideCm
            : 0;
          const lockedCm = locks[sh.id];

          // 自动重应用锁：如果 locked 且当前 cm 跟锁定值漂移了 > 0.005cm，按当前尺寸反算比例写回
          let reApplied = false;
          if (lockedCm != null && isRoundRect && minSideCm > 0
              && Math.abs(currentCm - lockedCm) > 0.005) {
            const targetCm = Math.min(lockedCm, minSideCm / 2);
            const newAdj = (targetCm / minSideCm) * ADJ_SCALE;
            if (Number.isFinite(newAdj)) {
              sh.adjustments.set(0, newAdj);
              currentCm = targetCm;   // 写回去之后新的 cm 就是 targetCm
              reApplied = true;
            }
          }

          // 调试：把所有 raw 值都打出来
          debugLines.push(`Shape id=${sh.id} name="${sh.name}"`);
          debugLines.push(`  width=${sh.width}pt  height=${sh.height}pt  minSide=${minSideCm.toFixed(2)}cm`);
          debugLines.push(`  adjustments.count = ${adjCount} (type=${typeof adjCount})`);
          debugLines.push(`  adjustments.get(0) = ${adjResult} (type=${typeof adjResult})`);
          debugLines.push(`  adjustments.get(0).value = ${adj} (type=${typeof adj})`);
          debugLines.push(`  isRoundRect 判定: ${isRoundRect ? 'true' : 'false'}`);
          if (lockedCm != null) {
            debugLines.push(`  锁定: ${lockedCm.toFixed(2)}cm${reApplied ? ' (本轮已重应用)' : ''}`);
          }
          debugLines.push('');

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
            reApplied,
          });
        }
      });
      renderUI();
      const dbg = $('debug-out');
      if (dbg) dbg.textContent = debugLines.join('\n') || '（无选中）';
      // 自动重应用锁的提示
      const reappliedCount = selectedShapes.filter((s) => s.reApplied).length;
      if (reappliedCount > 0) {
        showToast(`🔒 自动重应用了 ${reappliedCount} 个锁定的 R 角`);
      }
      // 根据当前选区是否有 locked 形状，启动/停止轮询监控
      const hasLocked = selectedShapes.some((s) => s.locked);
      if (hasLocked) startLockMonitor();
      else stopLockMonitor();
    } catch (err) {
      setStatus('选区', '读失败：' + (err.message || err), 'status-warn');
      showToast('读选区失败: ' + (err.message || err));
      const dbg = $('debug-out');
      if (dbg) dbg.textContent = '读失败：' + (err.message || err);
    }
  }

  // ---------------- 锁定监控（轮询稳定检测） ----------------
  // 目的：用户拖完形状松手后，自动把 R 角反算回锁定值
  // 策略：每 10ms poll 一次，若连续 4 次尺寸无变化（≈ 40ms）则视为松手
  // 期间尺寸在变（拖拽中）→ 跳过 apply，避免和用户的拖动手感冲突
  // 注意：apply 后不要调 refreshSelection()（会再触发一次完整 getSelectedShapes + sync，
  //       导致 PowerPoint 选区高亮重画，肉眼看是闪烁），直接改内存里的 currentCm 然后 renderUI

  const LOCK_POLL_MS = 10;
  const LOCK_STABLE_THRESHOLD = 4;  // 4 * 10ms = 40ms 稳定才 apply
  let lockMonitor = null;            // { timer, lastDims, stableCount }

  function startLockMonitor() {
    if (lockMonitor) return;
    lockMonitor = { timer: null, lastDims: {}, stableCount: {} };
    lockMonitor.timer = setInterval(monitorTick, LOCK_POLL_MS);
  }

  function stopLockMonitor() {
    if (!lockMonitor) return;
    clearInterval(lockMonitor.timer);
    lockMonitor = null;
  }

  async function monitorTick() {
    // 兜底：如果选区里没有 locked，关闭 monitor
    const locked = selectedShapes.filter((s) => s.locked);
    if (locked.length === 0) {
      stopLockMonitor();
      return;
    }
    let appliedIds = [];
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        // 只读 id + 尺寸，不读 adjustments（避免不必要的数据同步）
        sel.load('items/id,items/width,items/height');
        await ctx.sync();
        for (const sh of sel.items) {
          const id = sh.id;
          const target = locked.find((x) => x.id === id);
          if (!target) continue;
          const currentKey = `${sh.width.toFixed(4)}|${sh.height.toFixed(4)}`;
          const lastKey = lockMonitor.lastDims[id];
          if (lastKey === currentKey) {
            lockMonitor.stableCount[id] = (lockMonitor.stableCount[id] || 0) + 1;
          } else {
            lockMonitor.stableCount[id] = 0;
            lockMonitor.lastDims[id] = currentKey;
          }
          if (lockMonitor.stableCount[id] >= LOCK_STABLE_THRESHOLD) {
            const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
            if (minSideCm > 0) {
              const targetCm = Math.min(target.lockedCm, minSideCm / 2);
              const newAdj = (targetCm / minSideCm) * ADJ_SCALE;
              if (Number.isFinite(newAdj)) {
                sh.adjustments.set(0, newAdj);
                appliedIds.push(id);
              }
            }
          }
        }
      });
    } catch (_) {
      // 静默吞错：拖拽中可能 selection 临时为空 / shape 被删
    }
    if (appliedIds.length > 0) {
      // 不调 refreshSelection()，只改内存 + 重渲染 dialog UI
      for (const id of appliedIds) {
        const s = selectedShapes.find((x) => x.id === id);
        if (s && s.locked) s.currentCm = s.lockedCm;
      }
      renderUI();
      showToast(`🔒 自动重应用了 ${appliedIds.length} 个锁定的 R 角`);
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
      lockHint.title = 'CustomXmlPart 不可用，锁定暂存本机 localStorage（不会跟 .pptx 走）';
    } else if (lockBackend === 'customXmlPart') {
      lockHint.textContent = (lockHint.textContent || '') + ' · 跟文件走';
      lockHint.title = '锁定存到 .pptx 文件的 CustomXmlPart，换机器/发文件都会保留';
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
          // ADJ_SCALE=1 → newAdj 是 0~0.5 的小数比例，不能 round 到整数
          const newAdj = (targetCm / minSideCm) * ADJ_SCALE;
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

  /** 锁定 / 解锁 R 角（CustomXmlPart 优先跟随文件，localStorage 降级） */
  async function onToggleLock() {
    if (selectedShapes.length === 0) {
      showToast('请先在 PPT 里框选圆角矩形');
      return;
    }
    const allLocked = selectedShapes.every((s) => s.locked);
    const inputCm = parseFloat($('radius-input').value);
    let touched = 0;
    try {
      // 读当前所有锁（从与 refreshSelection 同一个 backend 读，避免分离）
      const { locks: currentLocks } = await loadLocks();
      const newLocks = { ...currentLocks };
      for (const s of selectedShapes) {
        if (allLocked) {
          delete newLocks[s.id];
        } else {
          const lockCm = Number.isFinite(inputCm) && inputCm > 0
            ? inputCm
            : s.currentCm;
          if (lockCm > 0) newLocks[s.id] = lockCm;
        }
        touched++;
      }
      // 写回（preferBackend 跟读保持一致：customXmlPart 优先）
      const backend = await saveLocks(newLocks, lockBackend);
      lockBackend = backend;
      const where = backend === 'customXmlPart' ? '（跟随 .pptx 文件）' : '（本机 localStorage）';
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
          const newAdj = (lockCm / minSideCm) * ADJ_SCALE;
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
