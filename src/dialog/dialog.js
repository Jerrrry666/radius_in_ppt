/*
 * dialog.js — R 角调整 Dialog 逻辑
 *
 * 职责：
 *   1. 监听选区变化，显示当前选区中圆角矩形的 R 角值（按选中数量的"代表值"）
 *   2. 用户输入数值 + 点"应用" → 写入选区所有圆角矩形
 *   3. "锁定/解锁"按钮 → 切换当前选区所有圆角矩形的锁定状态
 *   4. 锁定时，SelectionChanged 事件触发重新应用 R 角绝对值
 */

(function () {
  const $ = (id) => document.getElementById(id);

  let lastSelectionInfo = null; // { all, roundedRects, any, noneRounded }

  Office.onReady(() => {
    bindEvents();
    refresh();
    RadiusCore.installSelectionChangedAutoReapply();
  });

  function bindEvents() {
    $('apply-btn').addEventListener('click', onApply);
    $('lock-btn').addEventListener('click', onToggleLock);
    $('reapply-btn').addEventListener('click', onReapply);
    $('radius-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') onApply();
    });
    // 选区变化（用 Office 事件更可靠）
    if (Office.context.document.addHandlerAsync) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        refresh
      );
    }
  }

  /**
   * 拉取最新选区信息，更新 UI。
   */
  function refresh() {
    RadiusCore.getSelectionInfo()
      .then((info) => {
        lastSelectionInfo = info;
        renderStatus(info);
        renderCurrentRadius(info);
        renderLockButton(info);
      })
      .catch((err) => {
        console.error('[R 角调整] 读取选区失败:', err);
        $('status-text').textContent = '读取失败：' + (err && err.message ? err.message : err);
        $('status-card').className = 'status-card status-warn';
      });
  }

  function renderStatus(info) {
    const card = $('status-card');
    const text = $('status-text');
    if (!info.any) {
      card.className = 'status-card status-empty';
      text.textContent = '未选中任何形状';
      return;
    }
    if (info.noneRounded) {
      card.className = 'status-card status-warn';
      text.textContent = `${info.all.length} 个形状，但不是圆角矩形`;
      return;
    }
    if (info.roundedRects.length === 1) {
      card.className = 'status-card status-ok';
      text.textContent = '已选 1 个圆角矩形';
      return;
    }
    card.className = 'status-card status-ok';
    text.textContent = `已选 ${info.roundedRects.length} 个圆角矩形（多选）`;
  }

  function renderCurrentRadius(info) {
    const node = $('current-radius');
    if (!info.any || info.noneRounded) {
      node.textContent = '—';
      return;
    }
    const r = info.roundedRects;
    if (r.length === 1) {
      node.textContent = `${r[0].radiusCm.toFixed(2)} 厘米`;
      return;
    }
    // 多选：所有形状 R 角相同时显示该值，否则显示范围
    const first = r[0].radiusCm;
    const allSame = r.every((x) => Math.abs(x.radiusCm - first) < 0.005);
    if (allSame) {
      node.textContent = `${first.toFixed(2)} 厘米（多选相同）`;
    } else {
      const min = Math.min(...r.map((x) => x.radiusCm));
      const max = Math.max(...r.map((x) => x.radiusCm));
      node.textContent = `${min.toFixed(2)} ~ ${max.toFixed(2)} 厘米`;
    }
  }

  function renderLockButton(info) {
    const btn = $('lock-btn');
    const label = $('lock-label');
    const icon = $('lock-icon');
    const hint = $('lock-hint');
    if (!info.any || info.noneRounded) {
      btn.disabled = true;
      btn.classList.remove('is-locked');
      label.textContent = '锁定 R 角';
      icon.textContent = '🔓';
      hint.textContent = '需要先选中圆角矩形';
      return;
    }
    btn.disabled = false;
    const locks = RadiusCore.loadLocks();
    const r = info.roundedRects;
    const lockedCount = r.filter((s) => locks[s.id] && locks[s.id].locked).length;
    if (lockedCount === 0) {
      btn.classList.remove('is-locked');
      label.textContent = '锁定 R 角';
      icon.textContent = '🔓';
      hint.textContent = `当前：${lockedCount}/${r.length} 已锁定`;
    } else if (lockedCount === r.length) {
      btn.classList.add('is-locked');
      label.textContent = '解锁 R 角';
      icon.textContent = '🔒';
      hint.textContent = `当前：${lockedCount}/${r.length} 已锁定`;
    } else {
      btn.classList.add('is-locked');
      label.textContent = '全部锁定';
      icon.textContent = '🔒';
      hint.textContent = `当前：${lockedCount}/${r.length} 已锁定`;
    }
  }

  function onApply() {
    if (!lastSelectionInfo || lastSelectionInfo.roundedRects.length === 0) {
      showToast('请先选中至少一个圆角矩形');
      return;
    }
    const raw = $('radius-input').value;
    const cm = parseFloat(raw);
    if (!Number.isFinite(cm) || cm < 0) {
      showToast('请输入有效的 R 角数值（≥ 0）');
      return;
    }
    RadiusCore.applyRadiusToSelection(cm)
      .then((res) => {
        showToast(`已更新 ${res.updated} 个形状 R 角为 ${cm.toFixed(2)} 厘米`);
        // 如果当前选区有锁定标记，刷新锁定表里存的 radiusCm
        const locks = RadiusCore.loadLocks();
        for (const s of lastSelectionInfo.roundedRects) {
          if (locks[s.id] && locks[s.id].locked) {
            locks[s.id] = { radiusCm: cm, locked: true };
          }
        }
        RadiusCore.saveLocks(locks);
        refresh();
      })
      .catch((err) => {
        console.error(err);
        showToast('应用失败：' + (err && err.message ? err.message : err));
      });
  }

  function onToggleLock() {
    if (!lastSelectionInfo || lastSelectionInfo.roundedRects.length === 0) {
      showToast('请先选中至少一个圆角矩形');
      return;
    }
    const locks = RadiusCore.loadLocks();
    const r = lastSelectionInfo.roundedRects;
    const allLocked = r.every((s) => locks[s.id] && locks[s.id].locked);
    // 如果全锁定 → 解锁；否则 → 全部锁定
    const wantLock = !allLocked;
    RadiusCore.setLockOnSelection(wantLock)
      .then((res) => {
        showToast(
          wantLock
            ? `已锁定 ${res.updated} 个圆角矩形的 R 角绝对值`
            : `已解锁 ${res.updated} 个圆角矩形`
        );
        refresh();
      })
      .catch((err) => {
        console.error(err);
        showToast('操作失败：' + (err && err.message ? err.message : err));
      });
  }

  function onReapply() {
    RadiusCore.reapplyLocksToSelection()
      .then((res) => {
        if (res.reapplied === 0) {
          showToast('当前选区没有需要重新应用的锁定 R 角');
        } else {
          showToast(`已重新应用 ${res.reapplied} 个形状的锁定 R 角`);
        }
        refresh();
      })
      .catch((err) => {
        console.error(err);
        showToast('操作失败：' + (err && err.message ? err.message : err));
      });
  }

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
    toastTimer = setTimeout(() => el.classList.remove('show'), 2200);
  }
})();
