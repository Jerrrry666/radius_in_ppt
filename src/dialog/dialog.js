/*
 * dialog.js — R 角调整 Dialog（纯 UI，所有解析/写回走 server）
 *
 * 工作流：
 *   1. 打开 dialog → fetch /api/scan → 拿到所有圆角矩形
 *   2. 用户多选 + 输入 R 角
 *   3. 点「应用 R 角」→ fetch /api/apply → server 改 XML + 写回磁盘
 *   4. Mac PowerPoint 检测到 .pptx 改动，弹"是否重新载入"提示
 *
 * 完全在 client 端的：localStorage 锁、UI 状态。
 */

(function () {
  const $ = (id) => document.getElementById(id);

  let allShapes = [];        // 从 server 拿到的所有圆角矩形
  let selectedKeys = new Set();  // 用户选中的 "slideNum|id"

  const SERVER = 'http://localhost:3000';

  Office.onReady(() => {
    bindEvents();
    scan();
  });

  function bindEvents() {
    $('apply-btn').addEventListener('click', onApply);
    $('lock-btn').addEventListener('click', onToggleLock);
    $('reapply-btn').addEventListener('click', onReapply);
    $('select-all-btn').addEventListener('click', () => toggleAll(true));
    $('deselect-all-btn').addEventListener('click', () => toggleAll(false));
    $('rescan-btn').addEventListener('click', scan);
    $('radius-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') onApply();
    });
  }

  // ---------------- 扫描 ----------------

  async function scan() {
    setStatus('扫描中…', '正在读取 .pptx...', 'status-empty');
    $('list-count').textContent = '...';
    $('shape-list').innerHTML = '<div class="empty-list">扫描文档中…</div>';
    selectedKeys.clear();

    try {
      const res = await fetch(`${SERVER}/api/scan`);
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || `HTTP ${res.status}`);
      }
      const data = await res.json();
      allShapes = data.shapes || [];

      if (allShapes.length === 0) {
        setStatus('未找到圆角矩形', '这个文档里没有任何圆角矩形', 'status-warn');
        $('list-count').textContent = '0 个';
        $('shape-list').innerHTML = '<div class="empty-list">没有圆角矩形</div>';
        return;
      }

      setStatus('已找到圆角矩形', `${allShapes.length} 个`, 'status-ok');
      $('list-count').textContent = `${allShapes.length} 个`;
      renderShapeList(allShapes);
      toggleAll(true);  // 默认全选
    } catch (err) {
      console.error('[R 角调整] 扫描失败:', err);
      setStatus('扫描失败', err.message || String(err), 'status-warn');
      $('shape-list').innerHTML = `<div class="empty-list">扫描失败：${escapeHtml(err.message || String(err))}<br/><small style="color:#888">检查 server 是否在 3000 端口运行</small></div>`;
    }
  }

  // ---------------- 列表渲染 ----------------

  function renderShapeList(shapes) {
    const list = $('shape-list');
    list.innerHTML = '';
    for (const s of shapes) {
      const key = `${s.slideNum}|${s.id}`;
      const row = document.createElement('div');
      row.className = 'shape-row';
      row.dataset.key = key;
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = false;
      cb.addEventListener('change', () => onRowToggle(key, cb.checked, row));
      row.addEventListener('click', (e) => {
        if (e.target !== cb) {
          cb.checked = !cb.checked;
          onRowToggle(key, cb.checked, row);
        }
      });
      const name = document.createElement('div');
      name.className = 'shape-name';
      name.textContent = `S${s.slideNum} · ${s.name || '(无名)'}`;
      const meta = document.createElement('div');
      meta.className = 'shape-meta';
      meta.textContent = `${s.currentCm.toFixed(2)} cm`;
      row.appendChild(cb);
      row.appendChild(name);
      row.appendChild(meta);
      list.appendChild(row);
    }
  }

  function onRowToggle(key, checked, row) {
    if (checked) selectedKeys.add(key);
    else selectedKeys.delete(key);
    row.classList.toggle('selected', checked);
    updateCurrentRadius();
    updateLockButton();
  }

  function toggleAll(checked) {
    selectedKeys.clear();
    const rows = document.querySelectorAll('.shape-row');
    for (const row of rows) {
      const cb = row.querySelector('input[type=checkbox]');
      cb.checked = checked;
      row.classList.toggle('selected', checked);
      if (checked) selectedKeys.add(row.dataset.key);
    }
    updateCurrentRadius();
    updateLockButton();
  }

  // ---------------- UI 更新 ----------------

  function setStatus(label, text, cardClass) {
    $('status-text').textContent = text;
    $('status-card').className = 'status-card ' + cardClass;
  }

  function updateCurrentRadius() {
    const sel = allShapes.filter((s) => selectedKeys.has(`${s.slideNum}|${s.id}`));
    const node = $('current-radius');
    if (sel.length === 0) {
      node.textContent = '—';
      return;
    }
    if (sel.length === 1) {
      node.textContent = `${sel[0].currentCm.toFixed(2)} 厘米`;
      return;
    }
    const first = sel[0].currentCm;
    const allSame = sel.every((x) => Math.abs(x.currentCm - first) < 0.005);
    if (allSame) {
      node.textContent = `${first.toFixed(2)} 厘米（多选相同）`;
    } else {
      const min = Math.min(...sel.map((x) => x.currentCm));
      const max = Math.max(...sel.map((x) => x.currentCm));
      node.textContent = `${min.toFixed(2)} ~ ${max.toFixed(2)} 厘米`;
    }
  }

  function updateLockButton() {
    const btn = $('lock-btn');
    const label = $('lock-label');
    const icon = $('lock-icon');
    const hint = $('lock-hint');
    if (selectedKeys.size === 0) {
      btn.disabled = true;
      btn.classList.remove('is-locked');
      label.textContent = '锁定 R 角';
      icon.textContent = '🔓';
      hint.textContent = '请先选择圆角矩形';
      return;
    }
    btn.disabled = false;
    const locks = RadiusCore.loadLocks();
    let lockedCount = 0;
    for (const k of selectedKeys) {
      if (locks[k] && locks[k].locked) lockedCount++;
    }
    if (lockedCount === 0) {
      btn.classList.remove('is-locked');
      label.textContent = '锁定 R 角';
      icon.textContent = '🔓';
      hint.textContent = `当前：${lockedCount}/${selectedKeys.size} 已锁定`;
    } else if (lockedCount === selectedKeys.size) {
      btn.classList.add('is-locked');
      label.textContent = '解锁 R 角';
      icon.textContent = '🔒';
      hint.textContent = `当前：${lockedCount}/${selectedKeys.size} 已锁定`;
    } else {
      btn.classList.add('is-locked');
      label.textContent = '全部锁定';
      icon.textContent = '🔒';
      hint.textContent = `当前：${lockedCount}/${selectedKeys.size} 已锁定`;
    }
  }

  // ---------------- 操作 ----------------

  async function onApply() {
    if (selectedKeys.size === 0) {
      showToast('请先选择至少一个圆角矩形');
      return;
    }
    const cm = parseFloat($('radius-input').value);
    if (!Number.isFinite(cm) || cm < 0) {
      showToast('请输入有效的 R 角数值（≥ 0）');
      return;
    }

    // 构造 items
    const items = [];
    for (const k of selectedKeys) {
      const [slideNum, id] = k.split('|');
      items.push({ slideNum: parseInt(slideNum, 10), id, cm });
    }

    try {
      const res = await fetch(`${SERVER}/api/apply`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ shapes: items }),
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || `HTTP ${res.status}`);
      }
      const result = await res.json();

      // 更新 localStorage 锁定表
      const locks = RadiusCore.loadLocks();
      for (const it of items) {
        const k = `${it.slideNum}|${it.id}`;
        if (locks[k] && locks[k].locked) {
          locks[k] = { radiusCm: cm, locked: true };
        }
      }
      RadiusCore.saveLocks(locks);

      // 显示操作面板（因为 Mac PowerPoint 不会自动弹"文件已修改"）
      showActionPanel(
        `✅ 已更新 ${result.modified} 个圆角矩形为 ${cm.toFixed(2)} 厘米`,
        `Mac PowerPoint 不会自动检测文件被外部修改。\n请在 PowerPoint 里：\n1. 点顶部菜单「文件」\n2. 选「关闭」\n3. 再点「文件 → 打开」找到 ${result.path} 重开\n\n或者直接 Cmd+W 关闭再点文件名重开。`
      );

      // 重新扫描
      await scan();
    } catch (err) {
      console.error('[R 角调整] 应用失败:', err);
      showToast('应用失败：' + (err.message || err));
    }
  }

  function onToggleLock() {
    if (selectedKeys.size === 0) {
      showToast('请先选择圆角矩形');
      return;
    }
    const locks = RadiusCore.loadLocks();
    const allLocked = [...selectedKeys].every((k) => locks[k] && locks[k].locked);
    const wantLock = !allLocked;
    let n = 0;
    for (const k of selectedKeys) {
      if (wantLock) {
        const s = allShapes.find((x) => `${x.slideNum}|${x.id}` === k);
        if (!s) continue;
        locks[k] = { radiusCm: s.currentCm, locked: true };
      } else {
        if (locks[k]) delete locks[k];
      }
      n++;
    }
    RadiusCore.saveLocks(locks);
    showToast(wantLock ? `已锁定 ${n} 个圆角矩形` : `已解锁 ${n} 个圆角矩形`);
    updateLockButton();
  }

  function onReapply() {
    if (selectedKeys.size === 0) {
      showToast('请先选择圆角矩形');
      return;
    }
    const locks = RadiusCore.loadLocks();
    let applied = 0;
    let missing = 0;
    for (const k of selectedKeys) {
      const e = locks[k];
      if (!e || !e.locked) {
        missing++;
        continue;
      }
      const s = allShapes.find((x) => `${x.slideNum}|${x.id}` === k);
      if (!s) {
        missing++;
        continue;
      }
      $('radius-input').value = e.radiusCm.toFixed(2);
      applied++;
    }
    if (applied === 0) {
      showToast('选中的圆角矩形都没有锁定');
    } else {
      showToast(`${applied} 个有锁定，已填入 R 角值。请点「应用 R 角」写入。`);
    }
  }

  // ---------------- Toast / 操作面板 ----------------

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
    toastTimer = setTimeout(() => el.classList.remove('show'), 3000);
  }

  /**
   * 显示一个操作面板（替代 toast，用于需要用户后续操作的情况）
   * 包含一个"知道了"按钮关闭
   */
  function showActionPanel(title, message) {
    // 移除旧面板
    const old = document.querySelector('.action-panel');
    if (old) old.remove();

    const panel = document.createElement('div');
    panel.className = 'action-panel';
    panel.innerHTML = `
      <div class="action-panel-title">${escapeHtml(title)}</div>
      <div class="action-panel-msg">${escapeHtml(message).replace(/\n/g, '<br/>')}</div>
      <div class="action-panel-buttons">
        <button class="btn btn-secondary" id="action-ok">知道了</button>
      </div>
    `;
    document.body.appendChild(panel);
    $('action-ok').addEventListener('click', () => panel.remove());
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
})();
