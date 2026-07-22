/*
 * dialog.js — R 角调整 Dialog
 *
 * 路线（绕开 Mac PowerPoint 不支持 getSelectedDataAsync(Shape)）：
 *   1. 用 Office.context.document.getFileAsync("compressed") 拿 .pptx 压缩包
 *   2. JSZip 解压
 *   3. 解析所有 slide*.xml，列出所有圆角矩形
 *   4. 用户多选 + 输入 R 角
 *   5. 改 XML + 重新打包
 *   6. setFileAsync 写回
 *
 * "锁定"用 localStorage 存 { "slideNum|shapeId": { radiusCm, locked } }。
 */

(function () {
  const $ = (id) => document.getElementById(id);

  // 当前文档的所有圆角矩形
  let allShapes = [];
  // 用户选中的 key set: "slideNum|shapeId"
  let selectedKeys = new Set();
  // 当前 .pptx 压缩包（JSZip 实例）
  let currentZip = null;

  Office.onReady(() => {
    bindEvents();
    scanDocument();
  });

  function bindEvents() {
    $('apply-btn').addEventListener('click', onApply);
    $('lock-btn').addEventListener('click', onToggleLock);
    $('reapply-btn').addEventListener('click', onReapply);
    $('select-all-btn').addEventListener('click', () => toggleAll(true));
    $('deselect-all-btn').addEventListener('click', () => toggleAll(false));
    $('rescan-btn').addEventListener('click', scanDocument);
    $('radius-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') onApply();
    });
  }

  // ---------------- 扫描文档 ----------------

  async function scanDocument() {
    setStatus('扫描中…', '正在读取 .pptx...', 'status-empty');
    $('list-count').textContent = '...';
    $('shape-list').innerHTML = '<div class="empty-list">扫描文档中…</div>';
    selectedKeys.clear();

    try {
      const zip = await getPptxAsZip();
      currentZip = zip;
      const shapes = RadiusCore.parseRoundedRects(zip);
      allShapes = shapes;

      if (shapes.length === 0) {
        setStatus('未找到圆角矩形', '这个文档里没有任何圆角矩形', 'status-warn');
        $('list-count').textContent = '0 个';
        $('shape-list').innerHTML = '<div class="empty-list">没有圆角矩形</div>';
        return;
      }

      setStatus('已找到圆角矩形', `${shapes.length} 个`, 'status-ok');
      $('list-count').textContent = `${shapes.length} 个`;
      renderShapeList(shapes);
      updateCurrentRadius();
      updateLockButton();

      // 默认全选
      toggleAll(true);
    } catch (err) {
      console.error('[R 角调整] 扫描失败:', err);
      setStatus('扫描失败', err.message || String(err), 'status-warn');
      $('shape-list').innerHTML = `<div class="empty-list">扫描失败：${escapeHtml(err.message || String(err))}</div>`;
    }
  }

  /**
   * 用 Office.js 拿 .pptx 压缩包，转成 JSZip 实例。
   * Mac PowerPoint 上 getFileAsync("compressed") 是支持的。
   */
  async function getPptxAsZip() {
    if (!Office.context.document.getFileAsync) {
      throw new Error('当前 Office 版本不支持 getFileAsync');
    }
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 4 * 1024 * 1024 },  // 4MB
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error('getFileAsync 失败: ' + (result.error?.message || 'unknown')));
            return;
          }
          const file = result.value;
          const slices = file.sliceCount;
          const slicePromises = [];
          for (let i = 0; i < slices; i++) {
            slicePromises.push(new Promise((res, rej) => {
              file.getSliceAsync(i, (sliceResult) => {
                if (sliceResult.status === Office.AsyncResultStatus.Failed) {
                  rej(new Error('getSliceAsync 失败: ' + (sliceResult.error?.message || 'unknown')));
                  return;
                }
                res(sliceResult.value.data);  // base64 string
              });
            }));
          }
          Promise.all(slicePromises).then((dataArray) => {
            file.closeAsync();
            // 拼接 base64
            const base64 = dataArray.join('');
            // base64 -> Uint8Array
            const binStr = atob(base64);
            const bytes = new Uint8Array(binStr.length);
            for (let i = 0; i < binStr.length; i++) {
              bytes[i] = binStr.charCodeAt(i);
            }
            // JSZip 解析
            JSZip.loadAsync(bytes).then(resolve).catch(reject);
          }).catch(reject);
        }
      );
    });
  }

  // ---------------- 列表渲染 ----------------

  function renderShapeList(shapes) {
    const list = $('shape-list');
    list.innerHTML = '';
    for (const s of shapes) {
      const key = RadiusCore.getLockKey(s.slideNum, s.id);
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
    const sel = allShapes.filter((s) => selectedKeys.has(RadiusCore.getLockKey(s.slideNum, s.id)));
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
    const lockedCount = [...selectedKeys].filter((k) => locks[k] && locks[k].locked).length;
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
    if (!currentZip) {
      showToast('请先点击「重新扫描」');
      return;
    }

    try {
      // 1. 改所有选中的 shape（在 zip 里改 XML）
      const targets = allShapes.filter((s) => selectedKeys.has(RadiusCore.getLockKey(s.slideNum, s.id)));
      // 按 slideFile 分组
      const byFile = {};
      for (const t of targets) {
        if (!byFile[t.slideFile]) byFile[t.slideFile] = [];
        byFile[t.slideFile].push(t);
      }
      for (const [file, items] of Object.entries(byFile)) {
        let xml = await currentZip.file(file).async('string');
        for (const it of items) {
          const newVal = RadiusCore.cmToVal(cm, it.shortSideEmu);
          xml = RadiusCore.modifyShapeAdj(xml, it.id, newVal);
        }
        currentZip.file(file, xml);
      }

      // 2. 写回 Office
      await writeZipBack(currentZip);

      // 3. 更新本地锁定表（如果选中的有锁定）
      const locks = RadiusCore.loadLocks();
      for (const t of targets) {
        const k = RadiusCore.getLockKey(t.slideNum, t.id);
        if (locks[k] && locks[k].locked) {
          locks[k] = { radiusCm: cm, locked: true };
        }
      }
      RadiusCore.saveLocks(locks);

      showToast(`已更新 ${targets.length} 个圆角矩形为 ${cm.toFixed(2)} 厘米`);

      // 4. 重新扫描刷新 UI
      await scanDocument();
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
        // 从 allShapes 拿当前 R 角
        const s = allShapes.find((x) => RadiusCore.getLockKey(x.slideNum, x.id) === k);
        if (!s) continue;
        locks[k] = { radiusCm: s.currentCm, locked: true };
      } else {
        if (locks[k]) {
          delete locks[k];
        }
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
      const s = allShapes.find((x) => RadiusCore.getLockKey(x.slideNum, x.id) === k);
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

  // ---------------- 写回 .pptx ----------------

  /**
   * 把改完的 JSZip 重新打包成 base64，POST 到本地 server 写回磁盘。
   * Mac PowerPoint 检测到 .pptx 文件被外部修改会弹"是否重新载入"提示。
   */
  async function writeZipBack(zip) {
    const b64 = await zip.generateAsync({
      type: 'base64',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
    // 拿 .pptx 路径（从 server 的 /api/pptx-path 拿）
    const pathRes = await fetch('http://localhost:3000/api/pptx-path');
    if (!pathRes.ok) {
      throw new Error('拿 .pptx 路径失败（需要先双击 R 角调整.app）');
    }
    const pathData = await pathRes.json();
    const pptxPath = pathData.path;
    if (!pptxPath) throw new Error('没拿到 .pptx 路径');

    // POST 写回
    const saveRes = await fetch(`http://localhost:3000/api/save-pptx?path=${encodeURIComponent(pptxPath)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: b64,
    });
    if (!saveRes.ok) {
      const errText = await saveRes.text();
      throw new Error('写回失败: ' + errText);
    }
    const saveData = await saveRes.json();
    return saveData;
  }

  // ---------------- Toast ----------------

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

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
})();
