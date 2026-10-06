/* Ribbon UI routes into the existing pane -> core -> driver path. */
(function (root) {
  function createCommandController(pane, ui, origin) {
    let tail = Promise.resolve();
    let activeDialog = null;
    async function input() {
      await pane.execute('read');
      const token = pane.selectionToken();
      const initial = pane.inputValue();
      const url = origin + '/src/ribbon/radius-input.html?value=' + encodeURIComponent(initial.value) + '&unit=' + encodeURIComponent(initial.unit);
      const response = await new Promise((resolve, reject) => {
        let settled = false;
        const finish = (value, error) => {
          if (settled) return;
          settled = true;
          if (activeDialog) { activeDialog.close(); activeDialog = null; }
          if (error) reject(error); else resolve(value);
        };
        ui.showDialog(url, (event) => {
          try {
            if (event.origin && event.origin !== origin) throw new Error('Unexpected dialog origin');
            if (typeof event.message !== 'string' || event.message.length > 4096) throw new Error('Invalid dialog message');
            const value = JSON.parse(event.message);
            if (value.cancel === true) return finish(null);
            if (!Number.isFinite(value.value) || value.value < 0 || !['cm', '%'].includes(value.unit)) throw new Error('Invalid radius');
            finish(value);
          } catch (error) { finish(null, error); }
        }, () => finish(null)).then((dialog) => {
          if (settled) dialog.close(); else activeDialog = dialog;
        }, (error) => finish(null, error));
      });
      if (!response) return;
      if (pane.selectionToken() !== token) throw new Error('选区已改变，请重新设置R角 / Selection changed; try again.');
      await pane.execute('apply', response);
    }
    const routes = {
      RadiusSet: input,
      RadiusPreset01: () => pane.execute('apply', { value: 0.1, unit: 'cm' }),
      RadiusPreset03: () => pane.execute('apply', { value: 0.3, unit: 'cm' }),
      RadiusPreset05: () => pane.execute('apply', { value: 0.5, unit: 'cm' }),
      RadiusPresetZero: () => pane.execute('apply', { value: 0, unit: 'cm' }),
      RadiusLock: () => pane.execute('lock'),
      RadiusStrict: () => pane.execute('strict'),
      RadiusReapply: () => pane.execute('reapply'),
      RadiusPick: () => pane.execute('pick'),
      RadiusBrush: () => pane.execute('brush'),
      RadiusHide: () => ui.hidePane(),
    };
    const actions = {};
    Object.entries(routes).forEach(([name, route]) => {
      actions[name] = (event) => {
        const work = tail.then(async () => { await pane.ready; await route(); });
        const done = work.catch(async (error) => {
          console.error('[ribbon]', name, error.message, error.stack);
          pane.notify(error.message);
          try { await ui.showPane(); } catch (showError) { console.error('[ribbon/showPane]', showError.message); }
        }).finally(() => { if (event && typeof event.completed === 'function') event.completed(); });
        tail = done.catch(() => {});
        return done;
      };
    });
    return actions;
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = { createCommandController };
  if (root.RadiusPaneActions && root.OfficeUi) {
    root.PptDriver.onReady(() => {
      root.OfficeUi.associate(createCommandController(root.RadiusPaneActions, root.OfficeUi, root.location.origin));
    });
  }
})(typeof window !== 'undefined' ? window : globalThis);
