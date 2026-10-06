/* Office UI transport only; no shape, radius, lock or layout rules. */
(function (root) {
  function createOfficeUiDriver(office) {
    return {
      associate(actions) {
        Object.entries(actions).forEach(([name, action]) => {
          root[name] = action;
          if (office.actions && office.actions.associate) office.actions.associate(name, action);
        });
      },
      showPane() { return office.addin.showAsTaskpane(); },
      hidePane() { return office.addin.hide(); },
      showDialog(url, onMessage, onClose) {
        return new Promise((resolve, reject) => {
          office.context.ui.displayDialogAsync(url, { height: 38, width: 32, displayInIframe: false }, (result) => {
            if (result.status !== office.AsyncResultStatus.Succeeded) {
              reject(new Error(result.error && result.error.message || 'Could not open dialog')); return;
            }
            const dialog = result.value;
            dialog.addEventHandler(office.EventType.DialogMessageReceived, onMessage);
            dialog.addEventHandler(office.EventType.DialogEventReceived, onClose);
            resolve({ close: () => dialog.close() });
          });
        });
      },
    };
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = { createOfficeUiDriver };
  if (root.Office) root.OfficeUi = createOfficeUiDriver(root.Office);
})(typeof window !== 'undefined' ? window : globalThis);
