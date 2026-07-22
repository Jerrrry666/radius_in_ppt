/*
 * commands.js
 * 注册 ribbon 按钮回调。这里只做一件事：打开 R 角调整 Dialog。
 */

// 单一全局 Dialog 引用，避免重复打开
let radiusDialog = null;

Office.onReady(() => {
  // Office.js 加载完成后函数就绪。ribbon 按钮触发 openRadiusDialog。
});

/**
 * 由 manifest.xml 中 CustomTab Button 的 <FunctionName> 引用。
 * 在新窗口中打开 R 角调整面板。
 * @param {Office.AddinCommands.Event} event
 */
function openRadiusDialog(event) {
  // 关闭已打开的 dialog（避免重复）
  if (radiusDialog) {
    try {
      radiusDialog.close();
    } catch (_) {
      // 忽略：可能已经被用户关闭
    }
    radiusDialog = null;
  }

  const dialogUrl = new URL('../dialog/dialog.html', window.location.href).toString();

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {
      displayInIframe: true,
      width: 360,
      height: 420,
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error('[R 角调整] 打开 dialog 失败:', asyncResult.error);
      } else {
        radiusDialog = asyncResult.value;
        // Dialog 关闭时清理引用
        radiusDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          radiusDialog = null;
        });
      }
      // 通知 Office 命令执行完成
      event.completed();
    }
  );
}

// 注册到全局（Office Add-in 要求命令函数挂在 window 上）
window.openRadiusDialog = openRadiusDialog;
