(function () {
  const params = new URLSearchParams(location.search);
  const value = document.getElementById('value');
  const unit = document.getElementById('unit');
  const initial = Number(params.get('value'));
  value.value = Number.isFinite(initial) && initial >= 0 ? initial : 0.3;
  unit.value = params.get('unit') === '%' ? '%' : 'cm';
  let english = false;
  Office.onReady(() => {
    english = !String(Office.context.displayLanguage || navigator.language).toLowerCase().startsWith('zh');
    if (english) {
      document.documentElement.lang = 'en';
      document.getElementById('title').textContent = 'Set R-corner';
      document.getElementById('label').textContent = 'Corner radius';
      document.getElementById('hint').textContent = 'Clamped to half the short side. Protected shapes block the entire selection.';
      document.getElementById('apply').textContent = 'Apply';
      document.getElementById('cancel').textContent = 'Cancel';
    }
    document.getElementById('apply').disabled = false;
    document.getElementById('cancel').disabled = false;
  });
  document.getElementById('form').addEventListener('submit', (event) => {
    event.preventDefault();
    const number = value.value.trim() === '' ? NaN : Number(value.value);
    if (!Number.isFinite(number) || number < 0) {
      document.getElementById('error').textContent = english ? 'Enter a nonnegative number.' : '请输入非负数。'; return;
    }
    Office.context.ui.messageParent(JSON.stringify({ value: number, unit: unit.value }));
  });
  document.getElementById('cancel').addEventListener('click', () => Office.context.ui.messageParent(JSON.stringify({ cancel: true })));
})();
