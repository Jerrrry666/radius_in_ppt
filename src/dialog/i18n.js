// i18n.js - lightweight i18n module for the add-in
//
// Usage:
//   1. Include i18n-data.js BEFORE this file (it defines window.I18N_DATA = { zh: {...}, en: {...} })
//   2. In HTML, mark translatable nodes with data-i18n / data-i18n-placeholder /
//      data-i18n-title / data-i18n-aria-label
//   3. Call i18n.applyAll() after DOM ready (or call i18n.t('key') imperatively)
//
// Fallback chain: currentLang -> en -> key (literal).
// params: pass {name: value} to substitute {name} placeholders.

(function (global) {
  'use strict';

  var DATA = global.I18N_DATA || { zh: {}, en: {} };

  function getLang() {
    var navLang = (global.navigator && (global.navigator.language || global.navigator.userLanguage)) || 'en';
    var lang = (navLang || 'en').toLowerCase().split('-')[0];
    // supported: zh, en. Everything else -> en.
    return lang === 'zh' ? 'zh' : 'en';
  }

  var currentLang = getLang();
  var messages = DATA[currentLang] || DATA.en || {};
  var fallback = DATA.en || {};

  function t(key, params) {
    var s = messages[key];
    if (s === undefined) s = fallback[key];
    if (s === undefined) s = key; // last resort: return the key
    if (params) {
      Object.keys(params).forEach(function (k) {
        s = s.replace(new RegExp('\\{' + k + '\\}', 'g'), params[k]);
      });
    }
    return s;
  }

  function applyNode(el) {
    var key;
    key = el.getAttribute('data-i18n');
    if (key) el.textContent = t(key);
    key = el.getAttribute('data-i18n-placeholder');
    if (key) el.setAttribute('placeholder', t(key));
    key = el.getAttribute('data-i18n-title');
    if (key) el.setAttribute('title', t(key));
    key = el.getAttribute('data-i18n-aria-label');
    if (key) el.setAttribute('aria-label', t(key));
  }

  function applyAll(root) {
    root = root || document;
    var sel = '[data-i18n],[data-i18n-placeholder],[data-i18n-title],[data-i18n-aria-label]';
    var nodes = root.querySelectorAll(sel);
    for (var i = 0; i < nodes.length; i++) applyNode(nodes[i]);
  }

  global.i18n = {
    t: t,
    applyAll: applyAll,
    getLang: function () { return currentLang; }
  };
})(typeof window !== 'undefined' ? window : this);
