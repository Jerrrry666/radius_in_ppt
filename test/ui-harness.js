const fs = require('fs');
const path = require('path');
const vm = require('vm');
const core = require('../src/lib/radius-core');
const { createHostQueue } = require('../src/lib/host-queue');
const { createHarness } = require('./test-harness');

function createUiHarness(shapes, selected) {
  const host = createHarness({ shapes });
  let roots = selected || shapes;
  const nodes = new Map();
  const timers = new Map();
  let nextTimer = 1;
  class Element {
    constructor() {
      this.children = []; this.dataset = {}; this.style = {}; this.value = '';
      this.textContent = ''; this.className = ''; this.checked = false;
      this.classList = { add() {}, remove() {}, toggle() {} };
    }
    set innerHTML(value) { this.html = value; this.children = []; }
    get innerHTML() { return this.html || ''; }
    appendChild(node) { this.children.push(node); return node; }
    querySelector(selector) { return this.children.find((node) => node.className === selector.slice(1)) || null; }
    querySelectorAll() { return []; }
    addEventListener() {}
    setAttribute() {}
  }
  const document = {
    getElementById(id) { if (!nodes.has(id)) nodes.set(id, new Element()); return nodes.get(id); },
    createElement() { return new Element(); },
    createTextNode(text) { const node = new Element(); node.textContent = text; return node; },
    querySelector() { return null; }, querySelectorAll() { return []; }, addEventListener() {},
  };
  const logs = [];
  let running = 0, maxRunning = 0;
  let runOverride;
  const queue = createHostQueue(async (callback) => {
    running++; maxRunning = Math.max(maxRunning, running);
    try {
      if (runOverride) return await runOverride(callback);
      const result = await callback({});
      await host.driver.sync(); // model the final implicit sync too
      return result;
    } finally { running--; }
  });
  host.driver.selectedShapes = () => ({ items: roots });
  host.slide.id = 'test-slide';
  host.driver.slideById = (id) => {
    if (id !== host.slide.id) throw new Error('Unknown slide: ' + id);
    return host.slide;
  };
  host.driver.loadSelectionIdentity = async () => ({ slideId: 'test-slide', shapeIds: roots.map((s) => s.id) });
  const select = host.driver.selectShapes;
  host.driver.selectShapes = (slide, ids) => {
    select(slide, ids);
    roots = ids.map((id) => slide.shapes.items.find((shape) => shape.id === id)).filter(Boolean);
  };
  const i18n = { t: (key, values) => key + (values ? JSON.stringify(values) : ''), getLang: () => 'zh', applyAll() {} };
  const context = vm.createContext({
    document, i18n, console: { log: (...args) => logs.push(args), warn: (...args) => logs.push(args), error: (...args) => logs.push(args) },
    Date, Map, Set, Promise, Number, JSON,
    setTimeout(callback, delay) { const id = nextTimer++; timers.set(id, { callback, delay }); return id; },
    clearTimeout(id) { timers.delete(id); },
    setInterval(callback, delay) { const id = nextTimer++; timers.set(id, { callback, delay, interval: true }); return id; },
    clearInterval(id) { timers.delete(id); },
    Office: { onReady() {}, context: { document: { addHandlerAsync() {} } } },
    window: { RadiusCore: core, i18n, PptDriver: { createDriver: () => host.driver, run: queue.run, isBusy: () => queue.busy, onReady() {}, onSelectionChanged() {} } },
  });
  const exportForTests = `globalThis.ui = {
    refreshSelection, monitorTick, stopLockMonitor, startLockMonitor, handleSelectionChanged,
    loadLocksViaTags, saveLocksViaTags, updateLockTagForShape,
    loadLayoutTagsViaTags, saveLayoutTags, deleteLayoutTags,
    applyLayoutToChildren, syncLayoutChildrenRIfNeeded, applyLayoutFromUI,
    onApply, onToggleLock, onToggleStrict, onReapply,
    applyPipetteToSelection, runDriverSmokeTest,
    setBrush(source, sync) { pipetteSource = source; pipetteSyncStrict = sync; pipetteState = 'brushing'; },
    renderShapeList, renderCurrentRadius, scheduleLayoutApply,
    setState(shapes, layout) { selectedShapes = shapes; currentLayout = layout || null; },
    invalidate() { selectionEpoch++; stopLockMonitor(); },
    ignoreSelection(identity) { ignoredRestoredSelection = identity; layoutSelectionIgnoreUntil = Date.now() + 300; },
    getState() { return { selectedShapes, currentLayout, lockMonitor, selectionEpoch, monitorInFlight }; }
  };`;
  const source = fs.readFileSync(path.join(__dirname, '../src/dialog/dialog.js'), 'utf8');
  vm.runInContext(source.replace(/\}\)\(\);\s*$/, exportForTests + '\n})();'), context);
  return { ...host, ui: context.ui, nodes, timers, logs, queue,
    select(shapes) { roots = shapes; },
    failRun(fn) { runOverride = fn; },
    get maxRunning() { return maxRunning; },
    async fireTimer(delay) {
      const entry = Array.from(timers).find(([, timer]) => timer.delay === delay && !timer.interval);
      if (!entry) throw new Error('No pending timer: ' + delay);
      timers.delete(entry[0]);
      await entry[1].callback();
    },
  };
}
module.exports = { createUiHarness };
