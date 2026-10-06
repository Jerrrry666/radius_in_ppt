const assert = require('assert/strict');
const { createTestRunner, createHarness } = require('./test-harness');
const { makeFixtureShape, makeStandardFixture, PT_PER_CM } = require('./fixtures');
const { createDriver } = require('../src/lib/ppt-driver');
const { createHostQueue } = require('../src/lib/host-queue');
const { parseSlideGeometry, readSlideGeometry } = require('../src/lib/pptx-geometry');
const { zipSync, strToU8 } = require('../src/lib/vendor/fflate-0.8.2');
const core = require('../src/lib/radius-core');
const { createUiHarness } = require('./ui-harness');
const { createStaticServer } = require('../tools/serve');
const http = require('http');

const suite = createTestRunner();
const test = suite.test;
const deferred = () => { let resolve; const promise = new Promise((r) => { resolve = r; }); return { promise, resolve }; };
const turn = () => new Promise((resolve) => setImmediate(resolve));
const geometryXml = '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Round"/></p:nvSpPr><p:spPr><a:prstGeom prst="roundRect"/></p:spPr></p:sp>' +
  '<p:grpSp><p:sp><p:nvSpPr><p:cNvPr id="3"/></p:nvSpPr><p:spPr><a:prstGeom prst="plus"/></p:spPr></p:sp></p:grpSp>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="4"/></p:nvSpPr><p:spPr><a:custGeom/></p:spPr></p:sp>' +
  '</p:spTree></p:cSld></p:sld>';

test('OOXML distinguishes roundRect, adjustable cross and custom geometry in groups', () => {
  const map = parseSlideGeometry(geometryXml);
  assert.equal(map.get('2'), 'roundRect');
  assert.equal(map.get('3'), 'plus');
  assert.equal(map.get('4'), null);
  const d = createDriver({}, { shapeKinds: map });
  for (const [id, expected] of [['2', true], ['3', false], ['4', false], ['5', false]]) {
    assert.equal(d.isRoundRect({ id, adjustments: { count: 1 } }), expected);
  }
});

test('PPTX decoder reads compressed slide XML and excludes media', () => {
  const bytes = zipSync({ 'ppt/slides/slide1.xml': strToU8(geometryXml), 'ppt/media/image1.png': new Uint8Array([1]) });
  const map = readSlideGeometry(Buffer.from(bytes).toString('base64'));
  assert.equal(map.get('2'), 'roundRect');
  assert.equal(map.get('3'), 'plus');
  assert.throws(() => readSlideGeometry('invalid'), /./);
});

test('Real driver exports one slide once per operation and refreshes on demand', async () => {
  const bytes = zipSync({ 'ppt/slides/slide1.xml': strToU8(geometryXml) });
  let exports = 0;
  const d = createDriver({ presentation: { getSelectedSlides: () => ({ getItemAt: () => ({ exportAsBase64() {
    exports++; return { value: Buffer.from(bytes).toString('base64') };
  } }) }) }, sync: async () => {} });
  await d.loadShapeKinds(); await d.loadShapeKinds();
  assert.equal(exports, 1);
  await d.loadShapeKinds(true);
  assert.equal(exports, 2);
});

test('Mac adjustment reads load values before obtaining a fresh result snapshot', async () => {
  let queuedLoad = false, loaded = false, value = 0.2;
  const shape = { adjustments: { count: 1,
    load(fields) { assert.equal(fields, 'items/value'); queuedLoad = true; },
    get() {
      const snapshot = loaded ? value : undefined;
      return { get value() {
        if (snapshot == null) throw new Error('Adjustment value was not loaded');
        return snapshot;
      } };
    },
  } };
  const d = createDriver({ sync: async () => { if (queuedLoad) { loaded = true; queuedLoad = false; } } });
  assert.equal(await d.readAdjFraction(shape), 0.2);
  value = 0; loaded = false;
  assert.equal(await d.readAdjFraction(shape), 0);
});

test('Actual brush UI cannot clear protection even with syncStrict selected', async () => {
  for (const sourceStrict of [true, false]) {
    const a = makeFixtureShape({ id: 'a', adjFraction: 0.1, tags: { radiusLockStrict_v1: '1', radiusLock_v1: '0.2' } });
    const b = makeFixtureShape({ id: 'b' });
    const group = { id: 'group', _isGroup: true, _groupShapes: [a, b], _tags: {}, name: 'Group' };
    const h = createUiHarness([group]);
    await h.ui.refreshSelection();
    // A stale memory snapshot must still be stopped by the live tag preflight.
    h.ui.getState().selectedShapes[0].strictLocked = false;
    h.ui.setBrush({ cm: 0.8, sourceStrict }, true);
    h.calls.length = 0;
    await h.ui.applyPipetteToSelection();
    assert.equal(a._tags.radiusLockStrict_v1, '1');
    assert.equal(a._adjFraction, 0.1); assert.equal(b._adjFraction, 0);
    for (const method of ['deleteTag', 'addTag', 'setAdjFraction', 'ungroupShapeGroup']) h.assertNotCalled(method);
  }
});

test('Brush selection events refresh the target before checking source protection', async () => {
  const source = makeFixtureShape({ id: 'source', tags: { radiusLock_v1: '1.8', radiusLockStrict_v1: '1' } });
  const target = makeFixtureShape({ id: 'target', heightCm: 2, adjFraction: 0.1 });
  const protectedTarget = makeFixtureShape({ id: 'protected', adjFraction: 0.2,
    tags: { radiusLock_v1: '0.4', radiusLockStrict_v1: '1' } });
  const h = createUiHarness([source, target, protectedTarget], [source]);
  await h.ui.refreshSelection();
  h.ui.setBrush({ cm: 1.8, sourceStrict: true }, true);
  h.select([target]);
  await h.ui.handleSelectionChanged();
  assert.equal(target._adjFraction, 0.5);
  assert.equal(target._tags.radiusLock_v1, '1');
  assert.equal(target._tags.radiusLockStrict_v1, '1');
  assert.equal(h.ui.getState().selectedShapes[0].id, 'target');
  h.select([protectedTarget]); h.calls.length = 0;
  await h.ui.handleSelectionChanged();
  assert.equal(h.ui.getState().selectedShapes[0].id, 'protected');
  assert.equal(protectedTarget._adjFraction, 0.2);
  for (const method of ['setAdjFraction', 'addTag', 'deleteTag']) h.assertNotCalled(method);
});

test('Driver smoke test refuses protected shapes without touching their radius or tags', async () => {
  for (const tags of [{ radiusLockStrict_v1: '1' }, { radiusLock_v1: '0.2' }]) {
    const shape = makeFixtureShape({ id: 'protected', adjFraction: 0.1, tags });
    const h = createUiHarness([shape]);
    h.calls.length = 0;
    await h.ui.runDriverSmokeTest();
    for (const method of ['setAdjFraction', 'setBox', 'addTag', 'deleteTag']) h.assertNotCalled(method);
    assert.equal(shape._adjFraction, 0.1);
  }
});

test('Repeated smoke clicks cannot capture temporary values or overlap cleanup', async () => {
  const shape = makeFixtureShape({ id: 'smoke', adjFraction: 0.1 });
  const h = createUiHarness([shape]); h.slide.id = 'test-slide';
  const gate = deferred();
  const sync = h.driver.sync;
  let paused = false;
  h.driver.sync = async () => {
    if (!paused && h.calls.some((call) => call.method === 'setBox')) {
      paused = true;
      await gate.promise;
    }
    await sync();
  };
  const first = h.ui.runDriverSmokeTest();
  while (!paused) await turn();
  assert.equal(h.nodes.get('smoke-test-btn').disabled, true);
  await h.ui.runDriverSmokeTest();
  assert.equal(h.calls.filter((call) => call.method === 'setBox').length, 1);
  gate.resolve();
  await first;
  assert.equal(shape._adjFraction, 0.1); assert.equal(shape.left, 0);
  assert.equal(h.calls.filter((call) => call.method === 'setBox').length, 2);
  assert.equal(h.nodes.get('smoke-test-btn').disabled, false);
  assert.equal(h.logs.filter((args) => String(args[0]).includes('14/14 passed')).length, 1);
});

test('Smoke cleanup restores the original slide even after a switch to a duplicate shape ID', async () => {
  const shape = makeFixtureShape({ id: 'duplicate', adjFraction: 0.1, tags: { DRIVER_SMOKE_TEST_V1: 'original' } });
  const other = makeFixtureShape({ id: 'duplicate', adjFraction: 0.3, leftCm: 9 });
  const h = createUiHarness([shape]); h.slide.id = 'test-slide';
  const otherSlide = { id: 'other-slide', shapes: { items: [other] } };
  h.driver.slideById = (id) => id === h.slide.id ? h.slide : otherSlide;
  const sync = h.driver.sync; let switched = false;
  h.driver.sync = async () => {
    if (!switched && h.calls.some((call) => call.method === 'setBox')) {
      switched = true; h.select([other]); h.ui.invalidate();
      h.driver.activeSlide = () => otherSlide;
      h.driver.loadSelectionIdentity = async () => ({ slideId: otherSlide.id, shapeIds: [other.id] });
    }
    await sync();
  };
  await h.ui.runDriverSmokeTest();
  assert.equal(shape._adjFraction, 0.1); assert.equal(shape.left, 0);
  assert.equal(shape._tags.DRIVER_SMOKE_TEST_V1, 'original');
  assert.equal(other._adjFraction, 0.3); assert.equal(other.left, 9 * PT_PER_CM);
  assert.deepEqual(other._tags, {});
});

test('Host tag failures fail closed; confirmed absence returns null', async () => {
  let writes = 0;
  const shape = { id: '2', width: 100, height: 100,
    adjustments: { count: 1, set() { writes++; } },
    tags: { load() {}, items: [] } };
  const d = createDriver({ sync: async () => { throw new Error('GeneralException'); } }, { shapeKinds: new Map([['2', 'roundRect']]) });
  const r = await core.writeRadius(d, shape, 0.3);
  assert.equal(r.ok, false); assert.equal(writes, 0); assert.match(r.error, /GeneralException/);
  const working = createDriver({ sync: async () => {} });
  assert.equal(await working.readTag(shape, 'missing'), null);
  Object.defineProperty(shape.tags, 'items', { get() { throw new Error('PropertyNotLoaded'); } });
  await assert.rejects(working.readTag(shape, 'missing'), /PropertyNotLoaded/);
});

test('Tag write errors are returned rather than reported as success', async () => {
  const r = await core.writeLockState({ addTag() { throw new Error('Tag failed'); } }, {}, { lockedCm: 0, isStrict: true });
  assert.equal(r.ok, false); assert.equal(r.error, 'Tag failed');
});

test('Strict child outside selection blocks all layout writes before ungrouping', async () => {
  const f = makeStandardFixture();
  const child = f.layoutChildren[0];
  child._tags.radiusLockStrict_v1 = '1'; child._tags.radiusLock_v1 = '0.15';
  const group = { id: 'group', _isGroup: true, _groupShapes: [f.parent, ...f.layoutChildren], _tags: {}, name: 'Group' };
  const h = createUiHarness([group], [f.parent]);
  h.ui.setState([{ id: f.parent.id, isRoundRect: true, layoutRole: 'parent' }]);
  const r = await h.ui.applyLayoutToChildren(f.parent.id, { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'same' }, f.layoutChildren.map((s) => s.id));
  assert.equal(r.ok, false); assert.equal(r.reason, 'strict');
  ['ungroupShapeGroup', 'setBox', 'setAdjFraction', 'addTag'].forEach((method) => h.assertNotCalled(method));
});

test('Brush protection stores the actual clamped fixed radius, including zero', async () => {
  for (const requested of [2, 0]) {
    const shape = makeFixtureShape({ id: 'target', widthCm: 2, heightCm: 1 });
    const h = createHarness({ shapes: [shape] });
    const r = await core.applyPickedToSelection(h.driver, [shape], { cm: requested, sourceStrict: true }, { syncStrict: true });
    assert.equal(r.ok, true);
    const state = await core.readLockState(h.driver, shape);
    assert.equal(state.isStrict, true); assert.equal(state.lockedCm, Math.min(requested, 0.5));
  }
});

test('Zero locks survive read/write and synchronize when radius is changed', async () => {
  for (const key of ['radiusLock_v1', 'RADIUSLOCK_V1']) {
    const shape = makeFixtureShape({ id: 'zero', tags: { [key]: '0' } });
    const h = createHarness({ shapes: [shape] });
    assert.equal((await core.readLockState(h.driver, shape)).lockedCm, 0);
    const r = await core.writeRadius(h.driver, shape, 0.2);
    assert.equal(r.wasLocked, true); assert.equal(shape._tags[key], '0.2');
    assert.equal(Object.keys(shape._tags).length, 1);
  }
});

test('Actual UI recognizes and protects zero radius', async () => {
  const shape = makeFixtureShape({ id: 'zero', adjFraction: 0 });
  const h = createUiHarness([shape]);
  await h.ui.refreshSelection();
  assert.equal(h.ui.getState().selectedShapes[0].currentCm, 0);
  await h.ui.onToggleStrict(true);
  assert.equal(shape._tags.radiusLock_v1, '0'); assert.equal(shape._tags.radiusLockStrict_v1, '1');
  assert.equal(h.ui.getState().selectedShapes[0].locked, true);
});

test('An empty native selection clears layout choices and shows zero selected', async () => {
  const shape = makeFixtureShape({ id: 'old-selection' });
  const h = createUiHarness([shape]);
  await h.ui.refreshSelection();
  assert.ok(h.nodes.get('layout-setup-list').children.length > 0);
  h.select([]); await h.ui.handleSelectionChanged();
  assert.equal(h.nodes.get('layout-setup-list').children.length, 1);
  assert.equal(h.nodes.get('layout-setup-list').children[0].textContent, 'layoutHintEmpty');
  assert.equal(h.nodes.get('lock-hint').textContent, 'emptyShapes');
  assert.equal(h.nodes.get('status-text').textContent, 'statusShapeCountFmt{"count":0}');
});

test('Actual monitor performs zero writes during native group resize', async () => {
  const f = makeStandardFixture();
  const child = f.layoutChildren[0]; child._tags.radiusLock_v1 = '0.15'; child._adjFraction = 0.1; child._groupLevel = 1;
  f.parent._groupLevel = 1;
  const group = { id: 'group', name: 'Group', _isGroup: true, _groupShapes: [f.parent, child], _tags: {} };
  const h = createUiHarness([group]);
  await h.ui.refreshSelection(); await h.ui.monitorTick();
  h.calls.length = 0;
  child.width *= 2; child.height *= 2; f.parent.width *= 2; f.parent.height *= 2;
  await h.ui.monitorTick();
  ['setAdjFraction', 'setBox', 'addTag', 'ungroupShapeGroup'].forEach((method) => h.assertNotCalled(method));
  assert.ok(Array.from(h.timers.values()).some((timer) => timer.delay === 300));
  await h.fireTimer(300);
  const methods = h.calls.map((call) => call.method);
  assert.ok(methods.indexOf('ungroupShapeGroup') < methods.indexOf('setAdjFraction'));
  assert.ok(methods.indexOf('addGroup') > methods.indexOf('setAdjFraction'));
  assert.ok(Math.abs(child._adjFraction * child.height / PT_PER_CM - 0.15) < 1e-9);
});

test('Native yellow-handle changes refresh the displayed fixed radius after settling', async () => {
  const shape = makeFixtureShape({ id: 'native', heightCm: 2, adjFraction: 0.1,
    tags: { radiusLock_v1: '0.2' } });
  const h = createUiHarness([shape]);
  await h.ui.refreshSelection(); await h.ui.monitorTick();
  shape._adjFraction = 0.4;
  for (let i = 0; i < 5; i++) await h.ui.monitorTick();
  assert.equal(shape._tags.radiusLock_v1, '0.8');
  const row = h.nodes.get('shape-list').children[0];
  assert.ok(row.children.some((child) => child.innerHTML.includes('0.80cm')));
});

test('Monitor calls do not overlap and old selection work cannot write', async () => {
  const shape = makeFixtureShape({ id: 'locked', tags: { radiusLock_v1: '0.2' }, adjFraction: 0.1 });
  const h = createUiHarness([shape]);
  await h.ui.refreshSelection(); await h.ui.monitorTick();
  shape.width *= 2; shape.height *= 2;
  const entered = deferred(), release = deferred();
  const read = h.driver.readAdjFraction;
  h.driver.readAdjFraction = async (shape) => { entered.resolve(); await release.promise; return read(shape); };
  h.calls.length = 0;
  const tick = h.ui.monitorTick();
  await entered.promise;
  await h.ui.monitorTick();
  h.ui.invalidate(); h.select([]);
  release.resolve(); await tick;
  assert.equal(h.maxRunning, 1); h.assertNotCalled('setAdjFraction');
  assert.equal(h.ui.getState().lockMonitor.parentRSyncTimer, null);
});

test('Latest selection refresh wins while host reads are delayed', async () => {
  const a = makeFixtureShape({ id: 'a' }), b = makeFixtureShape({ id: 'b' });
  const h = createUiHarness([a, b], [a]);
  const entered = deferred(), release = deferred();
  const read = h.driver.readAdjFraction;
  let calls = 0;
  h.driver.readAdjFraction = async (shape) => {
    if (++calls === 1) { entered.resolve(); await release.promise; }
    return read(shape);
  };
  const old = h.ui.refreshSelection(); await entered.promise;
  h.select([b]); const latest = h.ui.refreshSelection();
  release.resolve(); await Promise.all([old, latest]);
  assert.equal(h.ui.getState().selectedShapes[0].id, 'b'); assert.equal(h.maxRunning, 1);
});

test('UI R-off layouts still update geometry, keeping adjustment unchanged', async () => {
  const f = makeStandardFixture(); const ids = f.layoutChildren.map((s) => s.id);
  f.parent._tags.layoutParent_v1 = JSON.stringify({ rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'off', childIds: ids });
  f.layoutChildren.forEach((shape) => { shape._adjFraction = 0.1; });
  const h = createUiHarness(f.allShapes, [f.parent]);
  await h.ui.refreshSelection(); h.calls.length = 0;
  await h.ui.syncLayoutChildrenRIfNeeded({ geometry: true });
  h.assertCallCount('setBox', 4); h.assertNotCalled('setAdjFraction');
  assert.ok(Math.abs(f.layoutChildren[0].width / PT_PER_CM - 5.6) < 1e-9);
});

test('Native layout monitoring resumes after each geometry synchronization', async () => {
  const f = makeStandardFixture(); const ids = f.layoutChildren.map((s) => s.id);
  f.parent._tags.layoutParent_v1 = JSON.stringify({ rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'off', childIds: ids });
  const h = createUiHarness(f.allShapes, [f.parent]);
  await h.ui.refreshSelection(); h.calls.length = 0;
  for (const dimension of ['width', 'height']) {
    await h.ui.monitorTick();
    f.parent[dimension] += PT_PER_CM;
    await h.ui.monitorTick(); await h.fireTimer(200);
    assert.ok(h.ui.getState().lockMonitor.timer != null, 'The next native edit must still be monitored');
  }
  h.assertCallCount('setBox', 8);
});

test('Multi-parent layout synchronization stops when the selection changes', async () => {
  const roots = [], parents = [];
  for (const id of ['one', 'two']) {
    const parent = makeFixtureShape({ id, widthCm: 12, heightCm: 8 });
    const child = makeFixtureShape({ id: id + '-child' });
    parent._tags.layoutParent_v1 = JSON.stringify({ rows: 1, cols: 1, padding: 0.3, gutter: 0.2, linkRMode: 'off', childIds: [child.id] });
    roots.push(parent, child); parents.push(parent);
  }
  const h = createUiHarness(roots, parents);
  await h.ui.refreshSelection(); h.calls.length = 0;
  const sync = h.driver.sync; let changed = false;
  h.driver.sync = async () => {
    if (!changed && h.calls.some((call) => call.method === 'setBox')) {
      changed = true; h.select([]); h.ui.invalidate();
    }
    await sync();
  };
  await h.ui.syncLayoutChildrenRIfNeeded({ geometry: true });
  h.assertCallCount('setBox', 1);
});

test('Transient mutation failures refresh protection state before resuming monitoring', async () => {
  for (const invoke of [(h) => h.ui.onApply(), (h) => h.ui.onToggleLock(),
    (h) => h.ui.onToggleStrict(true), (h) => h.ui.onReapply()]) {
    const shape = makeFixtureShape({ id: 'locked', adjFraction: 0.1, tags: { radiusLock_v1: '0.2' } });
    const h = createUiHarness([shape]);
    await h.ui.refreshSelection();
    h.nodes.get('radius-input').value = '0.4';
    let reject = true;
    h.failRun(async (callback) => {
      if (reject) {
        reject = false; shape._tags.radiusLockStrict_v1 = '1';
        throw new Error('Transient host failure');
      }
      return callback({});
    });
    await invoke(h);
    assert.equal(h.ui.getState().selectedShapes[0].strictLocked, true);
    assert.ok(h.ui.getState().lockMonitor.timer != null);
  }
});

test('UI wrappers settle on host rejection before callback or during implicit sync', async () => {
  const h = createUiHarness([]);
  h.failRun(async () => { throw new Error('Host rejected'); });
  for (const fn of [() => h.ui.loadLocksViaTags(), () => h.ui.saveLocksViaTags({}, {}),
    () => h.ui.updateLockTagForShape('x', 0, true), () => h.ui.loadLayoutTagsViaTags(),
    () => h.ui.saveLayoutTags('x', {}, []), () => h.ui.deleteLayoutTags('x', [])]) {
    const result = await fn(); assert.equal(result.ok, false);
  }
  h.failRun(async (callback) => { await callback({}); throw new Error('Implicit sync failed'); });
  assert.equal((await h.ui.loadLocksViaTags()).ok, false);
});

test('Queue retains mutual exclusion through final sync and recovers after failure', async () => {
  const release = deferred(); const order = [];
  const queue = createHostQueue(async (fn) => { const r = await fn(); order.push('sync'); await release.promise; return r; });
  const first = queue.run(async () => { order.push('first'); throw new Error('Failed'); });
  await assert.rejects(first, /Failed/);
  const second = queue.run(async () => { order.push('second'); });
  const third = queue.run(async () => { order.push('third'); });
  await turn(); assert.deepEqual(order, ['first', 'second', 'sync']);
  release.resolve(); await Promise.all([second, third]);
  assert.deepEqual(order, ['first', 'second', 'sync', 'third', 'sync']);
});

test('Shape names render as plain text even when they contain HTML', () => {
  const h = createUiHarness([]);
  const name = '<img src=x onerror="globalThis.compromised=true">';
  h.ui.setState([{ id: 'x', name, isRoundRect: true, currentCm: 0, locked: false }]);
  h.ui.renderShapeList();
  const nameNode = h.nodes.get('shape-list').children[0].children[0];
  assert.equal(nameNode.textContent, name); assert.equal(nameNode.innerHTML, '');
});

test('Static server survives malformed URLs and blocks private files/traversal', async () => {
  const server = createStaticServer();
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const get = (path) => new Promise((resolve, reject) => {
    http.get({ host: '127.0.0.1', port: server.address().port, path }, (response) => {
      response.resume(); response.on('end', () => resolve(response.statusCode));
    }).on('error', reject);
  });
  try {
    assert.equal(await get('/%'), 400);
    assert.equal(await get('/%00'), 400);
    assert.equal(await get('/.git/config'), 403);
    assert.equal(await get('/../radius_in_ppt2/manifest.xml'), 403);
    assert.equal(await get('/src/dialog/dialog.html'), 200);
    assert.equal(await get('/manifest.xml'), 200);
  } finally { await new Promise((resolve) => server.close(resolve)); }
});

test('Legacy strict-without-lock tags acquire a persistent baseline after settling', async () => {
  const shape = makeFixtureShape({ id: 'legacy', adjFraction: 0.1, tags: { radiusLockStrict_v1: '1' } });
  const h = createUiHarness([shape]);
  await h.ui.refreshSelection();
  assert.equal(h.ui.getState().selectedShapes[0].locked, true);
  h.calls.length = 0;
  await h.ui.monitorTick();
  h.assertNotCalled('addTag');
  await h.fireTimer(300);
  assert.ok(Math.abs(Number(shape._tags.radiusLock_v1) - 0.2) < 1e-9);
  assert.equal(shape._tags.radiusLockStrict_v1, '1');
});

test('Nested groups keep their hierarchy and metadata through radius writes', async () => {
  const shape = makeFixtureShape({ id: 'leaf', heightCm: 2 });
  const extra = makeFixtureShape({ id: 'extra' });
  const inner = { id: 'inner', _isGroup: true, _groupShapes: [shape, extra], _tags: { INNER: 'yes' }, name: 'Inner' };
  const sibling = makeFixtureShape({ id: 'sibling' });
  const root = { id: 'outer', _isGroup: true, _groupShapes: [inner, sibling], _tags: { OUTER: 'yes' }, name: 'Outer' };
  const h = createHarness({ shapes: [root] });
  const result = await core.applyRadiusToSelection(h.driver, [shape], 0.4);
  assert.equal(result.ok, true);
  assert.equal(h.slide.shapes.items.length, 1);
  const restored = h.slide.shapes.items[0];
  assert.equal(restored.name, 'Outer'); assert.equal(restored._tags.OUTER, 'yes');
  const restoredInner = restored._groupShapes.find((s) => s._isGroup);
  assert.equal(restoredInner.name, 'Inner'); assert.equal(restoredInner._tags.INNER, 'yes');
  assert.ok(restoredInner._groupShapes.includes(shape)); assert.ok(restoredInner._groupShapes.includes(extra));
  assert.ok(Math.abs(shape._adjFraction - 0.2) < 1e-9);
  assert.equal(extra._adjFraction, 0);
});

test('Rapid layout requests serialize and the latest input survives refresh', async () => {
  const f = makeStandardFixture();
  const ids = f.layoutChildren.map((s) => s.id);
  f.parent._tags.layoutParent_v1 = JSON.stringify({ rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'off', childIds: ids });
  const h = createUiHarness(f.allShapes, [f.parent]);
  await h.ui.refreshSelection();
  const entered = deferred(), release = deferred();
  const load = h.driver.loadShapeTree;
  let calls = 0;
  h.driver.loadShapeTree = async (...args) => {
    if (++calls === 1) { entered.resolve(); await release.promise; }
    return load(...args);
  };
  const first = h.ui.applyLayoutFromUI(); await entered.promise;
  h.ui.getState().currentLayout.params.padding = 0.7;
  const second = h.ui.applyLayoutFromUI(); release.resolve();
  await Promise.all([first, second]);
  assert.equal(h.maxRunning, 1);
  assert.equal(JSON.parse(f.parent._tags.layoutParent_v1).padding, 0.7);
  assert.equal(h.ui.getState().currentLayout.params.padding, 0.7);
});

test('Regroup protection window still processes a real user selection change', async () => {
  const a = makeFixtureShape({ id: 'a' }), b = makeFixtureShape({ id: 'b' });
  const h = createUiHarness([a, b], [a]);
  await h.ui.refreshSelection();
  h.ui.ignoreSelection({ slideId: 'test-slide', shapeIds: ['a'] });
  h.select([b]); await h.ui.handleSelectionChanged();
  assert.equal(h.ui.getState().selectedShapes[0].id, 'b');
});

test('An interrupted reapply cannot write after the selection changes', async () => {
  const shape = makeFixtureShape({ id: 'locked', tags: { radiusLock_v1: '0.4' } });
  const h = createUiHarness([shape]);
  await h.ui.refreshSelection();
  const entered = deferred(), release = deferred();
  const load = h.driver.loadTagsBulk;
  h.driver.loadTagsBulk = async (...args) => { entered.resolve(); await release.promise; return load(...args); };
  h.calls.length = 0;
  const reapply = h.ui.onReapply(); await entered.promise;
  h.ui.invalidate(); h.select([]); release.resolve(); await reapply;
  h.assertNotCalled('setAdjFraction');
});

test('Groups and their metadata recover when a radius transaction fails', async () => {
  const a = makeFixtureShape({ id: 'a' }), b = makeFixtureShape({ id: 'b' });
  const group = { id: 'group', name: 'Original', _isGroup: true, _groupShapes: [a, b], _tags: { KEEP: 'yes' } };
  const h = createHarness({ shapes: [group] });
  await assert.rejects(core.withWritableShapes(h.driver, ['a'], async () => { throw new Error('Write failed'); }), /Write failed/);
  assert.equal(h.slide.shapes.items.length, 1);
  const restored = h.slide.shapes.items[0];
  assert.equal(restored.name, 'Original'); assert.equal(restored._tags.KEEP, 'yes');
  assert.deepEqual(restored._groupShapes, [a, b]);
});

test('Metadata restore errors are reported after recovering the group hierarchy', async () => {
  const a = makeFixtureShape({ id: 'a' }), b = makeFixtureShape({ id: 'b' });
  const sibling = makeFixtureShape({ id: 'sibling' });
  const inner = { id: 'inner', name: 'Inner', _isGroup: true, _groupShapes: [a, b], _tags: { INNER: 'yes' } };
  const outer = { id: 'outer', name: 'Outer', _isGroup: true, _groupShapes: [inner, sibling], _tags: { OUTER: 'yes' } };
  const h = createHarness({ shapes: [outer] });
  const setName = h.driver.setShapeName;
  h.driver.setShapeName = (group, name) => {
    if (name === 'Inner') throw new Error('Name restore failed');
    setName(group, name);
  };
  const result = await core.applyRadiusToSelection(h.driver, [a], 0.4);
  assert.equal(result.ok, false); assert.match(result.error, /Name restore failed/);
  assert.equal(h.slide.shapes.items.length, 1);
  const root = h.slide.shapes.items[0];
  assert.equal(root.name, 'Outer'); assert.equal(root._tags.OUTER, 'yes');
  const restoredInner = root._groupShapes.find((s) => s._isGroup);
  assert.deepEqual(restoredInner._groupShapes, [a, b]);
  assert.equal(restoredInner._tags.INNER, 'yes');

  const f = makeStandardFixture();
  const layoutGroup = { id: 'layout', name: 'Layout', _isGroup: true,
    _groupShapes: [f.parent, ...f.layoutChildren], _tags: { KEEP: 'yes' } };
  const layout = createHarness({ shapes: [layoutGroup] });
  layout.driver.setShapeName = () => { throw new Error('Layout name restore failed'); };
  const r = await core.applyLayout(layout.driver, f.parent.id,
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'same' },
    f.layoutChildren.map((s) => s.id));
  assert.equal(r.ok, false); assert.match(r.error, /Layout name restore failed/);
  assert.equal(layout.slide.shapes.items.length, 1);
  assert.equal(layout.slide.shapes.items[0]._isGroup, true);
  assert.equal(layout.slide.shapes.items[0]._tags.KEEP, 'yes');
  assert.equal(layout.slide._selectedShapeIds.length, 1);
});

test('Windows launcher points to the actual flat package root', () => {
  const text = require('fs').readFileSync(require('path').join(__dirname, '../app/Windows/RadiusInPpt.bat'), 'utf8');
  assert.match(text, /set "APP_DIR=%~dp0"/);
  assert.match(text, /set "RES_DIR=%APP_DIR%"/);
  assert.doesNotMatch(text, /\/b ""/);
});

suite.run();
