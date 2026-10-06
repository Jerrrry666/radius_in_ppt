const assert = require('assert/strict');
const fs = require('fs');
const { createTestRunner } = require('./test-harness');
const { createUiHarness } = require('./ui-harness');
const { makeFixtureShape } = require('./fixtures');
const { createCommandController } = require('../src/ribbon/commands');
const { createOfficeUiDriver } = require('../src/lib/office-ui-driver');
const suite = createTestRunner();
const test = suite.test;
function controller(execute) {
  const calls = [];
  let token = 1, receive, closeEvent;
  const pane = { ready: Promise.resolve(), execute: execute || (async (...args) => calls.push(args)),
    selectionToken: () => token, inputValue: () => ({ value: 0.3, unit: 'cm' }), notify: (msg) => calls.push(['notify', msg]) };
  const ui = { showPane: async () => calls.push(['show']), hidePane: async () => calls.push(['hide']),
    showDialog: async (url, message, close) => { calls.push(['dialog', url]); receive = message; closeEvent = close; return { close: () => calls.push(['close']) }; } };
  return { calls, pane, ui, actions: createCommandController(pane, ui, 'http://localhost:3000'),
    changeSelection() { token++; }, message(value) { receive({ message: JSON.stringify(value), origin: 'http://localhost:3000' }); },
    close() { closeEvent({ error: 12006 }); } };
}
const tick = () => new Promise((resolve) => setImmediate(resolve));

test('Ribbon routes finish Office commands and retain zero as a valid preset', async () => {
  const h = controller(); let completed = 0;
  await h.actions.RadiusPresetZero({ completed() { completed++; } });
  assert.deepEqual(h.calls, [['apply', { value: 0, unit: 'cm' }]]); assert.equal(completed, 1);
});
test('Ribbon commands serialize complete actions, recover after failure, and always complete', async () => {
  let release; const gate = new Promise((resolve) => { release = resolve; }); const order = []; let completed = 0;
  const h = controller(async (action) => { order.push(action); if (action === 'lock') { await gate; throw new Error('host failure'); } });
  const a = h.actions.RadiusLock({ completed() { completed++; } });
  const b = h.actions.RadiusReapply({ completed() { completed++; } });
  await tick(); assert.deepEqual(order, ['lock']); release(); await Promise.all([a, b]);
  assert.deepEqual(order, ['lock', 'reapply']); assert.equal(completed, 2); assert(h.calls.some((x) => x[0] === 'show'));
});
test('Dialog submits through the common apply path and closes once', async () => {
  const h = controller(); const work = h.actions.RadiusSet({ completed() {} }); await tick();
  h.message({ value: 12, unit: '%' }); await work;
  assert.deepEqual(h.calls.find((x) => x[0] === 'apply'), ['apply', { value: 12, unit: '%' }]);
  assert.equal(h.calls.filter((x) => x[0] === 'close').length, 1);
});
test('Changing selection while entering a number rejects the old dialog without writing', async () => {
  const h = controller(); const work = h.actions.RadiusSet({ completed() {} }); await tick();
  h.changeSelection(); h.message({ value: 0.8, unit: 'cm' }); await work;
  assert(!h.calls.some((x) => x[0] === 'apply')); assert(h.calls.some((x) => x[0] === 'notify'));
});
test('Dialog cancellation and invalid payloads never write radius', async () => {
  for (const payload of [{ cancel: true }, { value: -1, unit: 'cm' }, { value: 2, unit: 'unknown' }]) {
    const h = controller(); const work = h.actions.RadiusSet({ completed() {} }); await tick(); h.message(payload); await work;
    assert(!h.calls.some((x) => x[0] === 'apply'));
  }
});
test('Real UI adapter applies to groups and keeps all strict targets protected', async () => {
  const a = makeFixtureShape({ id: 'a' }), b = makeFixtureShape({ id: 'b' });
  const group = { id: 'group', _isGroup: true, _groupShapes: [a, b], _tags: {}, name: 'Group' };
  const h = createUiHarness([group]);
  await h.paneActions.execute('apply', { value: 0.3, unit: 'cm' });
  assert(a._adjFraction > 0); assert(b._adjFraction > 0);
  a._tags.radiusLockStrict_v1 = '1'; a._tags.radiusLock_v1 = '0.3'; h.calls.length = 0;
  await assert.rejects(h.paneActions.execute('apply', { value: 0.8, unit: 'cm' }), /protection/);
  h.assertNotCalled('setAdjFraction'); h.assertNotCalled('deleteTag'); h.assertNotCalled('ungroupShapeGroup');
});
test('Top lock uses each current R, including zero, instead of a stale hidden input', async () => {
  const shape = makeFixtureShape({ id: 'zero', adjFraction: 0 }); const h = createUiHarness([shape]);
  h.nodes.set('radius-input', { value: '9' });
  await h.paneActions.execute('lock'); assert.equal(shape._tags.radiusLock_v1, '0');
});
test('Office UI driver registers actions and surfaces transport errors', async () => {
  const associated = []; const office = { actions: { associate: (...args) => associated.push(args) },
    AsyncResultStatus: { Succeeded: 'succeeded' }, context: { ui: { displayDialogAsync(url, opts, cb) { cb({ status: 'failed', error: { message: 'dialog unavailable' } }); } } } };
  const driver = createOfficeUiDriver(office); const action = () => {};
  driver.associate({ TestRadiusCommand: action }); assert.deepEqual(associated, [['TestRadiusCommand', action]]);
  await assert.rejects(driver.showDialog('url', () => {}, () => {}), /dialog unavailable/);
});
test('Manifest commands use one long shared runtime, a valid extension point, and resolved resources', () => {
  const manifest = fs.readFileSync(require('path').join(__dirname, '../manifest.xml'), 'utf8');
  assert(manifest.includes('xsi:type="PrimaryCommandSurface"')); assert(manifest.includes('lifetime="long"'));
  assert(!manifest.includes('<TaskpaneId>')); assert(!manifest.includes('<TaskpaneID>'));
  assert.equal((manifest.match(/<Runtime /g) || []).length, 1);
  const resources = new Set(Array.from(manifest.matchAll(/<bt:(?:String|Url|Image) id="([^"]+)"/g), (m) => m[1]));
  for (const match of manifest.matchAll(/resid="([^"]+)"/g)) assert(resources.has(match[1]), match[1]);
  const h = controller();
  for (const match of manifest.matchAll(/<FunctionName>([^<]+)<\/FunctionName>/g)) assert.equal(typeof h.actions[match[1]], 'function');
});
suite.run();
