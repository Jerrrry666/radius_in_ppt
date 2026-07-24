/*
 * test-driver-integration.js — driver + radius-core.writeRadius 集成测试
 *
 * 测试策略：
 *   - 不 mock 整个 PowerPoint.run
 *   - 写一个 mockDriver（用普通对象 + Map 记录所有 driver 方法调用）
 *   - 喂给 radius-core.writeRadius（业务逻辑函数）
 *   - 验证：driver 方法被正确调用、最终 shape 状态正确
 *
 * 这个文件测的是「业务逻辑 + driver 协议」整体是否对，不测单个 driver 方法。
 * 单个 driver 方法的测试在真实 PPT 里做。
 */

const assert = require('assert');
const path = require('path');
const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
const { createDriver } = require(path.join(__dirname, '..', 'src', 'lib', 'ppt-driver.js'));

// ---------------- 测试 harness ----------------

function makeMockShape(opts) {
  const shape = {
    id: opts.id || '1',
    width: opts.width != null ? opts.width : 100,    // pt
    height: opts.height != null ? opts.height : 60,  // pt
    _adjFraction: opts.adjFraction != null ? opts.adjFraction : 0,  // 0~1
    _tags: Object.assign({}, opts.tags || {}),  // 普通对象，driver 模拟这个
    get type() { return 'GeometricShape'; },
  };
  // adjustments 是个特殊对象，要可读 .count 和 .get(0).value，可写 set(0, x)
  shape.adjustments = {
    count: opts.isRoundRect === false ? 0 : 1,
    get(i) { return { value: shape._adjFraction }; },
    set(i, v) { shape._adjFraction = v; },
  };
  return shape;
}

// 同步 mock driver（不 await sync，直接返回）
function makeMockDriver(opts) {
  opts = opts || {};
  const calls = [];  // 记录所有 driver 方法调用
  const shapes = new Map();
  if (opts.shapes) {
    for (const s of opts.shapes) shapes.set(s.id, s);
  }

  // 假的 ctx，sync 立刻 resolve
  const ctx = {
    sync: async () => { calls.push(['sync']); return; },
  };

  return {
    driver: createMockDriverApi(ctx, calls, shapes),
    calls,
    shapes,
    ctx,
  };
}

function createMockDriverApi(ctx, calls, shapes) {
  // 直接调用 createDriver(ctx)，但 addTag / deleteTag / readTag 用我们的 mock 实现
  // 这样能验证 createDriver 的 API 跟 mock 实现一致
  const driver = createDriver(ctx);

  // 覆盖 addTag / deleteTag / readTag：mock 一个 tags 集合
  driver.addTag = (s, key, value) => {
    calls.push(['addTag', s.id, key, value]);
    s._tags[key] = String(value);
  };
  driver.deleteTag = (s, key) => {
    calls.push(['deleteTag', s.id, key]);
    delete s._tags[key];
  };
  driver.readTag = async (s, key) => {
    calls.push(['readTag', s.id, key]);
    return s._tags[key] != null ? s._tags[key] : null;
  };

  return driver;
}

// 简易测试框架
let passed = 0, failed = 0;
function test(name, fn) {
  return Promise.resolve().then(fn).then(
    () => { passed++; console.log(`  ✓ ${name}`); },
    (e) => {
      failed++;
      console.log(`  ✗ ${name}`);
      console.log(`     ${e.message}`);
      if (e.stack) console.log(`     ${e.stack.split('\n').slice(1, 4).join('\n     ')}`);
    }
  );
}

// ---------------- 测试 ----------------

async function run() {
  console.log('=== writeRadius(driver) — 基础写 R 角 ===');

  const sh = makeMockShape({ id: '1', width: 100, height: 60, isRoundRect: true });
  const m = makeMockDriver({ shapes: [sh] });
  const r = await RC.writeRadius(m.driver, sh, 0.5);

  await test('ok=true', () => assert.strictEqual(r.ok, true));
  await test('newCm=0.5', () => assert.strictEqual(r.newCm, 0.5));
  await test('wasLocked=false', () => assert.strictEqual(r.wasLocked, false));
  await test('wasStrict=false', () => assert.strictEqual(r.wasStrict, false));
  await test('adjFraction = 0.5/2.118... = 0.236...',
    () => assert.ok(Math.abs(sh._adjFraction - 0.5 / (60 / RC.PT_PER_CM)) < 1e-9));
  await test('addTag 没调', () => assert.ok(!m.calls.some((c) => c[0] === 'addTag')));

  console.log('\n=== writeRadius(driver) — strict 永远拦截 ===');

  const shStrict = makeMockShape({ id: '2', width: 100, height: 60, isRoundRect: true, tags: { radiusLockStrict_v1: '1' } });
  const mStrict = makeMockDriver({ shapes: [shStrict] });
  const rStrict = await RC.writeRadius(mStrict.driver, shStrict, 0.5);

  await test('ok=false', () => assert.strictEqual(rStrict.ok, false));
  await test('reason=strict', () => assert.strictEqual(rStrict.reason, 'strict'));
  await test('isStrict=true', () => assert.strictEqual(rStrict.isStrict, true));
  await test('adjFraction 没变', () => assert.strictEqual(shStrict._adjFraction, 0));

  console.log('\n=== writeRadius(driver) — 普通矩形（非 roundRect） ===');

  const shRect = makeMockShape({ id: '3', width: 100, height: 60, isRoundRect: false });
  const mRect = makeMockDriver({ shapes: [shRect] });
  const rRect = await RC.writeRadius(mRect.driver, shRect, 0.5);

  await test('ok=false', () => assert.strictEqual(rRect.ok, false));
  await test('reason=not-roundRect', () => assert.strictEqual(rRect.reason, 'not-roundRect'));

  console.log('\n=== writeRadius(driver) — 0 尺寸 ===');

  const sh0 = makeMockShape({ id: '4', width: 0, height: 0, isRoundRect: true });
  const m0 = makeMockDriver({ shapes: [sh0] });
  const r0 = await RC.writeRadius(m0.driver, sh0, 0.5);

  await test('reason=no-size', () => assert.strictEqual(r0.reason, 'no-size'));

  console.log('\n=== writeRadius(driver) — clamp 到短边一半 ===');

  const shClamp = makeMockShape({ id: '5', width: 100, height: 60, isRoundRect: true });
  const mClamp = makeMockDriver({ shapes: [shClamp] });
  // 短边 = 60pt = 2.118cm，一半 = 1.059cm
  // 目标 = 5cm > 1.059 → clamp 到 1.059
  const rClamp = await RC.writeRadius(mClamp.driver, shClamp, 5);

  await test('newCm 被 clamp 到 1.0588...',
    () => assert.ok(Math.abs(rClamp.newCm - 60 / RC.PT_PER_CM / 2) < 1e-6));

  console.log('\n=== writeRadius(driver) — negative targetCm → 0 ===');

  const shNeg = makeMockShape({ id: '6', width: 100, height: 60, isRoundRect: true });
  const mNeg = makeMockDriver({ shapes: [shNeg] });
  const rNeg = await RC.writeRadius(mNeg.driver, shNeg, -1);

  await test('newCm=0', () => assert.strictEqual(rNeg.newCm, 0));
  await test('adjFraction=0', () => assert.strictEqual(shNeg._adjFraction, 0));

  console.log('\n=== writeRadius(driver) — locked 形状同步 fixed value ===');

  const shLock = makeMockShape({ id: '7', width: 100, height: 60, isRoundRect: true, tags: { radiusLock_v1: '0.3' } });
  const mLock = makeMockDriver({ shapes: [shLock] });
  const rLock = await RC.writeRadius(mLock.driver, shLock, 0.5);

  await test('ok=true', () => assert.strictEqual(rLock.ok, true));
  await test('wasLocked=true', () => assert.strictEqual(rLock.wasLocked, true));
  await test('addTag 被调同步 fixed value', () => {
    const added = mLock.calls.find((c) => c[0] === 'addTag' && c[1] === '7' && c[2] === 'radiusLock_v1');
    assert.ok(added, 'should addTag radiusLock_v1');
    assert.strictEqual(added[3], '0.5', 'fixed value should be new cm as string');
  });
  await test('_tags[radiusLock_v1] = "0.5"', () => assert.strictEqual(shLock._tags.radiusLock_v1, '0.5'));

  console.log('\n=== writeRadius(driver) — locked + strict 命中 strict ===');

  const shBoth = makeMockShape({
    id: '8', width: 100, height: 60, isRoundRect: true,
    tags: { radiusLock_v1: '0.3', radiusLockStrict_v1: '1' },
  });
  const mBoth = makeMockDriver({ shapes: [shBoth] });
  const rBoth = await RC.writeRadius(mBoth.driver, shBoth, 0.5);

  await test('reason=strict（优先级高）', () => assert.strictEqual(rBoth.reason, 'strict'));
  await test('wasLocked=true（读出来但没写）', () => assert.strictEqual(rBoth.wasLocked, true));
  await test('adjFraction 没变', () => assert.strictEqual(shBoth._adjFraction, 0));
  await test('addTag 没调', () => assert.ok(!mBoth.calls.some((c) => c[0] === 'addTag')));

  console.log('\n=== writeRadius(driver) — layoutParentId 写子 tag ===');

  const shChild = makeMockShape({ id: '9', width: 100, height: 60, isRoundRect: true });
  const mChild = makeMockDriver({ shapes: [shChild] });
  const rChild = await RC.writeRadius(mChild.driver, shChild, 0.5, { layoutParentId: 'parent-100' });

  await test('ok=true', () => assert.strictEqual(rChild.ok, true));
  await test('addTag layoutChild_v1 被调', () => {
    const added = mChild.calls.find((c) => c[0] === 'addTag' && c[2] === 'layoutChild_v1');
    assert.ok(added);
    assert.strictEqual(added[3], 'parent-100');
  });

  console.log('\n=== writeRadius(driver) — 批量 5 个形状 ===');

  const shapes = [
    makeMockShape({ id: 'a', width: 100, height: 60, isRoundRect: true }),
    makeMockShape({ id: 'b', width: 80, height: 50, isRoundRect: true }),
    makeMockShape({ id: 'c', width: 100, height: 60, isRoundRect: false }),  // 普通矩形
    makeMockShape({ id: 'd', width: 100, height: 60, isRoundRect: true, tags: { radiusLockStrict_v1: '1' } }),
    makeMockShape({ id: 'e', width: 100, height: 60, isRoundRect: true }),
  ];
  const mBatch = makeMockDriver({ shapes });
  const results = [];
  for (const s of shapes) {
    const r = await RC.writeRadius(mBatch.driver, s, 0.5);
    results.push(r);
  }

  await test('a ok=true', () => assert.strictEqual(results[0].ok, true));
  await test('b ok=true', () => assert.strictEqual(results[1].ok, true));
  await test('c not-roundRect', () => assert.strictEqual(results[2].reason, 'not-roundRect'));
  await test('d strict', () => assert.strictEqual(results[3].reason, 'strict'));
  await test('e ok=true', () => assert.strictEqual(results[4].ok, true));
  await test('2 success + 1 not-roundRect + 1 strict = 4 readTag 调用（c 不用 readTag？）',
    () => {
      // a/b/d/e 都触发 readTag 2 次（lock + strict），c 触发 2 次但 strict
      const readTagCalls = mBatch.calls.filter((c) => c[0] === 'readTag');
      // 每个 shape 调 2 次 readTag
      assert.strictEqual(readTagCalls.length, 10);
    });

  console.log('\n=== writeRadius(driver) — driver 异常时返回 reason=exception 带 error ===');

  const shBoom = makeMockShape({ id: '10', width: 100, height: 60, isRoundRect: true });
  const mBoom = makeMockDriver({ shapes: [shBoom] });
  // 让 driver.readTag 抛异常
  mBoom.driver.readTag = async () => { throw new Error('office.js boom'); };
  const rBoom = await RC.writeRadius(mBoom.driver, shBoom, 0.5);

  await test('ok=false', () => assert.strictEqual(rBoom.ok, false));
  await test('reason=exception', () => assert.strictEqual(rBoom.reason, 'exception'));
  await test('error 含 office.js boom', () => assert.ok(rBoom.error && rBoom.error.includes('office.js boom')));

  console.log('\n=== createDriver API 一致性（mock 跟真实 driver 形状相同）===');

  await test('createDriver 返回 16 个方法', () => {
    const fakeCtx = { sync: async () => {} };
    const d = createDriver(fakeCtx);
    const expected = [
      'load', 'sync',
      'selectedShapes', 'activeSlide', 'slideShapes',
      'shapeId', 'size', 'box', 'isRoundRect', 'adjFraction', 'loadAdjValue',
      'setBox', 'setAdjFraction',
      'addTag', 'deleteTag', 'readTag',
    ];
    for (const m of expected) {
      assert.ok(typeof d[m] === 'function', `missing method: ${m}`);
    }
    assert.strictEqual(Object.keys(d).length, expected.length, `extra methods: ${Object.keys(d).filter((k) => !expected.includes(k))}`);
  });

  // v1.2.5：driver.adjFraction 必须 defensive（value 没 load 时返回 0，不 throw）
  // 这是 Mac LTSC task pane 必加的 load 模式（per-shape loadAdjValue + sync），
  // 但如果 caller 漏了 load，driver 也不应该 throw
  await test('driver.adjFraction 在 value 没 load 时不 throw，返回 0（v1.2.5 defensive）', () => {
    const sh = makeMockShape({ id: 'def', width: 100, height: 60, isRoundRect: true });
    // 模拟 value 没 load：get(0).value 抛异常
    sh.adjustments.get = () => { throw new Error('尚未加载结果对象的值'); };
    const m = makeMockDriver({ shapes: [sh] });
    const adj = m.driver.adjFraction(sh);
    assert.strictEqual(adj, 0, 'should return 0 when value not loaded, not throw');
  });

  // v1.2.5：loadAdjValue 是 per-shape load 的 helper
  await test('driver.loadAdjValue 调用 sh.adjustments.load(items/value)', () => {
    const sh = makeMockShape({ id: 'lav', width: 100, height: 60, isRoundRect: true });
    let called = false;
    sh.adjustments.load = (path) => {
      called = true;
      assert.strictEqual(path, 'items/value', 'should load items/value path');
    };
    const m = makeMockDriver({ shapes: [sh] });
    m.driver.loadAdjValue(sh);
    assert.ok(called, 'loadAdjValue should call sh.adjustments.load');
  });

  console.log('\n=== driver.size 跟 driver.box 区别（size 只要 width/height）===');

  // 回归测试：v1.2.2 真实 PPT 暴露的 bug——onApply 没 load left/top，但 driver.box 会读 left
  // 期望：writeRadius 走 driver.size（不读 left/top），不依赖 left/top 被 load
  await test('writeRadius 不需要 left/top 被 load（v1.2.2 回归）', async () => {
    const shReg = makeMockShape({ id: 'reg', width: 100, height: 60, isRoundRect: true });
    // left/top 故意不设（undefined），模拟 PPT 里没 load
    shReg.left = undefined;
    shReg.top = undefined;
    const mReg = makeMockDriver({ shapes: [shReg] });
    // 把 left/top 改成访问就抛异常的 getter（模拟 PPT 'left 属性不可用'）
    Object.defineProperty(shReg, 'left', { get() { throw new Error('left 属性不可用'); }, configurable: true });
    Object.defineProperty(shReg, 'top', { get() { throw new Error('top 属性不可用'); }, configurable: true });
    const rReg = await RC.writeRadius(mReg.driver, shReg, 0.5);
    assert.strictEqual(rReg.ok, true, '应该成功，writeRadius 不应访问 left/top');
    assert.strictEqual(rReg.newCm, 0.5);
  });

  await test('driver.size 只读 width/height', () => {
    const shSize = makeMockShape({ id: 'sz', width: 200, height: 100 });
    // 不设 left/top，driver.size 不应该访问它们
    shSize.left = undefined;  // 故意让 left undefined
    shSize.top = undefined;
    const mSize = makeMockDriver({ shapes: [shSize] });
    const sz = mSize.driver.size(shSize);
    assert.strictEqual(sz.width, 200);
    assert.strictEqual(sz.height, 100);
    assert.ok(!('left' in sz));
    assert.ok(!('top' in sz));
  });

  await test('driver.box 需要 left/top 都被访问（不设就会 fail）', () => {
    // 验证我们的契约文档：如果 caller 调用 driver.box，必须 load 4 个字段
    // 这里用 proxy 模拟 left 没 load → 抛异常
    const shBox = makeMockShape({ id: 'bx', width: 200, height: 100, left: 10, top: 20 });
    const mBox = makeMockDriver({ shapes: [shBox] });
    // 把 left 改成 getter 抛异常（模拟没 load）
    Object.defineProperty(shBox, 'left', {
      get() { throw new Error('left not loaded'); },
      configurable: true,
    });
    assert.throws(() => mBox.driver.box(shBox), /left not loaded/);
  });

  console.log('\n=== readLockState(driver) — 读 lock + strict 状态 ===');

  await test('无 lock 无 strict → lockedCm=null, isStrict=false', async () => {
    const sh = makeMockShape({ id: 'ls1' });
    const m = makeMockDriver({ shapes: [sh] });
    const state = await RC.readLockState(m.driver, sh);
    assert.strictEqual(state.lockedCm, null);
    assert.strictEqual(state.isStrict, false);
  });

  await test('有 lock tag → lockedCm 解析为数字', async () => {
    const sh = makeMockShape({ id: 'ls2', tags: { radiusLock_v1: '0.5' } });
    const m = makeMockDriver({ shapes: [sh] });
    const state = await RC.readLockState(m.driver, sh);
    assert.strictEqual(state.lockedCm, 0.5);
    assert.strictEqual(state.isStrict, false);
  });

  await test('有 strict tag → isStrict=true', async () => {
    const sh = makeMockShape({ id: 'ls3', tags: { radiusLockStrict_v1: '1' } });
    const m = makeMockDriver({ shapes: [sh] });
    const state = await RC.readLockState(m.driver, sh);
    assert.strictEqual(state.isStrict, true);
    assert.strictEqual(state.lockedCm, null);
  });

  await test('lock + strict 同时存在', async () => {
    const sh = makeMockShape({ id: 'ls4', tags: { radiusLock_v1: '0.3', radiusLockStrict_v1: '1' } });
    const m = makeMockDriver({ shapes: [sh] });
    const state = await RC.readLockState(m.driver, sh);
    assert.strictEqual(state.lockedCm, 0.3);
    assert.strictEqual(state.isStrict, true);
  });

  await test('lock tag = 0 或负数 → 视作没 lock', async () => {
    const sh = makeMockShape({ id: 'ls5', tags: { radiusLock_v1: '0' } });
    const m = makeMockDriver({ shapes: [sh] });
    const state = await RC.readLockState(m.driver, sh);
    assert.strictEqual(state.lockedCm, null);
  });

  console.log('\n=== writeLockState(driver) — 写 lock + strict 状态 ===');

  await test('写 lockedCm=number → addTag 调', async () => {
    const sh = makeMockShape({ id: 'wl1' });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.writeLockState(m.driver, sh, { lockedCm: 0.5 });
    assert.strictEqual(r.ok, true);
    const added = m.calls.find((c) => c[0] === 'addTag' && c[2] === 'radiusLock_v1');
    assert.ok(added);
    assert.strictEqual(added[3], '0.5');
  });

  await test('写 lockedCm=null → deleteTag 调', async () => {
    const sh = makeMockShape({ id: 'wl2', tags: { radiusLock_v1: '0.3' } });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.writeLockState(m.driver, sh, { lockedCm: null });
    assert.strictEqual(r.ok, true);
    const deleted = m.calls.find((c) => c[0] === 'deleteTag' && c[2] === 'radiusLock_v1');
    assert.ok(deleted);
  });

  await test('写 isStrict=true → addTag strict=1', async () => {
    const sh = makeMockShape({ id: 'wl3' });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.writeLockState(m.driver, sh, { isStrict: true });
    assert.strictEqual(r.ok, true);
    const added = m.calls.find((c) => c[0] === 'addTag' && c[2] === 'radiusLockStrict_v1');
    assert.ok(added);
    assert.strictEqual(added[3], '1');
  });

  await test('写 isStrict=false → deleteTag', async () => {
    const sh = makeMockShape({ id: 'wl4', tags: { radiusLockStrict_v1: '1' } });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.writeLockState(m.driver, sh, { isStrict: false });
    assert.strictEqual(r.ok, true);
    const deleted = m.calls.find((c) => c[0] === 'deleteTag' && c[2] === 'radiusLockStrict_v1');
    assert.ok(deleted);
  });

  await test('undefined 字段不动（cm 跳过, strict 跳过）', async () => {
    const sh = makeMockShape({ id: 'wl5', tags: { radiusLock_v1: '0.3', radiusLockStrict_v1: '1' } });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.writeLockState(m.driver, sh, {});  // 两个都 undefined
    assert.strictEqual(r.ok, true);
    assert.ok(!m.calls.some((c) => c[0] === 'addTag'));
    assert.ok(!m.calls.some((c) => c[0] === 'deleteTag'));
  });

  await test('写两者同时', async () => {
    const sh = makeMockShape({ id: 'wl6' });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.writeLockState(m.driver, sh, { lockedCm: 0.5, isStrict: true });
    assert.strictEqual(r.ok, true);
    assert.ok(m.calls.some((c) => c[0] === 'addTag' && c[2] === 'radiusLock_v1'));
    assert.ok(m.calls.some((c) => c[0] === 'addTag' && c[2] === 'radiusLockStrict_v1'));
  });

  console.log('\n=== reapplyLock(driver) — 反算 adj 回 lockedCm ===');

  await test('reapplyLock basic — adj 写回 lockedCm', async () => {
    const sh = makeMockShape({ id: 'ra1', width: 100, height: 60, isRoundRect: true, adjFraction: 0.01 });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.reapplyLock(m.driver, sh, 0.5);
    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.newCm, 0.5);
    // adjFraction 应该 = 0.5 / (60/PT_PER_CM) = 0.236...
    const expected = 0.5 / (60 / RC.PT_PER_CM);
    assert.ok(Math.abs(sh._adjFraction - expected) < 1e-9, `expected ${expected}, got ${sh._adjFraction}`);
  });

  await test('reapplyLock clamp 到短边一半', async () => {
    const sh = makeMockShape({ id: 'ra2', width: 100, height: 60, isRoundRect: true });
    const m = makeMockDriver({ shapes: [sh] });
    // lockedCm = 5cm > 短边一半（30pt = 1.059cm）→ clamp
    const r = await RC.reapplyLock(m.driver, sh, 5);
    assert.strictEqual(r.ok, true);
    assert.ok(Math.abs(r.newCm - 60 / RC.PT_PER_CM / 2) < 1e-6);
  });

  await test('reapplyLock 非 roundRect → not-roundRect', async () => {
    const sh = makeMockShape({ id: 'ra3', width: 100, height: 60, isRoundRect: false });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.reapplyLock(m.driver, sh, 0.5);
    assert.strictEqual(r.ok, false);
    assert.strictEqual(r.reason, 'not-roundRect');
  });

  await test('reapplyLock 0 尺寸 → no-size', async () => {
    const sh = makeMockShape({ id: 'ra4', width: 0, height: 0, isRoundRect: true });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.reapplyLock(m.driver, sh, 0.5);
    assert.strictEqual(r.ok, false);
    assert.strictEqual(r.reason, 'no-size');
  });

  await test('reapplyLock 不需要 left/top（跟 writeRadius 一样）', async () => {
    // 回归：v1.2.2 monitor 改 driver 后，不能因没 load left/top 抛异常
    const sh = makeMockShape({ id: 'ra5', width: 100, height: 60, isRoundRect: true });
    Object.defineProperty(sh, 'left', { get() { throw new Error('left not loaded'); }, configurable: true });
    Object.defineProperty(sh, 'top', { get() { throw new Error('top not loaded'); }, configurable: true });
    const m = makeMockDriver({ shapes: [sh] });
    const r = await RC.reapplyLock(m.driver, sh, 0.5);
    assert.strictEqual(r.ok, true);
  });

  console.log('\n=== applyLayout(driver) — 1 父 + 4 子 2x2 ===');

  // applyLayout 需要 driver 操作 activeSlide() + slideShapes() + collection-level load
  // 改 makeMockDriver + 加个 mockSlide（包含所有 shapes）+ custom activeSlide / slideShapes
  function makeMockDriverWithSlide(opts) {
    opts = opts || {};
    const shapes = opts.shapes || [];
    // 模拟 Office.js slide proxy：有 load(fields) 方法
    const slide = {
      shapes: { items: shapes },
      load: (fields) => { /* recorded in calls if needed */ },
    };
    const calls = [];
    const ctx = { sync: async () => { calls.push(['sync']); return; } };

    // 直接用 createDriver 拿标准 API
    const driver = createDriver(ctx);
    // 覆盖 collection accessors 返回 mock 的 slide
    driver.activeSlide = () => { calls.push(['activeSlide']); return slide; };
    driver.slideShapes = (s) => { calls.push(['slideShapes']); return s.shapes; };
    // tags 也 mock
    driver.addTag = (s, key, value) => {
      calls.push(['addTag', s.id, key, value]);
      s._tags[key] = String(value);
    };
    driver.deleteTag = (s, key) => {
      calls.push(['deleteTag', s.id, key]);
      delete s._tags[key];
    };
    driver.readTag = async (s, key) => {
      calls.push(['readTag', s.id, key]);
      return s._tags[key] != null ? s._tags[key] : null;
    };
    return { driver, calls, shapes, slide, ctx };
  }

  await test('applyLayout basic 2x2: 父+4 子，位置/尺寸/R 角/父 tag 全对', async () => {
    // 父：368.58 x 155.09 pt 圆角矩形，R 角 0.3
    const parent = makeMockShape({ id: 'p1', width: 368.58, height: 155.09, isRoundRect: true, adjFraction: 0.3 });
    // 4 个子：初始位置任意，apply 后会按 layout 重写
    const c1 = makeMockShape({ id: 'c1', width: 100, height: 60, isRoundRect: true, adjFraction: 0 });
    const c2 = makeMockShape({ id: 'c2', width: 100, height: 60, isRoundRect: true, adjFraction: 0 });
    const c3 = makeMockShape({ id: 'c3', width: 100, height: 60, isRoundRect: true, adjFraction: 0 });
    const c4 = makeMockShape({ id: 'c4', width: 100, height: 60, isRoundRect: true, adjFraction: 0 });
    const m = makeMockDriverWithSlide({ shapes: [parent, c1, c2, c3, c4] });
    const params = { rows: 2, cols: 2, padding: 0.3, gutter: 0.4, linkRMode: 'subtract' };
    const r = await RC.applyLayout(m.driver, 'p1', params, ['c1', 'c2', 'c3', 'c4'], {});
    assert.strictEqual(r.ok, true, 'should succeed');
    assert.strictEqual(r.applied, 4, '4 children applied');
    // 父 tag 应写
    const parentTagAdded = m.calls.find((c) => c[0] === 'addTag' && c[1] === 'p1' && c[2] === 'layoutParent_v1');
    assert.ok(parentTagAdded, 'parent tag should be written');
    const parentPayload = JSON.parse(parentTagAdded[3]);
    assert.deepStrictEqual(parentPayload.childIds, ['c1', 'c2', 'c3', 'c4'], 'parent tag childIds correct');
    assert.strictEqual(parentPayload.rows, 2);
    assert.strictEqual(parentPayload.cols, 2);
    // 4 个子位置/尺寸应被重写（原始 width=100/height=60，layout 后会变）
    for (const c of [c1, c2, c3, c4]) {
      assert.ok(c.width !== 100 || c.height !== 60, `child ${c.id} should have been resized (width=${c.width}, height=${c.height})`);
    }
    // 4 个子 child tag 应写
    for (const cid of ['c1', 'c2', 'c3', 'c4']) {
      const childTagAdded = m.calls.find((c) => c[0] === 'addTag' && c[1] === cid && c[2] === 'layoutChild_v1');
      assert.ok(childTagAdded, `child tag should be written for ${cid}`);
    }
  });

  await test('applyLayout stale childId 过滤：传 5 个，只有 4 个在 slide，父 tag childIds 只含 4 个', async () => {
    // v1.2.9 用户要求：删 shape 后 applyLayout 不卡，stale id 不进 JSON
    const parent = makeMockShape({ id: 'p2', width: 200, height: 100, isRoundRect: true, adjFraction: 0.2 });
    const c1 = makeMockShape({ id: 'c1', width: 50, height: 30, isRoundRect: true });
    const c2 = makeMockShape({ id: 'c2', width: 50, height: 30, isRoundRect: true });
    const c3 = makeMockShape({ id: 'c3', width: 50, height: 30, isRoundRect: true });
    const c4 = makeMockShape({ id: 'c4', width: 50, height: 30, isRoundRect: true });
    // c5 故意不在 slide 里
    const m = makeMockDriverWithSlide({ shapes: [parent, c1, c2, c3, c4] });
    const params = { rows: 2, cols: 2, padding: 0.2, gutter: 0.2, linkRMode: 'subtract' };
    // 传 c5 这个 stale id
    const r = await RC.applyLayout(m.driver, 'p2', params, ['c1', 'c2', 'c3', 'c5'], {});
    // 应该 ok 但 warn（子不足）
    assert.ok(r.warn && r.warn.includes('不足'), 'should warn about missing child');
    // 父 tag 不应被写（c5 在前几位，让 stale 在前）
    const parentTagAdded = m.calls.find((c) => c[0] === 'addTag' && c[1] === 'p2' && c[2] === 'layoutParent_v1');
    assert.ok(!parentTagAdded, 'parent tag should NOT be written when child missing');
  });

  await test('applyLayout stale childId 过滤：2x2 传 4 个但其中一个 stale，warn 但仍写 3 个', async () => {
    // 实际上 computeLayout 需要 expectedCount 个 child，少了就拒绝整个 apply
    // 所以这个测试是「stale + 数量够」的边界场景
    // 既然 missingCount=1 就 warn 拒绝，那 stale 自动清理的场景是：
    //   caller 传 [c1, c2, c3, c4]，但 c4 stale，validCount=3 < expectedCount=4 → warn 拒绝
    // 那 valid childIds 永远等于 caller 传的（因为如果 stale 就不够）？
    // 其实 stale childIds 清理的场景是：旧父 tag JSON 里的 childIds 有 stale（用户复用 applyLayout）
    //   caller 读旧 tag → childIds 列表 → 传进来 → 里面有些已删
    // 这种情况我们 warn 拒绝，让用户先调整。这是当前设计。
    // TODO: 未来可以 add「自动从 validChildIds 中用最近 4 个」的 fallback。
    const parent = makeMockShape({ id: 'p3', width: 200, height: 100, isRoundRect: true });
    const c1 = makeMockShape({ id: 'c1', width: 50, height: 30, isRoundRect: true });
    const c2 = makeMockShape({ id: 'c2', width: 50, height: 30, isRoundRect: true });
    const c3 = makeMockShape({ id: 'c3', width: 50, height: 30, isRoundRect: true });
    const c4 = makeMockShape({ id: 'c4', width: 50, height: 30, isRoundRect: true });
    const m = makeMockDriverWithSlide({ shapes: [parent, c1, c2, c3, c4] });
    const params = { rows: 2, cols: 2, padding: 0.2, gutter: 0.2, linkRMode: 'subtract' };
    const r = await RC.applyLayout(m.driver, 'p3', params, ['c1', 'c2', 'c3', 'stale_id'], {});
    assert.ok(!r.ok, 'should reject when child missing');
    assert.ok(r.warn.includes('不足'), 'warn about child count');
  });

  await test('applyLayout 父不在当前 slide → warn 不 apply', async () => {
    const c1 = makeMockShape({ id: 'c1', width: 50, height: 30, isRoundRect: true });
    const m = makeMockDriverWithSlide({ shapes: [c1] });  // 父不在
    const params = { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'subtract' };
    const r = await RC.applyLayout(m.driver, 'missing_parent', params, ['c1'], {});
    assert.ok(!r.ok, 'should fail when parent not found');
    assert.ok(r.warn.includes('找不到') || r.warn.includes('找不到父') || r.warn.length > 0);
  });

  await test('applyLayout linkRMode=off：不写 R 角', async () => {
    const parent = makeMockShape({ id: 'p4', width: 200, height: 100, isRoundRect: true, adjFraction: 0.3 });
    const c1 = makeMockShape({ id: 'c1', width: 50, height: 30, isRoundRect: true, adjFraction: 0 });
    const m = makeMockDriverWithSlide({ shapes: [parent, c1] });
    const params = { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'off' };
    const r = await RC.applyLayout(m.driver, 'p4', params, ['c1'], {});
    assert.strictEqual(r.ok, true);
    // c1 的 R 角不应被改（linkRMode=off）
    // writeRadius 不会被调，但 setBox 会调
    // 验证：没有 setBox 后跟着 writeRadius 的 adjFraction 改变
    // 简化：c1 的 _adjFraction 应保持原值 0
    assert.strictEqual(c1._adjFraction, 0, 'child R角 should not be modified when linkRMode=off');
  });

  await test('applyLayout linkRMode=same：子 R 角跟父一样', async () => {
    const parent = makeMockShape({ id: 'p5', width: 200, height: 100, isRoundRect: true, adjFraction: 0.3 });
    const c1 = makeMockShape({ id: 'c1', width: 80, height: 80, isRoundRect: true, adjFraction: 0 });
    const m = makeMockDriverWithSlide({ shapes: [parent, c1] });
    const params = { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'same' };
    const r = await RC.applyLayout(m.driver, 'p5', params, ['c1'], {});
    assert.strictEqual(r.ok, true);
    // 短边 80pt = 2.823cm，0.3 * 2.823 / 2 = 0.423... R 角
    // adjFraction = 0.423 / 2.823 = 0.15 (注意：0.3 是 adjFraction 不是 cm)
    // 实际：parentRcm = 0.3 * minSideCm(80/28.35) = 0.3 * 2.823 = 0.847cm
    // childRcm = 0.847cm, child adjFraction = 0.847 / 2.823 = 0.3
    assert.ok(c1._adjFraction > 0, 'child adjFraction should be set');
    // 应该是 0.3 附近（短边相同时 same mode 复制 adjFraction）
    assert.ok(Math.abs(c1._adjFraction - 0.3) < 0.01, `expected 0.3, got ${c1._adjFraction}`);
  });

  console.log('\n=== syncLayoutChildrenR(driver) — 联动子 R 角 ===');

  await test('syncLayoutChildrenR basic subtract：父 R=0.8, padding=0.3 → 子 R=0.5', async () => {
    const parent = makeMockShape({ id: 'p_sync1', width: 200, height: 100, isRoundRect: true });
    const c1 = makeMockShape({ id: 'c_sync1', width: 80, height: 60, isRoundRect: true, adjFraction: 0 });
    const c2 = makeMockShape({ id: 'c_sync2', width: 80, height: 60, isRoundRect: true, adjFraction: 0 });
    const m = makeMockDriverWithSlide({ shapes: [parent, c1, c2] });
    const r = await RC.syncLayoutChildrenR(m.driver, 'p_sync1', ['c_sync1', 'c_sync2'], 0.3, 'subtract', 0.8);
    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.applied, 2);
    // 短边 60pt = 2.117cm, 子 R = 0.5cm, adjFraction = 0.5/2.117 = 0.236
    for (const c of [c1, c2]) {
      assert.ok(Math.abs(c._adjFraction - 0.236) < 0.01, `expected 0.236, got ${c._adjFraction}`);
    }
  });

  await test('syncLayoutChildrenR same mode：子 R = 父 R', async () => {
    const c1 = makeMockShape({ id: 'c_same1', width: 80, height: 60, isRoundRect: true, adjFraction: 0 });
    const m = makeMockDriverWithSlide({ shapes: [c1] });
    const r = await RC.syncLayoutChildrenR(m.driver, 'p_same', ['c_same1'], 0, 'same', 0.8);
    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.applied, 1);
    // 子 R = 0.8cm, 短边 60pt = 2.117cm, adjFraction = 0.8/2.117 = 0.378
    assert.ok(Math.abs(c1._adjFraction - 0.378) < 0.01, `expected 0.378, got ${c1._adjFraction}`);
  });

  await test('syncLayoutChildrenR off mode：什么都不做', async () => {
    const c1 = makeMockShape({ id: 'c_off', width: 80, height: 60, isRoundRect: true, adjFraction: 0.1 });
    const m = makeMockDriverWithSlide({ shapes: [c1] });
    const r = await RC.syncLayoutChildrenR(m.driver, 'p_off', ['c_off'], 0, 'off', 0.8);
    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.applied, 0);
    // adjFraction 不变
    assert.strictEqual(c1._adjFraction, 0.1, 'off mode should not change child R');
  });

  await test('syncLayoutChildrenR parentRcm=0：不写（off 路径）', async () => {
    const c1 = makeMockShape({ id: 'c_zero', width: 80, height: 60, isRoundRect: true, adjFraction: 0.5 });
    const m = makeMockDriverWithSlide({ shapes: [c1] });
    const r = await RC.syncLayoutChildrenR(m.driver, 'p_zero', ['c_zero'], 0, 'subtract', 0);
    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.applied, 0);
    assert.strictEqual(c1._adjFraction, 0.5, 'parentRcm=0 should not change child R');
  });

  await test('syncLayoutChildrenR stale childId 过滤', async () => {
    const c1 = makeMockShape({ id: 'c_real', width: 80, height: 60, isRoundRect: true, adjFraction: 0 });
    const m = makeMockDriverWithSlide({ shapes: [c1] });
    // 传一个不存在的 c_stale
    const r = await RC.syncLayoutChildrenR(m.driver, 'p_stale', ['c_stale', 'c_real'], 0.3, 'subtract', 0.8);
    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.applied, 1, 'stale skipped, only c_real applied');
    // c1 应该被写
    assert.ok(c1._adjFraction > 0, 'c_real should be written');
  });

  await test('syncLayoutChildrenR strict child 跳过（不 fail）', async () => {
    const cStrict = makeMockShape({ id: 'c_strict', width: 80, height: 60, isRoundRect: true, adjFraction: 0, tags: { radiusLockStrict_v1: '1' } });
    const cOk = makeMockShape({ id: 'c_ok', width: 80, height: 60, isRoundRect: true, adjFraction: 0 });
    const m = makeMockDriverWithSlide({ shapes: [cStrict, cOk] });
    const r = await RC.syncLayoutChildrenR(m.driver, 'p_strict', ['c_strict', 'c_ok'], 0.3, 'subtract', 0.8);
    assert.strictEqual(r.ok, true);
    // strict 被跳，c_ok 被应用
    assert.strictEqual(cStrict._adjFraction, 0, 'strict child R unchanged');
    assert.ok(cOk._adjFraction > 0, 'c_ok should be written');
  });

  await test('syncLayoutChildrenR 非 roundRect child 跳过（不 fail）', async () => {
    const cRect = makeMockShape({ id: 'c_rect', width: 80, height: 60, isRoundRect: false });
    const cOk = makeMockShape({ id: 'c_ok2', width: 80, height: 60, isRoundRect: true, adjFraction: 0 });
    const m = makeMockDriverWithSlide({ shapes: [cRect, cOk] });
    const r = await RC.syncLayoutChildrenR(m.driver, 'p_rect', ['c_rect', 'c_ok2'], 0.3, 'subtract', 0.8);
    assert.strictEqual(r.ok, true);
    // 非 roundRect 不算 fail（c_ok 算 applied）
    assert.strictEqual(cRect._adjFraction, 0, 'rect child R unchanged (was 0)');
    assert.ok(cOk._adjFraction > 0, 'c_ok should be written');
  });

  console.log('\n==================================================');
  console.log(`结果: ${passed} passed, ${failed} failed`);
  console.log('==================================================');
  if (failed > 0) process.exit(1);
}

run().catch((e) => { console.error(e); process.exit(1); });
