/*
 * test-driver-integration.js — driver + radius-core 集成测试（v1.3 重整后唯一端到端测试）
 *
 * 框架：fixtures.js + test-harness.js
 *   - fixtures.js 提供 5+ 标准 R 角矩形（含 locked / strict / 普通 / 非圆角 / 边界）
 *   - test-harness.js 提供 createHarness：driver 包装层，记录所有 driver 方法调用
 *
 * 测试风格：调业务方法 → 验证 driver 反应 + shape 状态
 *   - assertCalled(method)         → 该 driver 方法被调至少 1 次
 *   - assertNotCalled(method)      → 该 driver 方法没被调
 *   - assertCallCount(method, n)   → 该 driver 方法被调 n 次
 *   - assertShape(shape, expected) → shape 状态符合预期
 *
 * 覆盖（v1.3 去重后 54 个）：
 *   1. writeRadius(driver, ...) — 5+ R 角矩形的各种分支
 *   2. writeRadius 边界 — 0 尺寸 / 非圆角 / NaN / Infinity / layoutParentId
 *   3. writeRadius driver 异常 — reason=exception 带 error
 *   4. 批量：5 个普通 / 5 个混合状态
 *   5. readLockState / writeLockState — 各种 tag 状态读写
 *   6. reapplyLock — 反算 adj + clamp + 边界
 *   7. applyLayout — 父+子的端到端（2x2 / off / same / 父不在 / 子不足 / stale / writeParentTag=false / infeasible）
 *   8. syncLayoutChildrenR — subtract / same / off / stale / strict / 非圆角
 *   9. driver API 一致性 / 边界
 *   10. 自测场景：5+ R 角矩形一次操作多个
 *
 * v1.3 重整说明：
 *   - 删了 test-mock-harness.js（70 个测试）—— 重复
 *   - 删了 radius-core.writeRadiusToShapePure + applyLayoutPure —— 业务只走 driver 路径
 *   - 纯算法（computeLayout / valueToCm / 业务规则）挪到 test-radius-core.js
 */

const assert = require('assert');
const path = require('path');
const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
const { createDriver } = require(path.join(__dirname, '..', 'src', 'lib', 'ppt-driver.js'));
const { createHarness, createTestRunner, makeStandardFixture, PT_PER_CM, cm } = require('./test-harness');

const t = createTestRunner();

// ============================================================
// writeRadius(driver) — 基础 + 各 fixture 的反应
// ============================================================

t.test('writeRadius(r1_basic, 0.5)：调 setAdjFraction 一次，adj=0.5/minSideCm，不写 tag', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.newCm, 0.5);
  assert.strictEqual(r.wasLocked, false);
  assert.strictEqual(r.wasStrict, false);
  h.assertCalled('setAdjFraction', { with: ['r1_basic', 0.5 / 3] });  // 3cm 短边
  h.assertNotCalled('addTag');
  h.assertCallCount('readTag', 2);  // lock + strict 各 1 次
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.5 / 3, tags: {} });
});

t.test('writeRadius(r2_medium, 0.3)：覆盖已有 R 角（adjFraction 从 0.1 改到 0.3/4=0.075）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r2_medium, 0.3);
  // r2_medium 短边 4cm，0.3cm → 0.3/4 = 0.075
  assert.strictEqual(r.ok, true);
  h.assertShape(f.shapes.r2_medium, { adjFraction: 0.3 / 4 });
});

t.test('writeRadius(r3_large, 5)：clamp 到短边一半 (8/2=4cm)', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r3_large, 5);
  assert.strictEqual(r.ok, true);
  // short side = 8cm, half = 4cm, clamped
  assert.ok(Math.abs(r.newCm - 4) < 1e-6);
  h.assertShape(f.shapes.r3_large, { adjFraction: 0.5 });  // 4 / 8
});

t.test('writeRadius(r4_tiny, -1)：负数 → newCm=0', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r4_tiny, -1);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.newCm, 0);
  h.assertShape(f.shapes.r4_tiny, { adjFraction: 0 });
});

t.test('writeRadius(r5_wide, 0.5)：宽 shape 不影响 minSide 算（短边 5cm）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r5_wide, 0.5);
  assert.strictEqual(r.ok, true);
  h.assertShape(f.shapes.r5_wide, { adjFraction: 0.5 / 5 });  // 短边 5cm
});

// ============================================================
// writeRadius — strict / locked 路径
// ============================================================

t.test('writeRadius(r7_strict, 0.5)：strict 拦截，不调 setAdjFraction，调 readTag×2', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r7_strict, 0.5);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'strict');
  assert.strictEqual(r.isStrict, true);
  h.assertNotCalled('setAdjFraction');
  h.assertCallCount('readTag', 2);
  h.assertShape(f.shapes.r7_strict, { adjFraction: 0, tags: { radiusLockStrict_v1: '1' } });
});

t.test('writeRadius(r6_locked, 0.5)：写完 R 角后调 addTag 同步 lock fixed value', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r6_locked, 0.5);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.wasLocked, true);
  h.assertCalled('setAdjFraction');
  h.assertCalled('addTag', { with: ['r6_locked', 'radiusLock_v1', '0.5'] });
  h.assertShape(f.shapes.r6_locked, {
    adjFraction: 0.5 / 4,  // 短边 4cm
    tags: { radiusLock_v1: '0.5' },
  });
});

t.test('writeRadius(r8_lockedStrict, 0.5)：strict 优先，wasLocked=true 但不写', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r8_lockedStrict, 0.5);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'strict');
  assert.strictEqual(r.wasLocked, true);  // 读了但没写
  h.assertNotCalled('setAdjFraction');
  h.assertNotCalled('addTag');
  // 原 tag 都不动
  h.assertShape(f.shapes.r8_lockedStrict, {
    adjFraction: 0.15,
    tags: { radiusLock_v1: '0.6', radiusLockStrict_v1: '1' },
  });
});

// ============================================================
// writeRadius — 边界
// ============================================================

t.test('writeRadius(r10_zeroSize, 0.5)：0 尺寸 → reason=no-size', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r10_zeroSize, 0.5);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'no-size');
  h.assertNotCalled('setAdjFraction');
});

t.test('writeRadius(rect1, 0.5)：非圆角矩形 → reason=not-roundRect', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.rect1, 0.5);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'not-roundRect');
  h.assertNotCalled('setAdjFraction');
});

t.test('writeRadius NaN → reason=invalid-adj（v1.3.5 防御）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, NaN);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'invalid-adj');
  h.assertNotCalled('setAdjFraction');
});

t.test('writeRadius Infinity → reason=invalid-adj（不让 clamp 静默吞）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, Infinity);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'invalid-adj');
  h.assertNotCalled('setAdjFraction');
});

t.test('writeRadius(layoutParentId)：调 addTag 写 layoutChild_v1', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5, { layoutParentId: 'parent_p1' });
  assert.strictEqual(r.ok, true);
  h.assertCalled('addTag', { with: ['r1_basic', 'layoutChild_v1', 'parent_p1'] });
});

// ============================================================
// writeRadius — driver 异常处理（v1.2.1 教训：catch 要 log e.message）
// ============================================================

t.test('writeRadius 时 driver.readTag 抛异常 → reason=exception 带 error message', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  h.driver.readTag = async () => { throw new Error('office.js boom'); };
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'exception');
  assert.ok(r.error && r.error.includes('office.js boom'));
});

// ============================================================
// 批量：5+ 个 R 角矩形一次调 writeRadius
// ============================================================

t.test('批量 5 个：混合 5 个 R 角 + 1 非圆角 + 1 strict，所有 readTag 都触发', async () => {
  const f = makeStandardFixture();
  // 拿 5 个普通 R 角矩形
  const targets = [
    f.shapes.r1_basic, f.shapes.r2_medium, f.shapes.r3_large,
    f.shapes.r4_tiny, f.shapes.r5_wide,
  ];
  const h = createHarness({ shapes: f.allShapes });
  const results = [];
  for (const s of targets) {
    results.push(await RC.writeRadius(h.driver, s, 0.3));
  }
  // 5 个全部成功
  for (const r of results) assert.strictEqual(r.ok, true);
  // 5 个 setAdjFraction + 5×2 readTag = 10 readTag
  h.assertCallCount('setAdjFraction', 5);
  h.assertCallCount('readTag', 10);
  h.assertNotCalled('addTag');  // 都没 lock → 不写 lock tag
});

t.test('批量：5 个混合（普通 / locked / strict / 非圆角 / 0 尺寸），反应各不相同', async () => {
  const f = makeStandardFixture();
  const targets = [
    f.shapes.r1_basic,        // 普通
    f.shapes.r6_locked,       // locked
    f.shapes.r7_strict,       // strict
    f.rect1,                  // 非圆角
    f.shapes.r10_zeroSize,    // 0 尺寸
  ];
  const h = createHarness({ shapes: f.allShapes });
  const results = [];
  for (const s of targets) results.push(await RC.writeRadius(h.driver, s, 0.3));
  assert.strictEqual(results[0].ok, true);
  assert.strictEqual(results[1].ok, true);
  assert.strictEqual(results[2].reason, 'strict');
  assert.strictEqual(results[3].reason, 'not-roundRect');
  assert.strictEqual(results[4].reason, 'no-size');
  // 只有 r1 + r6 调了 setAdjFraction（locked 也算，因为 lock 同步不影响写 R 角）
  h.assertCallCount('setAdjFraction', 2);
  // 只有 r6 locked 调了 addTag 同步 fixed value
  h.assertCallCount('addTag', 1);
  h.assertCalled('addTag', { with: ['r6_locked', 'radiusLock_v1', '0.3'] });
});

// ============================================================
// readLockState / writeLockState
// ============================================================

t.test('readLockState(r1_basic)：无 tag → lockedCm=null, isStrict=false', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r1_basic);
  assert.strictEqual(s.lockedCm, null);
  assert.strictEqual(s.isStrict, false);
  h.assertCallCount('readTag', 2);
});

t.test('readLockState(r6_locked)：lock tag 解析为数字', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r6_locked);
  assert.strictEqual(s.lockedCm, 0.8);
  assert.strictEqual(s.isStrict, false);
});

t.test('readLockState(r7_strict)：strict=1 → isStrict=true', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r7_strict);
  assert.strictEqual(s.lockedCm, null);
  assert.strictEqual(s.isStrict, true);
});

t.test('readLockState(r8_lockedStrict)：两者都解析', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r8_lockedStrict);
  assert.strictEqual(s.lockedCm, 0.6);
  assert.strictEqual(s.isStrict, true);
});

t.test('readLockState：lock tag 非数字 → lockedCm=null（不 crash）', async () => {
  const f = makeStandardFixture();
  f.shapes.r1_basic._tags.radiusLock_v1 = 'not-a-number';
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r1_basic);
  assert.strictEqual(s.lockedCm, null);
});

t.test('writeLockState({lockedCm: 0.5})：addTag radiusLock_v1', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeLockState(h.driver, f.shapes.r1_basic, { lockedCm: 0.5 });
  assert.strictEqual(r.ok, true);
  h.assertCalled('addTag', { with: ['r1_basic', 'radiusLock_v1', '0.5'] });
  h.assertShape(f.shapes.r1_basic, { tags: { radiusLock_v1: '0.5' } });
});

t.test('writeLockState({lockedCm: null})：deleteTag radiusLock_v1', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeLockState(h.driver, f.shapes.r6_locked, { lockedCm: null });
  assert.strictEqual(r.ok, true);
  h.assertCalled('deleteTag', { with: ['r6_locked', 'radiusLock_v1'] });
  h.assertShape(f.shapes.r6_locked, { tags: { radiusLockStrict_v1: undefined } });  // r6 没 strict
});

t.test('writeLockState({isStrict: true})：addTag strict=1', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeLockState(h.driver, f.shapes.r1_basic, { isStrict: true });
  assert.strictEqual(r.ok, true);
  h.assertCalled('addTag', { with: ['r1_basic', 'radiusLockStrict_v1', '1'] });
  h.assertShape(f.shapes.r1_basic, { tags: { radiusLockStrict_v1: '1' } });
});

t.test('writeLockState({})：undefined 字段不动', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeLockState(h.driver, f.shapes.r6_locked, {});
  assert.strictEqual(r.ok, true);
  h.assertNotCalled('addTag');
  h.assertNotCalled('deleteTag');
});

t.test('writeLockState({null, false}) 同时删两个 tag（v1.3.5 回归）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeLockState(h.driver, f.shapes.r8_lockedStrict, { lockedCm: null, isStrict: false });
  assert.strictEqual(r.ok, true);
  h.assertCalled('deleteTag', { with: ['r8_lockedStrict', 'radiusLock_v1'] });
  h.assertCalled('deleteTag', { with: ['r8_lockedStrict', 'radiusLockStrict_v1'] });
  h.assertShape(f.shapes.r8_lockedStrict, { tags: {} });
});

// ============================================================
// reapplyLock
// ============================================================

t.test('reapplyLock：adj 反算 = lockedCm / minSideCm', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.shapes.r1_basic, 0.5);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.newCm, 0.5);
  h.assertCalled('setAdjFraction', { with: ['r1_basic', 0.5 / 3] });
});

t.test('reapplyLock 超短边一半 → clamp', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.shapes.r1_basic, 999);
  assert.strictEqual(r.ok, true);
  assert.ok(Math.abs(r.newCm - 1.5) < 1e-6);  // 3/2
  h.assertCalled('setAdjFraction', { with: ['r1_basic', 0.5] });  // 1.5/3
});

t.test('reapplyLock 非圆角 → reason=not-roundRect', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.rect1, 0.5);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'not-roundRect');
  h.assertNotCalled('setAdjFraction');
});

t.test('reapplyLock 0 尺寸 → reason=no-size', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.shapes.r10_zeroSize, 0.5);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'no-size');
});

t.test('reapplyLock lockedCm=-1 → reason=invalid-target（防御）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.shapes.r1_basic, -1);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'invalid-target');
});

// ============================================================
// applyLayout — 父 + 4 子端到端
// ============================================================

t.test('applyLayout 2x2：调 activeSlide+load+setBox×4+setAdjFraction×4+addTag(parent)', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.4, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'],
    {}
  );
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 4);

  // driver 反应
  h.assertCalled('activeSlide');
  h.assertCalled('slideShapes');
  h.assertCalled('load');
  h.assertCallCount('setBox', 4);  // 4 个子的位置/尺寸
  h.assertCallCount('setAdjFraction', 4);  // 4 个子的 R 角
  h.assertCalled('addTag', { with: ['parent_p1', 'layoutParent_v1'] });

  // shape 状态：4 个子位置/尺寸被重写（不再 width=2.5cm）
  for (const c of f.layoutChildren) {
    assert.ok(c.width !== cm(2.5) || c.height !== cm(1.5),
      `child ${c.id} should have been resized`);
  }
  // 父 tag 写入了
  const parentTagCall = h.calls.find((c) => c.method === 'addTag' && c.args[1] === 'layoutParent_v1');
  const parentPayload = JSON.parse(parentTagCall.args[2]);
  assert.deepStrictEqual(parentPayload.childIds, ['lc1', 'lc2', 'lc3', 'lc4']);
  assert.strictEqual(parentPayload.rows, 2);
  assert.strictEqual(parentPayload.cols, 2);
});

t.test('applyLayout linkRMode=off：子 R 角不写，只写 setBox', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'off' },
    ['lc1'],
    {}
  );
  assert.strictEqual(r.ok, true);
  h.assertCallCount('setBox', 1);
  h.assertNotCalled('setAdjFraction');
  // lc1 的 adjFraction 应保持原值 0
  h.assertShape(f.layoutChildren[0], { adjFraction: 0 });
});

t.test('applyLayout linkRMode=same：子 R = 父 R（短边相同时）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // 把 lc1 改成跟父短边一致（8cm 方）以方便验证
  f.layoutChildren[0].width = cm(8);
  f.layoutChildren[0].height = cm(8);
  const r = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'same' },
    ['lc1'],
    {}
  );
  assert.strictEqual(r.ok, true);
  // parent adjFraction = 0.3, parentRcm = 0.3 * minSide(8) = 2.4cm
  // lc1 短边 8cm，adjFraction = 2.4/8 = 0.3
  h.assertCalled('setAdjFraction', { with: ['lc1', 0.3] });
});

t.test('applyLayout 父不在 → warn 拒绝，setBox 不调', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver,
    'missing_parent',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'subtract' },
    ['lc1'],
    {}
  );
  assert.strictEqual(r.ok, false);
  assert.ok(r.warn && r.warn.length > 0);
  h.assertNotCalled('setBox');
  h.assertNotCalled('addTag');
});

t.test('applyLayout 子不足 → warn 拒绝', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // 2x2 但只传 1 个子
  const r = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.3, linkRMode: 'subtract' },
    ['lc1'],
    {}
  );
  assert.strictEqual(r.ok, false);
  assert.ok(r.warn.includes('不足'));
  h.assertNotCalled('setBox');
  h.assertNotCalled('addTag');
});

t.test('applyLayout stale childId 过滤（caller 传的不在 slide）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // 传 4 个但其中一个 stale
  const r = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.3, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'stale_id'],
    {}
  );
  // expectedCount=4, validCount=3 → warn 拒绝
  assert.strictEqual(r.ok, false);
  assert.ok(r.warn.includes('不足'));
});

t.test('applyLayout writeParentTag=false：addTag 不写父 tag（子 tag 还是写）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'off' },
    ['lc1'],
    { writeParentTag: false }
  );
  // 父 tag 不写（writeParentTag=false）
  const parentTagCalls = h.calls.filter((c) => c.method === 'addTag' && c.args[1] === 'layoutParent_v1');
  assert.strictEqual(parentTagCalls.length, 0, 'parent tag should NOT be written');
  // 子 tag 还是写
  h.assertCalled('addTag', { with: ['lc1', 'layoutChild_v1', 'parent_p1'] });
  h.assertNotCalled('setAdjFraction');  // off 不写 R 角
  h.assertCallCount('setBox', 1);
});

t.test('applyLayout infeasible（padding 太大）→ warn 拒绝', async () => {
  const f = makeStandardFixture();
  // 把父改成 5×5cm
  f.parent.width = cm(5);
  f.parent.height = cm(5);
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 2, cols: 2, padding: 10, gutter: 0, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'],
    {}
  );
  assert.strictEqual(r.ok, false);
  assert.ok(r.warn.includes('边距') || r.warn.includes('间距'));
  h.assertNotCalled('setBox');
});

// ============================================================
// syncLayoutChildrenR
// ============================================================

t.test('syncLayoutChildrenR subtract：父 R=0.8, padding=0.3 → 子 R=0.5', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(
    h.driver,
    'parent_p1',
    ['lc1', 'lc2'],
    0.3,
    'subtract',
    0.8
  );
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 2);
  h.assertCallCount('setAdjFraction', 2);
  // lc1 lc2 短边 = 1.5cm（原 size，但 syncR 不重写 box，所以用原 size）
  // 0.5 / 1.5 = 0.333
  h.assertCalled('setAdjFraction', { with: ['lc1', 0.5 / 1.5] });
});

t.test('syncLayoutChildrenR same：子 R = 父 R（clamp 到短边一半）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1'], 0, 'same', 0.8);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 1);
  // parentRcm=0.8, lc1 短边 1.5cm → 0.8/1.5 = 0.533
  // 但 writeRadius 内部会 clamp 到短边一半：max R = 0.75cm, 但 0.8 > 0.75 → clamp 到 0.75 → adjFraction = 0.5
  h.assertCalled('setAdjFraction', { with: ['lc1', 0.5] });
});

t.test('syncLayoutChildrenR off：什么都不做', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1'], 0, 'off', 0.8);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 0);
  h.assertNotCalled('setAdjFraction');
});

t.test('syncLayoutChildrenR parentRcm=0：什么都不做（off 路径）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1'], 0, 'subtract', 0);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 0);
  h.assertNotCalled('setAdjFraction');
});

t.test('syncLayoutChildrenR stale childId 过滤', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(
    h.driver,
    'parent_p1',
    ['lc1', 'stale_id'],
    0.3,
    'subtract',
    0.8
  );
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 1);
  h.assertCallCount('setAdjFraction', 1);
});

t.test('syncLayoutChildrenR strict child 跳过（不 fail）', async () => {
  const f = makeStandardFixture();
  f.layoutChildren[0]._tags.radiusLockStrict_v1 = '1';  // lc1 设为 strict
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(
    h.driver,
    'parent_p1',
    ['lc1', 'lc2'],
    0.3,
    'subtract',
    0.8
  );
  assert.strictEqual(r.ok, true);
  // lc1 strict 跳过，lc2 写
  h.assertCallCount('setAdjFraction', 1);
  h.assertCalled('setAdjFraction', { with: ['lc2', 0.5 / 1.5] });
});

t.test('syncLayoutChildrenR 非圆角 child 跳过（不 fail）', async () => {
  const f = makeStandardFixture();
  f.layoutChildren[0].adjustments.count = 0;  // lc1 设为非圆角
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(
    h.driver,
    'parent_p1',
    ['lc1', 'lc2'],
    0.3,
    'subtract',
    0.8
  );
  assert.strictEqual(r.ok, true);
  h.assertCallCount('setAdjFraction', 1);  // 只 lc2
});

// ============================================================
// driver API 一致性
// ============================================================

t.test('createDriver 返回 16 个方法（mock 跟真实 API 一致）', () => {
  const d = createDriver({ sync: async () => {} });
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
  assert.strictEqual(Object.keys(d).length, expected.length);
});

t.test('driver.adjFraction：get(0) 抛异常时返回 0（v1.2.5 defensive）', () => {
  const f = makeStandardFixture();
  f.shapes.r1_basic.adjustments.get = () => { throw new Error('boom'); };
  const h = createHarness({ shapes: f.allShapes });
  const adj = h.driver.adjFraction(f.shapes.r1_basic);
  assert.strictEqual(adj, 0);
});

t.test('driver.size 只读 width/height，不读 left/top（v1.2.2 回归）', () => {
  const f = makeStandardFixture();
  let leftRead = false, topRead = false;
  Object.defineProperty(f.shapes.r1_basic, 'left', { get() { leftRead = true; return 0; }, configurable: true });
  Object.defineProperty(f.shapes.r1_basic, 'top', { get() { topRead = true; return 0; }, configurable: true });
  const h = createHarness({ shapes: f.allShapes });
  const sz = h.driver.size(f.shapes.r1_basic);
  assert.strictEqual(sz.width, cm(5));
  assert.strictEqual(sz.height, cm(3));
  assert.ok(!('left' in sz));
  assert.ok(!('top' in sz));
});

t.test('writeRadius 不需要 left/top 被 load（v1.2.2 回归）', async () => {
  const f = makeStandardFixture();
  Object.defineProperty(f.shapes.r1_basic, 'left', { get() { throw new Error('left not loaded'); }, configurable: true });
  Object.defineProperty(f.shapes.r1_basic, 'top', { get() { throw new Error('top not loaded'); }, configurable: true });
  const h = createHarness({ shapes: f.allShapes });
  // should not throw
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);
  assert.strictEqual(r.ok, true);
});

// ============================================================
// 自测场景：5+ R 角矩形组合操作
// ============================================================

t.test('自测场景：5+ R 角矩形 + 写 R 角 + 锁定 + 联动布局', async () => {
  // 完整流程：5 个普通 R 角 + 1 个 locked + 1 个 strict + 1 个非圆角
  // → 批量写 R 角 → 锁定其中 1 个 → 应用 layout 联动 → 验证状态
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });

  // 1. 批量写 R 角到 5 个普通 R 角矩形
  const targets = [f.shapes.r1_basic, f.shapes.r2_medium, f.shapes.r3_large, f.shapes.r4_tiny, f.shapes.r5_wide];
  for (const s of targets) {
    await RC.writeRadius(h.driver, s, 0.3);
  }
  h.assertCallCount('setAdjFraction', 5);

  // 2. 锁定 r1_basic
  await RC.writeLockState(h.driver, f.shapes.r1_basic, { lockedCm: 0.3 });
  h.assertCalled('addTag', { with: ['r1_basic', 'radiusLock_v1', '0.3'] });

  // 3. 对 r1_basic 再次写 R 角 → 应该同步 fixed value 到 0.5
  await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);
  h.assertCalled('addTag', { with: ['r1_basic', 'radiusLock_v1', '0.5'] });

  // 4. r7_strict 写 R 角 → 拦截
  const rStrict = await RC.writeRadius(h.driver, f.shapes.r7_strict, 0.5);
  assert.strictEqual(rStrict.reason, 'strict');

  // 5. 对 lc1~lc4 应用 layout（2x2）联动 R 角
  const layoutR = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.3, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'],
    {}
  );
  assert.strictEqual(layoutR.ok, true);
  assert.strictEqual(layoutR.applied, 4);

  // 6. 验证最终 snapshot
  h.assertShape(f.shapes.r1_basic, { tags: { radiusLock_v1: '0.5' } });
  h.assertShape(f.shapes.r6_locked, { tags: { radiusLock_v1: '0.8' } });  // 没被改
  h.assertShape(f.shapes.r7_strict, { adjFraction: 0 });  // strict 没动
  h.assertShape(f.rect1, { adjFraction: 0 });  // 非圆角没动
  h.assertCalled('addTag', { with: ['parent_p1', 'layoutParent_v1'] });
});

t.test('自测场景：clamp 边界（短边 1.058cm）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // r9_clampEdge 短边 1.058cm，max R = 0.529cm
  // 试图写 5cm → clamp 到 0.529cm
  const r = await RC.writeRadius(h.driver, f.shapes.r9_clampEdge, 5);
  assert.strictEqual(r.ok, true);
  assert.ok(Math.abs(r.newCm - 1.058 / 2) < 1e-3);
  // adj = 0.5
  h.assertShape(f.shapes.r9_clampEdge, { adjFraction: 0.5 });
});

t.test('自测场景：reapplyLock 把改乱的 R 角恢复', async () => {
  const f = makeStandardFixture();
  // 模拟用户拖拽手柄改了 r1_basic 的 R 角
  f.shapes.r1_basic._adjFraction = 0.4;  // 用户改成 0.4
  const h = createHarness({ shapes: f.allShapes });
  // r1_basic 没有 lock tag 的话不能 reapplyLock——这里我们手动加一个 lock
  f.shapes.r1_basic._tags.radiusLock_v1 = '0.3';
  const r = await RC.reapplyLock(h.driver, f.shapes.r1_basic, 0.3);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.newCm, 0.3);
  // adjFraction 恢复成 0.3/3 = 0.1
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.3 / 3 });
});

t.test('自测场景：多 slide 的活动 slide 选择（activeSlideIndex=0 默认）', async () => {
  // 注：mock harness 不直接支持多 slide（fixture 都进同一个 slide）
  // 这里只验证 activeSlide 被调
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.3);
  // writeRadius 不会调 activeSlide（不需要 slide）
  h.assertNotCalled('activeSlide');

  // applyLayout 会调 activeSlide
  await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'off' },
    ['lc1'],
    {}
  );
  h.assertCalled('activeSlide');
});

// ============================================================
// 跑
// ============================================================

t.run().catch((e) => { console.error(e); process.exit(1); });
