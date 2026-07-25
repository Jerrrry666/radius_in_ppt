/*
 * test-features.js — 功能测试（v1.3）
 *
 * 目的：测业务函数（writeRadius / applyLayout / syncLayoutChildrenR /
 *      readLockState / writeLockState / reapplyLock）的功能正确性。
 *
 * 测试方法：模拟交互反馈
 *   - 拿 fixture（5+ R 角矩形）
 *   - 调业务方法
 *   - 验证 shape 最终状态（assertShape）
 *
 * 不测什么：
 *   - driver 层（ppt-driver.js 16 个方法）—— Mac LTSC 兼容性在真实 PPT 测
 *   - driver 内部调用次数 / 顺序 —— 功能测试不关心
 *
 * driver 调用记录（h.calls）保留在 harness 里供 debug 用，但不作为主断言。
 */

const path = require('path');
const assert = require('assert');
const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
const { createHarness, createTestRunner, makeStandardFixture, PT_PER_CM, cm } = require('./test-harness');

const t = createTestRunner();

// ============================================================
// writeRadius — 调功能后 shape 状态对不对
// ============================================================

t.test('writeRadius(普通, 0.5)：R 角变成 0.5cm（adj = 0.5/短边）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.newCm, 0.5);
  // 短边 3cm，adj = 0.5/3
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.5 / 3 });
});

t.test('writeRadius(已有 R 角, 0.3)：覆盖成新值', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeRadius(h.driver, f.shapes.r2_medium, 0.3);
  // 短边 4cm
  h.assertShape(f.shapes.r2_medium, { adjFraction: 0.3 / 4 });
});

t.test('writeRadius(超过短边一半)：R 角被 clamp 到 短边一半', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r3_large, 5);

  assert.strictEqual(r.ok, true);
  // 短边 8cm，max R = 4cm
  assert.ok(Math.abs(r.newCm - 4) < 1e-6);
  h.assertShape(f.shapes.r3_large, { adjFraction: 0.5 });  // 4/8
});

t.test('writeRadius(负数)：R 角变 0', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r4_tiny, -1);
  assert.strictEqual(r.newCm, 0);
  h.assertShape(f.shapes.r4_tiny, { adjFraction: 0 });
});

t.test('writeRadius(宽矩形)：按短边算（5cm）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeRadius(h.driver, f.shapes.r5_wide, 0.5);
  h.assertShape(f.shapes.r5_wide, { adjFraction: 0.5 / 5 });
});

t.test('writeRadius(clamp 边界 短边 1.058cm)：超过 0.529cm 都被 clamp', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r9_clampEdge, 5);
  assert.ok(Math.abs(r.newCm - 1.058 / 2) < 1e-3);
  h.assertShape(f.shapes.r9_clampEdge, { adjFraction: 0.5 });
});

// ============================================================
// writeRadius — strict / locked / 边界
// ============================================================

t.test('writeRadius(strict 形状)：R 角不变，tag 不动', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r7_strict, 0.5);

  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.reason, 'strict');
  h.assertShape(f.shapes.r7_strict, {
    adjFraction: 0,  // 没变
    tags: { radiusLockStrict_v1: '1' },  // 没动
  });
});

t.test('writeRadius(locked 形状)：R 角写入 + tag 同步到新 R 角', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r6_locked, 0.5);

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.wasLocked, true);
  h.assertShape(f.shapes.r6_locked, {
    adjFraction: 0.5 / 4,  // 短边 4cm
    tags: { radiusLock_v1: '0.5' },  // 同步
  });
});

t.test('writeRadius(locked + strict)：strict 优先，R 角和 tag 都不动', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r8_lockedStrict, 0.5);

  assert.strictEqual(r.reason, 'strict');
  assert.strictEqual(r.wasLocked, true);  // 读了但没写
  h.assertShape(f.shapes.r8_lockedStrict, {
    adjFraction: 0.15,  // 没变
    tags: { radiusLock_v1: '0.6', radiusLockStrict_v1: '1' },  // 没动
  });
});

t.test('writeRadius(0 尺寸)：拒绝', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r10_zeroSize, 0.5);
  assert.strictEqual(r.reason, 'no-size');
});

t.test('writeRadius(非圆角矩形)：拒绝', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.rect1, 0.5);
  assert.strictEqual(r.reason, 'not-roundRect');
});

t.test('writeRadius(NaN)：拒绝（v1.3.5 防御）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, NaN);
  assert.strictEqual(r.reason, 'invalid-adj');
});

t.test('writeRadius(Infinity)：拒绝（不让 clamp 静默吞）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, Infinity);
  assert.strictEqual(r.reason, 'invalid-adj');
});

t.test('writeRadius(layoutParentId)：写子 tag', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5, { layoutParentId: 'parent_p1' });
  h.assertShape(f.shapes.r1_basic, {
    adjFraction: 0.5 / 3,
    tags: { layoutChild_v1: 'parent_p1' },
  });
});

t.test('writeRadius 时 driver 抛异常：返回 reason=exception + error message', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  h.driver.readTag = async () => { throw new Error('office.js boom'); };
  const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);
  assert.strictEqual(r.reason, 'exception');
  assert.ok(r.error && r.error.includes('office.js boom'));
});

// ============================================================
// 批量写 R 角
// ============================================================

t.test('批量写 5 个普通 R 角矩形：5 个 R 角都更新', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const targets = [
    f.shapes.r1_basic, f.shapes.r2_medium, f.shapes.r3_large,
    f.shapes.r4_tiny, f.shapes.r5_wide,
  ];
  for (const s of targets) await RC.writeRadius(h.driver, s, 0.3);

  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.3 / 3 });
  h.assertShape(f.shapes.r2_medium, { adjFraction: 0.3 / 4 });
  h.assertShape(f.shapes.r3_large, { adjFraction: 0.3 / 8 });  // 注意 0.3 < 8/2=4 不 clamp
  h.assertShape(f.shapes.r4_tiny, { adjFraction: 0.3 / 1.5 });
  h.assertShape(f.shapes.r5_wide, { adjFraction: 0.3 / 5 });
});

t.test('批量写 5 个混合状态：每个 shape 反应各不相同', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const targets = [
    f.shapes.r1_basic,     // 普通
    f.shapes.r6_locked,    // locked
    f.shapes.r7_strict,    // strict
    f.rect1,               // 非圆角
    f.shapes.r10_zeroSize, // 0 尺寸
  ];
  const results = [];
  for (const s of targets) results.push(await RC.writeRadius(h.driver, s, 0.3));

  assert.strictEqual(results[0].ok, true);
  assert.strictEqual(results[1].ok, true);
  assert.strictEqual(results[2].reason, 'strict');
  assert.strictEqual(results[3].reason, 'not-roundRect');
  assert.strictEqual(results[4].reason, 'no-size');

  // 成功写 R 角的：r1 + r6（locked 也写）
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.3 / 3 });
  h.assertShape(f.shapes.r6_locked, {
    adjFraction: 0.3 / 4,
    tags: { radiusLock_v1: '0.3' },  // 同步
  });
  // 没写的：保持原状态
  h.assertShape(f.shapes.r7_strict, { adjFraction: 0 });
  h.assertShape(f.rect1, { adjFraction: 0 });
});

// ============================================================
// readLockState / writeLockState
// ============================================================

t.test('readLockState(无 tag)：lockedCm=null, isStrict=false', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r1_basic);
  assert.deepStrictEqual(s, { lockedCm: null, isStrict: false });
});

t.test('readLockState(lock tag)：lockedCm 解析为数字', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r6_locked);
  assert.deepStrictEqual(s, { lockedCm: 0.8, isStrict: false });
});

t.test('readLockState(strict tag)：isStrict=true', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r7_strict);
  assert.deepStrictEqual(s, { lockedCm: null, isStrict: true });
});

t.test('readLockState(lock + strict)：两者都解析', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r8_lockedStrict);
  assert.deepStrictEqual(s, { lockedCm: 0.6, isStrict: true });
});

t.test('readLockState(lock tag 非数字)：lockedCm=null（防御）', async () => {
  const f = makeStandardFixture();
  f.shapes.r1_basic._tags.radiusLock_v1 = 'not-a-number';
  const h = createHarness({ shapes: f.allShapes });
  const s = await RC.readLockState(h.driver, f.shapes.r1_basic);
  assert.strictEqual(s.lockedCm, null);
});

t.test('writeLockState({lockedCm: 0.5})：写 lock tag', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeLockState(h.driver, f.shapes.r1_basic, { lockedCm: 0.5 });
  h.assertShape(f.shapes.r1_basic, { tags: { radiusLock_v1: '0.5' } });
});

t.test('writeLockState({lockedCm: null})：删 lock tag', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeLockState(h.driver, f.shapes.r6_locked, { lockedCm: null });
  h.assertShape(f.shapes.r6_locked, { tags: { radiusLock_v1: undefined } });
});

t.test('writeLockState({isStrict: true})：写 strict tag', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeLockState(h.driver, f.shapes.r1_basic, { isStrict: true });
  h.assertShape(f.shapes.r1_basic, { tags: { radiusLockStrict_v1: '1' } });
});

t.test('writeLockState({})：不动（undefined 字段跳过）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeLockState(h.driver, f.shapes.r6_locked, {});
  h.assertShape(f.shapes.r6_locked, {
    tags: { radiusLock_v1: '0.8' },  // 没动
  });
});

t.test('writeLockState({null, false}) 同时删两个 tag（v1.3.5 回归）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.writeLockState(h.driver, f.shapes.r8_lockedStrict, { lockedCm: null, isStrict: false });
  h.assertShape(f.shapes.r8_lockedStrict, { tags: {} });
});

// ============================================================
// reapplyLock — 反算 adj 回 lockedCm
// ============================================================

t.test('reapplyLock(0.5cm)：R 角被设回 0.5cm', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.shapes.r1_basic, 0.5);
  assert.strictEqual(r.ok, true);
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.5 / 3 });
});

t.test('reapplyLock(超过短边一半)：clamp 到短边一半', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.shapes.r1_basic, 999);
  assert.strictEqual(r.ok, true);
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.5 });  // 1.5/3
});

t.test('reapplyLock(非圆角)：拒绝', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.rect1, 0.5);
  assert.strictEqual(r.reason, 'not-roundRect');
});

t.test('reapplyLock(0 尺寸)：拒绝', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.shapes.r10_zeroSize, 0.5);
  assert.strictEqual(r.reason, 'no-size');
});

t.test('reapplyLock(lockedCm=-1)：拒绝（防御）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.reapplyLock(h.driver, f.shapes.r1_basic, -1);
  assert.strictEqual(r.reason, 'invalid-target');
});

t.test('reapplyLock(模拟用户拖歪 R 角后)：恢复', async () => {
  const f = makeStandardFixture();
  // 模拟用户拖拽手柄改了 R 角
  f.shapes.r1_basic._adjFraction = 0.4;
  f.shapes.r1_basic._tags.radiusLock_v1 = '0.3';
  const h = createHarness({ shapes: f.allShapes });
  await RC.reapplyLock(h.driver, f.shapes.r1_basic, 0.3);
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.3 / 3 });
});

// ============================================================
// applyLayout — 父 + 子端到端
// ============================================================

t.test('applyLayout 2x2：父子位置/尺寸/R 角/父 tag 全对', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.4, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'], {}
  );

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 4);

  // 4 个子位置/尺寸被重写（不再 width=2.5cm/height=1.5cm）
  for (const c of f.layoutChildren) {
    assert.ok(c.width !== cm(2.5) || c.height !== cm(1.5),
      `child ${c.id} 应被重写`);
  }

  // 父 tag 写入了
  const parentTag = JSON.parse(f.parent._tags.layoutParent_v1);
  assert.deepStrictEqual(parentTag.childIds, ['lc1', 'lc2', 'lc3', 'lc4']);
  assert.strictEqual(parentTag.rows, 2);
  assert.strictEqual(parentTag.cols, 2);
});

t.test('applyLayout linkRMode=off：子位置/尺寸改，但 R 角不动', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const originalAdj = f.layoutChildren[0]._adjFraction;
  await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'off' },
    ['lc1'], {}
  );
  h.assertShape(f.layoutChildren[0], {
    adjFraction: originalAdj,  // 没动
  });
});

t.test('applyLayout linkRMode=same：子 R 角 = 父 R 角（按短边）', async () => {
  const f = makeStandardFixture();
  // 改成跟父短边一致以方便验证
  f.layoutChildren[0].width = cm(8);
  f.layoutChildren[0].height = cm(8);
  const h = createHarness({ shapes: f.allShapes });
  await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'same' },
    ['lc1'], {}
  );
  // parent adjFraction=0.3, parentRcm = 0.3 * 8 = 2.4cm
  // lc1 短边 8cm，adjFraction = 2.4/8 = 0.3
  h.assertShape(f.layoutChildren[0], { adjFraction: 0.3 });
});

t.test('applyLayout 父不在 → 拒绝，4 个子位置/尺寸/R 角/父 tag 都不动', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const origLc1 = { ...f.layoutChildren[0], _tags: { ...f.layoutChildren[0]._tags } };
  const origLc1Adj = f.layoutChildren[0]._adjFraction;

  const r = await RC.applyLayout(
    h.driver, 'missing_parent',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'subtract' },
    ['lc1'], {}
  );
  assert.strictEqual(r.ok, false);
  assert.ok(r.warn && r.warn.length > 0);

  // lc1 不动
  h.assertShape(f.layoutChildren[0], { adjFraction: origLc1Adj });
  assert.strictEqual(f.layoutChildren[0]._tags.layoutChild_v1, undefined);
  assert.strictEqual(f.parent._tags.layoutParent_v1, undefined);
});

t.test('applyLayout 子不足 → 拒绝，4 个子状态都不动', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const origAdj = f.layoutChildren[0]._adjFraction;

  // 2x2 但只传 1 个
  const r = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.3, linkRMode: 'subtract' },
    ['lc1'], {}
  );
  assert.strictEqual(r.ok, false);
  assert.ok(r.warn.includes('不足'));
  h.assertShape(f.layoutChildren[0], { adjFraction: origAdj });
});

t.test('applyLayout stale childId 过滤：4 个但 1 个 stale → 拒绝', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.3, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'stale_id'], {}
  );
  assert.strictEqual(r.ok, false);
  assert.ok(r.warn.includes('不足'));
});

t.test('applyLayout writeParentTag=false：父 tag 不写', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 1, cols: 1, padding: 0, gutter: 0, linkRMode: 'off' },
    ['lc1'], { writeParentTag: false }
  );
  h.assertShape(f.parent, { tags: { layoutParent_v1: undefined } });
});

t.test('applyLayout infeasible（padding 太大）：拒绝', async () => {
  const f = makeStandardFixture();
  f.parent.width = cm(5);
  f.parent.height = cm(5);
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 10, gutter: 0, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'], {}
  );
  assert.strictEqual(r.ok, false);
  assert.ok(r.warn.includes('边距') || r.warn.includes('间距'));
});

// ============================================================
// syncLayoutChildrenR — 联动子 R 角
// ============================================================

t.test('syncLayoutChildrenR subtract：父 R=0.8, padding=0.3 → 子 R=0.5', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1', 'lc2'], 0.3, 'subtract', 0.8);

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 2);
  h.assertShape(f.layoutChildren[0], { adjFraction: 0.5 / 1.5 });  // 短边 1.5cm
  h.assertShape(f.layoutChildren[1], { adjFraction: 0.5 / 1.5 });
});

t.test('syncLayoutChildrenR same：子 R = 父 R（clamp 到短边一半）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // 0.8 / 1.5 = 0.533，但 clamp 到 0.5（短边一半 = 0.75，0.8 > 0.75 → 0.5 adjFraction）
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1'], 0, 'same', 0.8);
  assert.strictEqual(r.ok, true);
  h.assertShape(f.layoutChildren[0], { adjFraction: 0.5 });
});

t.test('syncLayoutChildrenR off：什么都不做', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const origAdj = f.layoutChildren[0]._adjFraction;
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1'], 0, 'off', 0.8);
  assert.strictEqual(r.applied, 0);
  h.assertShape(f.layoutChildren[0], { adjFraction: origAdj });
});

t.test('syncLayoutChildrenR parentRcm=0：什么都不做', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const origAdj = f.layoutChildren[0]._adjFraction;
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1'], 0, 'subtract', 0);
  assert.strictEqual(r.applied, 0);
  h.assertShape(f.layoutChildren[0], { adjFraction: origAdj });
});

t.test('syncLayoutChildrenR stale childId 过滤：1 个写 1 个跳', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1', 'stale_id'], 0.3, 'subtract', 0.8);
  assert.strictEqual(r.applied, 1);
  h.assertShape(f.layoutChildren[0], { adjFraction: 0.5 / 1.5 });
});

t.test('syncLayoutChildrenR strict child 跳过：R 角不动，caller 业务照样 ok', async () => {
  const f = makeStandardFixture();
  f.layoutChildren[0]._tags.radiusLockStrict_v1 = '1';  // lc1 设为 strict
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1', 'lc2'], 0.3, 'subtract', 0.8);
  assert.strictEqual(r.ok, true);
  // strict 跳过（adjFraction 不变）
  h.assertShape(f.layoutChildren[0], { adjFraction: 0 });
  // lc2 写
  h.assertShape(f.layoutChildren[1], { adjFraction: 0.5 / 1.5 });
});

t.test('syncLayoutChildrenR 非圆角 child 跳过', async () => {
  const f = makeStandardFixture();
  f.layoutChildren[0].adjustments.count = 0;  // lc1 设为非圆角
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1', 'lc2'], 0.3, 'subtract', 0.8);
  assert.strictEqual(r.ok, true);
  h.assertShape(f.layoutChildren[0], { adjFraction: 0 });
  h.assertShape(f.layoutChildren[1], { adjFraction: 0.5 / 1.5 });
});

// ============================================================
// 自测场景：5+ R 角矩形组合操作
// ============================================================

t.test('自测：批量写 R 角 → 锁定 → 写 R 角 → 应用 layout 联动', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });

  // 1. 批量写 R 角到 5 个普通 R 角矩形
  const targets = [f.shapes.r1_basic, f.shapes.r2_medium, f.shapes.r3_large, f.shapes.r4_tiny, f.shapes.r5_wide];
  for (const s of targets) await RC.writeRadius(h.driver, s, 0.3);
  for (const s of targets) {
    const expected = 0.3 / Math.min(s.width, s.height) * PT_PER_CM;
    h.assertShape(s, { adjFraction: expected });
  }

  // 2. 锁定 r1_basic
  await RC.writeLockState(h.driver, f.shapes.r1_basic, { lockedCm: 0.3 });
  h.assertShape(f.shapes.r1_basic, { tags: { radiusLock_v1: '0.3' } });

  // 3. 再次写 r1_basic → 应该同步 fixed value 到 0.5
  await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);
  h.assertShape(f.shapes.r1_basic, {
    adjFraction: 0.5 / 3,
    tags: { radiusLock_v1: '0.5' },  // 同步
  });

  // 4. r7_strict 写 R 角 → 拦截
  const rStrict = await RC.writeRadius(h.driver, f.shapes.r7_strict, 0.5);
  assert.strictEqual(rStrict.reason, 'strict');

  // 5. 对 lc1~lc4 应用 layout（2x2）联动 R 角
  const layoutR = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.3, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'], {}
  );
  assert.strictEqual(layoutR.ok, true);
  assert.strictEqual(layoutR.applied, 4);
  // 父 tag 写入
  h.assertShape(f.parent, { tags: { layoutParent_v1: (v) => !!v } });

  // 6. 验证 r6_locked 没被改（没参与）
  h.assertShape(f.shapes.r6_locked, { tags: { radiusLock_v1: '0.8' } });
});

// ============================================================
// 跑
// ============================================================

t.run().catch((e) => { console.error(e); process.exit(1); });
