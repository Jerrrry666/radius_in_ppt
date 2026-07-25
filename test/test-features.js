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
// detectLayoutParentChanges + syncLayoutChildrenR 串联：bug #6 集成
// ============================================================
//
// 复现 bug #6：用户选父 + 子 → 在 PPT 里直接拖父的 R 角黄色滑块（不走 task pane）
//   1. monitorTick 之前：lastCm 已知父 R=0.8cm
//   2. 用户拖到 1.5cm
//   3. monitorTick 看到 currentCm=1.5，detectLayoutParentChanges 返回 [parentId, lastCm=0.8, newCm=1.5]
//   4. caller 调 syncLayoutChildrenR(driver, parentId, childIds, padding, mode, newCm=1.5)
//   5. 子 R 角应该被重写
//
// 这个测试覆盖了"monitorTick 末尾的联动 hook 逻辑"，但因为 monitorTick 在 dialog.js 里
// 不便直接测，所以用两个函数串联来证明：detectLayoutParentChanges 输出 → syncLayoutChildrenR 输入
// 两者能正确对接。

t.test('集成：bug #6 — 拖父 R 角 → detectLayoutParentChanges 触发 → syncLayoutChildrenR 写子 R', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });

  // 父初始 R = 0.3 * 8 = 2.4cm（fixture 设定），子 R = 0（fixture 设定）
  // 上次 knownCm = 2.4（monitorTick 上一轮已记）
  const knownCm = { parent_p1: 2.4 };

  // 模拟 monitorTick 看到 currentCm 变成 1.5（用户拖动）
  // selectedShapes 内存：layoutRole='parent', currentCm=1.5
  const selectedShapes = [{ id: 'parent_p1', layoutRole: 'parent', currentCm: 1.5 }];
  const changes = RC.detectLayoutParentChanges(knownCm, selectedShapes);
  assert.strictEqual(changes.length, 1);
  assert.strictEqual(changes[0].parentId, 'parent_p1');
  assert.strictEqual(changes[0].lastCm, 2.4);
  assert.strictEqual(changes[0].newCm, 1.5);

  // 模拟 dialog.js：拿到 changes 后调 syncLayoutChildrenR（用 newCm 作为 parentRcm）
  // fixture 的 layout tag 没挂在 childIds 上，所以这里手动喂
  // padding=0.3, linkRMode='subtract', parentRcm=1.5 → subRcm = max(0, 1.5 - 0.3) = 1.2
  const r = await RC.syncLayoutChildrenR(
    h.driver,
    'parent_p1',
    ['lc1', 'lc2'],
    0.3,
    'subtract',
    changes[0].newCm  // = 1.5
  );
  assert.strictEqual(r.ok, true);
  // lc1 短边 1.5cm，新 R = 1.2cm → adj = 1.2/1.5
  // 但 PowerPoint 几何约束：R 不能超过短边一半（0.75cm），clamp 到 0.75cm → adj = 0.5
  h.assertShape(f.layoutChildren[0], { adjFraction: 0.75 / 1.5 });
  h.assertShape(f.layoutChildren[1], { adjFraction: 0.75 / 1.5 });
});

t.test('集成：bug #6 — 父 R 角没变（detectLayoutParentChanges 返回空）→ 不调 syncLayoutChildrenR', async () => {
  // 模拟：monitorTick 看到 currentCm 跟上次一样（容差内）→ 不应触发联动
  const knownCm = { parent_p1: 1.0 };
  const selectedShapes = [{ id: 'parent_p1', layoutRole: 'parent', currentCm: 1.0001 }];
  const changes = RC.detectLayoutParentChanges(knownCm, selectedShapes);
  assert.strictEqual(changes.length, 0);
  // caller 应该不调 syncLayoutChildrenR（节省一次 PowerPoint.run）
});

t.test('集成：bug #6 — 首次见到（lastCm null）只记录不触发 sync', async () => {
  // 模拟：monitorTick 第一次跑（lastCm 是空）→ 不应触发 sync
  const knownCm = {};
  const selectedShapes = [{ id: 'parent_p1', layoutRole: 'parent', currentCm: 1.0 }];
  const changes = RC.detectLayoutParentChanges(knownCm, selectedShapes);
  assert.strictEqual(changes.length, 1);
  assert.strictEqual(changes[0].lastCm, null);
  // caller 应该把 lastCm 记到 1.0，但不调 syncLayoutChildrenR（避免启动时无意义重写）
});

t.test('集成：bug #6 — 父 R 角 linkRMode=off → detect 检测到变化但 syncLayoutChildrenR 跳过', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });

  // 父 R 角变了
  const knownCm = { parent_p1: 2.4 };
  const selectedShapes = [{ id: 'parent_p1', layoutRole: 'parent', currentCm: 1.5 }];
  const changes = RC.detectLayoutParentChanges(knownCm, selectedShapes);
  assert.strictEqual(changes.length, 1);

  // 但 dialog.js 里如果父的 layoutParams.linkRMode = 'off' → 调 syncLayoutChildrenR 时传 'off' → 啥也不做
  // 模拟：调用前先把子 R 角改成非 0（验证 syncLayoutChildrenR 真的不写）
  f.layoutChildren[0]._adjFraction = 0.5;  // 当前 R 角 = 0.5 * 1.5 = 0.75cm
  const r = await RC.syncLayoutChildrenR(
    h.driver,
    'parent_p1',
    ['lc1', 'lc2'],
    0.3,
    'off',  // 联动关闭
    changes[0].newCm
  );
  assert.strictEqual(r.ok, true);
  // 子的 adjFraction 应该不变（syncLayoutChildrenR linkRMode=off 直接 return）
  h.assertShape(f.layoutChildren[0], { adjFraction: 0.5 });
  h.assertShape(f.layoutChildren[1], { adjFraction: 0 });  // 之前就是 0
});

// ============================================================
// loadLayoutTags + saveLayoutTags — refreshSelection 链路
// ============================================================

t.test('loadLayoutTags: 读父 + 子 tag 完整', async () => {
  const f = makeStandardFixture();
  // 模拟 PPT 状态：父有 layout tag + 4 个子有 child tag
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = JSON.stringify({
    rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract',
    childIds: ['lc1', 'lc2', 'lc3', 'lc4'],
  });
  for (const c of f.layoutChildren) {
    c._tags[RC.LAYOUT_CHILD_TAG_KEY] = 'parent_p1';
  }
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.loadLayoutTags(h.driver, h.slide.shapes);
  assert.strictEqual(r.ok, true);
  // 父
  assert.ok(r.parents['parent_p1']);
  assert.strictEqual(r.parents['parent_p1'].rows, 2);
  assert.strictEqual(r.parents['parent_p1'].cols, 2);
  assert.deepStrictEqual(r.parents['parent_p1'].childIds, ['lc1', 'lc2', 'lc3', 'lc4']);
  // 子
  assert.strictEqual(r.childOf['lc1'], 'parent_p1');
  assert.strictEqual(r.childOf['lc2'], 'parent_p1');
  assert.strictEqual(r.childOf['lc3'], 'parent_p1');
  assert.strictEqual(r.childOf['lc4'], 'parent_p1');
  // 没 stale
  assert.deepStrictEqual(r.staleParents['parent_p1'] || [], []);
});

t.test('loadLayoutTags: stale childId 过滤（已删 / 跨 slide）', async () => {
  const f = makeStandardFixture();
  // 父 tag 里写了 4 个子，但实际只有 2 个在 slide（lc3, lc4 已被删 / 跨 slide）
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = JSON.stringify({
    rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract',
    childIds: ['lc1', 'lc2', 'lc3_stale', 'lc4_stale'],
  });
  // 子 tag 也不在了
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.loadLayoutTags(h.driver, h.slide.shapes);
  assert.strictEqual(r.ok, true);
  // parents['parent_p1'].childIds 应该只剩 lc1 + lc2（stale 被过滤）
  assert.deepStrictEqual(r.parents['parent_p1'].childIds, ['lc1', 'lc2']);
  // staleParents 报告哪些被剔除了
  assert.deepStrictEqual(r.staleParents['parent_p1'], ['lc3_stale', 'lc4_stale']);
});

t.test('loadLayoutTags: 父 tag 损坏 → 跳过（不 throw）', async () => {
  const f = makeStandardFixture();
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = 'garbage-not-json';
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.loadLayoutTags(h.driver, h.slide.shapes);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.parents['parent_p1'], undefined);  // 解析失败 → 没 parents
});

t.test('loadLayoutTags: 父 tag 缺 childIds → 跳过', async () => {
  const f = makeStandardFixture();
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = JSON.stringify({ rows: 2, cols: 2 });
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.loadLayoutTags(h.driver, h.slide.shapes);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.parents['parent_p1'], undefined);
});

t.test('loadLayoutTags: 兼容旧版 linkR boolean', async () => {
  const f = makeStandardFixture();
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = JSON.stringify({
    rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkR: false, childIds: ['lc1'],
  });
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.loadLayoutTags(h.driver, h.slide.shapes);
  assert.strictEqual(r.parents['parent_p1'].linkRMode, 'off');
});

t.test('saveLayoutTags: 写父 + 子 tag 完整', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.saveLayoutTags(
    h.driver, h.slide,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4']
  );
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.writtenChildIds, ['lc1', 'lc2', 'lc3', 'lc4']);
  assert.deepStrictEqual(r.staleChildIds, []);
  // 父 tag 写了
  const parentTag = JSON.parse(f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY]);
  assert.strictEqual(parentTag.rows, 2);
  assert.strictEqual(parentTag.cols, 2);
  assert.strictEqual(parentTag.linkRMode, 'subtract');
  assert.deepStrictEqual(parentTag.childIds, ['lc1', 'lc2', 'lc3', 'lc4']);
  // 子 tag 写了
  for (const c of f.layoutChildren) {
    assert.strictEqual(c._tags[RC.LAYOUT_CHILD_TAG_KEY], 'parent_p1');
  }
});

t.test('saveLayoutTags: stale childId 跳过（不写 tag，不报错）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.saveLayoutTags(
    h.driver, h.slide,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'stale_id_1', 'stale_id_2']
  );
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.writtenChildIds, ['lc1', 'lc2']);
  assert.deepStrictEqual(r.staleChildIds, ['stale_id_1', 'stale_id_2']);
  // 父 tag 里 childIds 只剩 valid 的
  const parentTag = JSON.parse(f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY]);
  assert.deepStrictEqual(parentTag.childIds, ['lc1', 'lc2']);
  // 子 tag 只写了 lc1, lc2
  assert.strictEqual(f.layoutChildren[0]._tags[RC.LAYOUT_CHILD_TAG_KEY], 'parent_p1');
  assert.strictEqual(f.layoutChildren[1]._tags[RC.LAYOUT_CHILD_TAG_KEY], 'parent_p1');
});

t.test('saveLayoutTags: 父不在当前 slide → 返回错误，不写任何 tag', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.saveLayoutTags(
    h.driver, h.slide,
    'parent_NOT_EXIST',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    ['lc1']
  );
  assert.strictEqual(r.ok, false);
  // 任何 tag 都不应该被写
  assert.strictEqual(f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY], undefined);
  assert.strictEqual(f.layoutChildren[0]._tags[RC.LAYOUT_CHILD_TAG_KEY], undefined);
});

t.test('集成：saveLayoutTags → loadLayoutTags round-trip 一致', async () => {
  // 写完再读，验证数据完整保留
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  await RC.saveLayoutTags(
    h.driver, h.slide,
    'parent_p1',
    { rows: 3, cols: 2, padding: 0.4, gutter: 0.1, linkRMode: 'same' },
    ['lc1', 'lc2', 'lc3', 'lc4']
  );
  // 重置 calls（harness 内部 recordCall 全程都在）
  h.reset();
  const r = await RC.loadLayoutTags(h.driver, h.slide.shapes);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.parents['parent_p1'].rows, 3);
  assert.strictEqual(r.parents['parent_p1'].cols, 2);
  assert.strictEqual(r.parents['parent_p1'].linkRMode, 'same');
  assert.deepStrictEqual(r.parents['parent_p1'].childIds, ['lc1', 'lc2', 'lc3', 'lc4']);
  assert.strictEqual(r.childOf['lc1'], 'parent_p1');
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
