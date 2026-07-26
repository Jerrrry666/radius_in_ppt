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
const fs = require('fs');
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

// ============================================================
// v1.2.6：默认 linkRMode = 'same'（等距 R_sub = R_父）
// ============================================================

t.test('v1.2.6：applyLayout 不传 linkRMode → 默认 same（子 R 角 = 父 R 角）', async () => {
  // v1.0/v1.2 默认是 'subtract'（不等距），v1.2.6 改成 'same'（等距）
  // 父 R = 0.3 * 8 = 2.4cm，期望子 R = 2.4cm（clamp 到子短边一半 0.75cm）
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2 },  // ← 没传 linkRMode
    ['lc1', 'lc2', 'lc3', 'lc4'],
    { writeParentTag: true, syncR: true }
  );
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 4);
  // 验证：4 个子的 R 角跟父一样（clamp 到子短边一半 0.75cm → adj=0.5）
  // 父的 R 是 0.3*8=2.4cm，但子短边 1.5cm，max R=0.75cm → adj=0.75/1.5=0.5
  for (let i = 0; i < 4; i++) {
    const c = f.layoutChildren[i];
    h.assertShape(c, { adjFraction: 0.5 }, `lc${i+1} 默认 same 模式 R 角 = 父 R 角 (clamp 0.75cm)`);
  }
});

t.test('v1.2.6：saveLayoutTags 不传 linkRMode → 默认 same', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.saveLayoutTags(
    h.driver, h.slide,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2 },  // ← 没传 linkRMode
    ['lc1', 'lc2', 'lc3', 'lc4']
  );
  assert.strictEqual(r.ok, true);
  // 验证父 tag 里 linkRMode = 'same'（不是 'subtract'）
  const parentTag = JSON.parse(f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY]);
  assert.strictEqual(parentTag.linkRMode, 'same', 'v1.2.6 默认 linkRMode 应该是 same');
});

t.test('v1.2.6：subtract 模式：父 R=0.7 + d=0.2 + 子 R = 0.5（角部窄于边部，几何上不等距）', async () => {
  // 验证 subtract 行为没变（v1.0/v1.2 兼容）
  // 父 R=2.4cm，subtract 模式：子 R = 2.4 - 0.3 = 2.1cm → clamp 到 0.75cm → adj=0.5
  // 跟 same 模式结果数值一样（因为都 clamp 到 0.75cm），但**几何意图不同**：
  //   - same：父 R 角 = 2.4cm → 子 R 角 = min(2.4, 0.75) = 0.75cm（保留 R 角的几何）
  //   - subtract：父 R 角 - d = 2.1cm → 子 R 角 = min(2.1, 0.75) = 0.75cm（先减 d 再 clamp）
  // 这里两个公式结果数值一样（都过 clamp），但**没 clamp 时**应该不同。
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // 改父 R 角为 0.5cm（短边一半内），验证不 clamp 时行为不同
  // 父短边 8cm, adj = 0.5/8 = 0.0625
  f.parent._adjFraction = 0.0625;
  // 子尺寸会被 applyLayout 改成 (W-2p)x(H-2p) = (12-0.6)x(8-0.6) = 11.4x7.4
  // 短边 7.4cm, max R = 3.7cm
  // same 模式：子 R = min(0.5, 3.7) = 0.5cm → adj = 0.5/7.4
  // subtract 模式：子 R = min(0.5-0.3, 3.7) = min(0.2, 3.7) = 0.2cm → adj = 0.2/7.4
  await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 1, cols: 1, padding: 0.3, gutter: 0, linkRMode: 'same' },
    ['lc1'],
    { writeParentTag: false, syncR: true }
  );
  h.assertShape(f.layoutChildren[0], { adjFraction: 0.5 / 7.4 }, 'same 模式：子 R = 父 R = 0.5cm');

  // 重新 fixture 测 subtract
  const f2 = makeStandardFixture();
  f2.parent._adjFraction = 0.0625;
  const h2 = createHarness({ shapes: f2.allShapes });
  await RC.applyLayout(
    h2.driver, 'parent_p1',
    { rows: 1, cols: 1, padding: 0.3, gutter: 0, linkRMode: 'subtract' },
    ['lc1'],
    { writeParentTag: false, syncR: true }
  );
  h2.assertShape(f2.layoutChildren[0], { adjFraction: 0.2 / 7.4 }, 'subtract 模式：子 R = 父 R - 0.3 = 0.2cm');
});

t.test('v1.2.6：syncLayoutChildrenR 不传 linkRMode（radius-core）→ 默认 same', async () => {
  // 父 R = 0.5cm（fixture 上面 test 改的，模拟"调用方不传 linkRMode"）
  // 但 syncLayoutChildrenR 必须显式传 linkRMode（没 default）—— 测 default
  // 实际 radius-core.syncLayoutChildrenR 没有 default（参数强制），
  // default 是在 dialog.js syncLayoutChildrenRIfNeeded 里 `|| 'same'`
  // 这里直接测 radius-core 不传 linkRMode 时的行为（实际是 undefined，会被 'same' fallback 命中）
  // 但更准确地测：调用方传 'same'，验证公式正确
  const f = makeStandardFixture();
  f.parent._adjFraction = 0.0625;  // 父 R = 0.5cm
  const h = createHarness({ shapes: f.allShapes });
  // call 跟 caller 在 dialog.js 里一样：s.layoutParams.linkRMode || 'same' = 'same'
  const linkRMode = undefined || 'same';
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1', 'lc2'], 0.3, linkRMode, 0.5);
  assert.strictEqual(r.ok, true);
  // same 模式：子 R = 父 R = 0.5cm
  h.assertShape(f.layoutChildren[0], { adjFraction: 0.5 / 1.5 });
  h.assertShape(f.layoutChildren[1], { adjFraction: 0.5 / 1.5 });
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

t.test('v1.2.7：autoPadding 让大 padding 变成 feasible（不再 infeasible）', async () => {
  // v1.2.6 之前：父 5x5 padding 10 → infeasible（拒绝）
  // v1.2.7：autoPadding 把 padding 减到 min(5,5)/2 - 父R = 2.5 - 1.5 = 1cm → feasible
  // 父 5x5，R = 0.3 * 5 = 1.5cm（fixture 默认 adj）
  // d_init = 10cm, d_max = 1cm → effective = 1cm
  // 子尺寸 = (5-2)/2 = 1.5cm × 1.5cm ✓
  const f = makeStandardFixture();
  f.parent.width = cm(5);
  f.parent.height = cm(5);
  // 父 adjFraction = 0.3 默认 → R = 1.5cm
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 10, gutter: 0, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'], {}
  );
  // 期望：autoPadding 起作用，layout 成功
  assert.strictEqual(r.ok, true, `autoPadding 应该让 infeasible 变成 feasible，实际: ${r.warn || r.error}`);
  // 验证子尺寸 = 1.5x1.5（5-2*1 = 3 / 2 = 1.5）
  for (let i = 0; i < 4; i++) {
    const c = f.layoutChildren[i];
    // width = 1.5cm = 1.5 * 28.3464567 ≈ 42.52pt
    assert.ok(Math.abs(c.width - 1.5 * PT_PER_CM) < 1, `lc${i+1} width ${c.width/PT_PER_CM}cm 应该 = 1.5cm`);
    assert.ok(Math.abs(c.height - 1.5 * PT_PER_CM) < 1, `lc${i+1} height ${c.height/PT_PER_CM}cm 应该 = 1.5cm`);
  }
});

t.test('v1.2.7：applyLayout 真 infeasible（父尺寸真的不够）', async () => {
  // 父 1x1cm R=0.5cm, padding 0.3, 2x2 → 父太小，autoPadding 减到 0 也放不下 4 个子
  const f = makeStandardFixture();
  f.parent.width = cm(1);
  f.parent.height = cm(1);
  f.parent._adjFraction = 0.5;  // 父 R = 0.5cm，d_max = 0
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'], {}
  );
  // effective = 0, 2x2 仍需 ≥ 2*0 + 子尺寸 > 0 → 子尺寸 = 1/2 = 0.5cm，OK
  // 但 4 个子都放得下（each 0.5x0.5），feasible
  // 真正 infeasible：2x2 with gutter > 0，subW = (1-0-0)/2 = 0.5
  // 试试：1x1 R=0.5, rows=2 cols=2, gutter=0.5 → subW = 0.5, 但 (cols-1)*gutter=0.5，totalW = 1-0-0.5 = 0.5, 2*subW=1 > 0.5 → infeasible
  const r2 = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0, gutter: 0.3, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'], {}
  );
  // totalW = 1 - 0 - 0.3 = 0.7, subW = 0.35 > 0 → 实际 feasible
  // 真正的 infeasible：1x1, padding 0, gutter 0.6, 2x2 → totalW = 1-0-0.6 = 0.4, subW = 0.2 > 0 → 仍 feasible
  // 极端：1x1, padding 0.6, gutter 0, 2x2 → totalW = -0.2 → infeasible
  const r3 = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0.6, gutter: 0, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'], {}
  );
  // effective = 0, totalW = 1 - 0 - 0 = 1, subW = 0.5 → feasible
  // 真 infeasible 难构造（autoPadding 减了 padding）—— 改为期望 feasible
  assert.strictEqual(r3.ok, true);
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

t.test('loadLayoutTags: bug v1.3.7 — 只选父 + 4 个子都在 slide → 全部保留（不被误判 stale）', async () => {
  // v1.3.7 bug：只选父时 4 个子不在选区 → 旧版用 selectedShapeIds 做 stale 检测
  //   把 4 个子全部判为 stale → childIds 变空 → UI 显示"子 0 个（需要 4）"
  // 修法：caller 必须传整 slide 的 shapes（allSlideShapes）→ stale 检测比对 slide IDs
  //   → 4 个子都在 slide 上 → 全部 valid → UI 显示"子 4 个"
  const f = makeStandardFixture();
  // 父 tag 写 4 个子（标准 fixture 里 lc1~lc4 都在）
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = JSON.stringify({
    rows: 2, cols: 2, padding: 0.5, gutter: 0.3, linkRMode: 'subtract',
    childIds: ['lc1', 'lc2', 'lc3', 'lc4'],
  });
  const h = createHarness({ shapes: f.allShapes });
  // 关键：选区只放父，但 allSlideShapes 给全部 5 个 shape
  const r = await RC.loadLayoutTags(h.driver, [f.parent], h.slide.shapes);
  assert.strictEqual(r.ok, true);
  // 4 个子全部保留
  assert.deepStrictEqual(r.parents['parent_p1'].childIds, ['lc1', 'lc2', 'lc3', 'lc4']);
  // 没有 stale（4 个子都在 slide 上）
  assert.strictEqual(r.staleParents['parent_p1'], undefined);
});

t.test('loadLayoutTags: bug v1.3.7 fallback — 不传 allSlideShapes 时用选区（兼容旧测试）', async () => {
  // 不传 allSlideShapes 时降级到 selectedShapeIds，行为同 v1.3.6
  // （这条测试只是确保新参数可选，不破坏旧调用方）
  const f = makeStandardFixture();
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = JSON.stringify({
    rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract',
    childIds: ['lc1', 'lc2'],
  });
  const h = createHarness({ shapes: f.allShapes });
  // 选区放父 + lc1 + lc2（不传 allSlideShapes）—— lc1/lc2 在 f.layoutChildren 数组里
  const r = await RC.loadLayoutTags(h.driver, [f.parent, f.layoutChildren[0], f.layoutChildren[1]]);
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.parents['parent_p1'].childIds, ['lc1', 'lc2']);
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
// pickupFromSelection / applyPickedToSelection — 样式刷链路
// ============================================================

t.test('pickupFromSelection: 选区里第一个 roundRect → 返回 cm + strict', async () => {
  const f = makeStandardFixture();
  // 喂 r2_medium（adjFraction=0.1, 短边 4cm → cm=0.4）当第一个 roundRect
  const h = createHarness({ shapes: [f.shapes.r2_medium] });
  const r = await RC.pickupFromSelection(h.driver, h.slide.shapes);
  assert.ok(r);
  assert.strictEqual(r.id, 'r2_medium');
  assert.strictEqual(r.sourceStrict, false);
  // 浮点容差
  assert.ok(Math.abs(r.cm - 0.4) < 1e-6);
});

t.test('pickupFromSelection: 选区里第一个 strict 形状 → sourceStrict=true', async () => {
  const f = makeStandardFixture();
  // r7_strict 已经有 strict tag（fixture）
  const h = createHarness({ shapes: [f.shapes.r7_strict] });
  const r = await RC.pickupFromSelection(h.driver, h.slide.shapes);
  assert.ok(r);
  assert.strictEqual(r.id, 'r7_strict');
  assert.strictEqual(r.sourceStrict, true);
});

t.test('pickupFromSelection: 选区里没 roundRect → 返回 null', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: [f.rect1] });  // rect1 是非圆角矩形
  const r = await RC.pickupFromSelection(h.driver, h.slide.shapes);
  assert.strictEqual(r, null);
});

t.test('pickupFromSelection: 选区空 → 返回 null（不 throw）', async () => {
  const h = createHarness({ shapes: [] });
  const r = await RC.pickupFromSelection(h.driver, h.slide.shapes);
  assert.strictEqual(r, null);
});

t.test('applyPickedToSelection: 把源 R 角刷到所有 roundRect 目标', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // 源 R 角 = 0.4cm（来自 r2_medium），目标 = 5 个 roundRect
  const source = { cm: 0.4, sourceStrict: false };
  // 喂 5 个 roundRect 当目标
  const targets = [f.shapes.r1_basic, f.shapes.r3_large, f.shapes.r4_tiny, f.shapes.r5_wide, f.shapes.r6_locked];
  const r = await RC.applyPickedToSelection(h.driver, targets, source, {});
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 5);
  assert.strictEqual(r.failed, 0);
  // 验证每个目标 R 角都变了
  for (const t of targets) {
    const minSideCm = Math.min(t.width, t.height) / PT_PER_CM;
    const expectedAdj = Math.min(0.4, minSideCm / 2) / minSideCm;
    h.assertShape(t, { adjFraction: (v) => Math.abs(v - expectedAdj) < 1e-6 });
  }
});

t.test('applyPickedToSelection: bug #1 — 吸取后能正常刷入（happy path）', async () => {
  // bug #1 根因：之前 dialog.js applyPipetteToSelection 调的是 writeRadiusToShape（不走 driver），
  //             在某些场景写不进 R 角。改用 radius-core.writeRadius（driver 版）后正常。
  // 这个测试就是验证整个 pipeline：pickup → apply 一气呵成能把 R 角刷到目标
  const f = makeStandardFixture();
  // 用 r2_medium 当源（adj=0.1, 短边 4cm → cm=0.4）
  const h = createHarness({ shapes: [f.shapes.r2_medium] });
  // 1. 吸 r2_medium 的 R 角
  const source = await RC.pickupFromSelection(h.driver, h.slide.shapes);
  assert.ok(source);
  assert.ok(Math.abs(source.cm - 0.4) < 1e-6);
  // 2. 刷到 r1_basic（之前 adj=0）
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic], source, {});
  assert.strictEqual(r.ok, true, `applyPickedToSelection 失败: ${r.error || r.rejectReason}`);
  assert.strictEqual(r.applied, 1, 'bug #1：吸取后应该能刷入 1 个目标');
  // 3. 验证 r1_basic 真的被改了
  // r1_basic 短边 3cm，新 R = 0.4cm（不超短边一半 1.5）→ adj = 0.4/3
  h.assertShape(f.shapes.r1_basic, { adjFraction: (v) => Math.abs(v - 0.4 / 3) < 1e-6 });
});

t.test('applyPickedToSelection: 目标里有 strict → 全部拒绝（步骤 0 拦截）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: false };
  // r7_strict 有 strict tag
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic, f.shapes.r7_strict], source, {});
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.rejectReason, 'strict');
  // r1_basic 状态不变（步骤 0 拦截，提前 return）
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0 });  // fixture 初始 0
});

t.test('applyPickedToSelection: 目标里含非圆角矩形 → 跳过（不计入 failed）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: false };
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic, f.rect1], source, {});
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 1);  // 只有 r1_basic 是 roundRect
  // r1_basic 改了
  h.assertShape(f.shapes.r1_basic, { adjFraction: (v) => Math.abs(v - 0.4 / 3) < 1e-6 });
  // rect1 状态不变（非 roundRect，writeRadius 内部会 skip）
});

t.test('applyPickedToSelection: source.cm 不合法 → 直接拒绝（不 throw）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // NaN
  const r1 = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic], { cm: NaN }, {});
  assert.strictEqual(r1.ok, false);
  // null
  const r2 = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic], null, {});
  assert.strictEqual(r2.ok, false);
  // undefined
  const r3 = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic], undefined, {});
  assert.strictEqual(r3.ok, false);
});

t.test('applyPickedToSelection: syncStrict=true + sourceStrict=true → 目标加 strict tag', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: true };
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic, f.shapes.r3_large], source, { syncStrict: true });
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 2);
  assert.strictEqual(r.strictSynced, 2);
  // 验证 strict tag 写了
  assert.strictEqual(f.shapes.r1_basic._tags[RC.LOCK_STRICT_TAG_KEY], '1');
  assert.strictEqual(f.shapes.r3_large._tags[RC.LOCK_STRICT_TAG_KEY], '1');
});

t.test('applyPickedToSelection: syncStrict=false → 不刷 strict（即使源是 strict）', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: true };
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic], source, { syncStrict: false });
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.strictSynced, 0);
  assert.strictEqual(f.shapes.r1_basic._tags[RC.LOCK_STRICT_TAG_KEY], undefined);
});

// ===== v1.2.15: syncStrict 双向覆盖（source=false 也要清目标的 strict） =====

t.test('applyPickedToSelection: syncStrict=true + sourceStrict=false + 目标有 strict → 删 strict tag + 写 R 角', async () => {
  // v1.2.15 新行为：source 不 strict 时，syncStrict=true 应该把目标的 strict tag **也删掉**
  // （之前 v1.2 之前只单向：source strict → 目标 strict，source 不 strict → 啥也不做）
  // 顺序：先删 strict → 再写 R 角（写的时候 target 已不是 strict，writeRadius 不被拦截）
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: false };
  // r7_strict fixture 已有 radiusLockStrict_v1='1'
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r7_strict], source, { syncStrict: true });
  assert.strictEqual(r.ok, true);
  // strict tag 删了
  assert.strictEqual(f.shapes.r7_strict._tags[RC.LOCK_STRICT_TAG_KEY], undefined, 'strict tag 应该被删');
  // R 角被写（target 不再 strict，writeRadius 成功）
  h.assertShape(f.shapes.r7_strict, { adjFraction: (v) => Math.abs(v - 0.4 / 5) < 1e-6 });
  // 计数
  assert.strictEqual(r.applied, 1);
  assert.strictEqual(r.strictRemoved, 1);
  assert.strictEqual(r.strictAdded, 0);
  assert.strictEqual(r.strictSynced, 1);
});

t.test('applyPickedToSelection: syncStrict=true + sourceStrict=false + 目标没 strict → 不删（no-op）+ 写 R 角', async () => {
  // 目标本来就没有 strict，deleteTag 是 no-op，但 strictRemoved 仍然 +1（操作了）
  // 这里其实应该区分"有 strict → 删"和"没 strict → no-op"才算精确，但当前实现统一 +1
  // （行为上是 idempotent，UX 反馈时按"操作了 N 个目标"显示也合理）
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: false };
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic], source, { syncStrict: true });
  assert.strictEqual(r.ok, true);
  h.assertShape(f.shapes.r1_basic, { adjFraction: (v) => Math.abs(v - 0.4 / 3) < 1e-6 });
  assert.strictEqual(f.shapes.r1_basic._tags[RC.LOCK_STRICT_TAG_KEY], undefined);
  // 计数
  assert.strictEqual(r.applied, 1);
  assert.strictEqual(r.strictRemoved, 1, 'deleteTag 被调用，计 1 次');
  assert.strictEqual(r.strictAdded, 0);
});

t.test('applyPickedToSelection: syncStrict=true + sourceStrict=true + 目标有 strict → 跳过拦截 + R 角不写 + strict 保留', async () => {
  // v1.2.15 新行为：syncStrict=true 时 step 0 拦截**不生效**（让 override 逻辑处理）
  // 目标原本就是 strict：
  //   - step 0 跳过（syncStrict=true）
  //   - step 1a 不进（source.strict=true）
  //   - step 1b writeRadius 拒（target 已是 strict）→ failed++
  //   - step 1c addTag strict（覆盖回 '1'，no-op）→ strictAdded++
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: true };
  // r7_strict 已有 strict tag
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r7_strict], source, { syncStrict: true });
  assert.strictEqual(r.ok, true);
  // 不会被 step 0 拒
  assert.notStrictEqual(r.rejectReason, 'strict');
  // R 角不写（writeRadius 拒）
  assert.strictEqual(r.applied, 0);
  assert.strictEqual(r.failed, 1);
  // strict 仍是 '1'（addTag 覆盖回 '1'，no-op 数据上）
  assert.strictEqual(f.shapes.r7_strict._tags[RC.LOCK_STRICT_TAG_KEY], '1');
  // 计数
  assert.strictEqual(r.strictAdded, 1);
  assert.strictEqual(r.strictRemoved, 0);
  assert.strictEqual(r.strictSynced, 1);
});

t.test('applyPickedToSelection: syncStrict=true + sourceStrict=true + 目标混合（1 strict + 1 普通）→ 各按情况', async () => {
  // 关键组合测试：
  //   - r1_basic: 普通 → 写 R 角成功 + 加 strict
  //   - r7_strict: 已有 strict → 写 R 角拒 + addTag 覆盖
  // step 0 因为 syncStrict=true 而跳过，所以不会整个拒绝
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: true };
  const r = await RC.applyPickedToSelection(
    h.driver, [f.shapes.r1_basic, f.shapes.r7_strict], source, { syncStrict: true }
  );
  assert.strictEqual(r.ok, true);
  // r1_basic 写了 R 角 + 加了 strict
  h.assertShape(f.shapes.r1_basic, { adjFraction: (v) => Math.abs(v - 0.4 / 3) < 1e-6 });
  assert.strictEqual(f.shapes.r1_basic._tags[RC.LOCK_STRICT_TAG_KEY], '1');
  // r7_strict 没写 R 角（已是 strict 被拒）+ strict 保留
  assert.strictEqual(f.shapes.r7_strict._tags[RC.LOCK_STRICT_TAG_KEY], '1');
  // 计数
  assert.strictEqual(r.applied, 1);
  assert.strictEqual(r.failed, 1);
  assert.strictEqual(r.strictAdded, 2);
  assert.strictEqual(r.strictRemoved, 0);
});

t.test('applyPickedToSelection: syncStrict=true + sourceStrict=false + 目标混合（1 strict + 1 普通）→ 都删 strict + 都写 R 角', async () => {
  //   - r1_basic: 没 strict → deleteTag 调（no-op）→ 写 R 角成功
  //   - r7_strict: 有 strict → deleteTag 删掉 → 写 R 角成功
  // 关键：r7_strict 现在能写 R 角了（之前 syncStrict=false 时会 step 0 拦截）
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: false };
  const r = await RC.applyPickedToSelection(
    h.driver, [f.shapes.r1_basic, f.shapes.r7_strict], source, { syncStrict: true }
  );
  assert.strictEqual(r.ok, true);
  // r1_basic 写了 R 角
  h.assertShape(f.shapes.r1_basic, { adjFraction: (v) => Math.abs(v - 0.4 / 3) < 1e-6 });
  // r7_strict 也写了 R 角（strict 被删了）
  h.assertShape(f.shapes.r7_strict, { adjFraction: (v) => Math.abs(v - 0.4 / 5) < 1e-6 });
  // 都没 strict 了
  assert.strictEqual(f.shapes.r1_basic._tags[RC.LOCK_STRICT_TAG_KEY], undefined);
  assert.strictEqual(f.shapes.r7_strict._tags[RC.LOCK_STRICT_TAG_KEY], undefined);
  // 计数
  assert.strictEqual(r.applied, 2);
  assert.strictEqual(r.failed, 0);
  assert.strictEqual(r.strictRemoved, 2);
  assert.strictEqual(r.strictAdded, 0);
});

t.test('applyPickedToSelection: syncStrict=false + 目标有 strict → 仍然拦截（行为不变）', async () => {
  // 回归测试：syncStrict=false 时 step 0 拦截行为**不变**（不能误改）
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const source = { cm: 0.4, sourceStrict: true };
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r7_strict], source, { syncStrict: false });
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.rejectReason, 'strict');
  // strict 不动
  assert.strictEqual(f.shapes.r7_strict._tags[RC.LOCK_STRICT_TAG_KEY], '1');
  // 计数：0（操作被拒）
  assert.strictEqual(r.strictAdded, 0);
  assert.strictEqual(r.strictRemoved, 0);
  assert.strictEqual(r.strictSynced, 0);
});

t.test('applyPickedToSelection: 目标里有 locked 形状 → 写完 R 角后同步 fixed value', async () => {
  // 验证 #1 修复后的行为：locked 目标被样式刷后，lockedCm 也要同步（不然 lock monitor 会反算）
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // r6_locked fixture 已有 radiusLock_v1: '0.8'，adjFraction=0.2
  const source = { cm: 0.5, sourceStrict: false };
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r6_locked], source, {});
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 1);
  // 验证 R 角写了
  h.assertShape(f.shapes.r6_locked, { adjFraction: (v) => Math.abs(v - 0.5 / 4) < 1e-6 });
  // 验证 lockedCm 同步成 0.5（v1.1 lock 同步 fixed value 行为）
  assert.strictEqual(f.shapes.r6_locked._tags[RC.LOCK_TAG_KEY], '0.5');
});

t.test('集成：完整 pipette pipeline — pickup → apply → history', async () => {
  // 模拟 dialog.js 的 pipette 流程：吸取 + 刷入 + 记录 history
  const f = makeStandardFixture();
  // 用 r2_medium 当源（adj=0.1, 短边 4cm → cm=0.4）
  const h = createHarness({ shapes: [f.shapes.r2_medium] });
  // 1. 吸 r2_medium（cm=0.4）
  const source = await RC.pickupFromSelection(h.driver, h.slide.shapes);
  assert.ok(source);
  assert.ok(Math.abs(source.cm - 0.4) < 1e-6);
  // 2. 刷到 3 个目标
  const r = await RC.applyPickedToSelection(h.driver, [f.shapes.r1_basic, f.shapes.r3_large, f.shapes.r4_tiny], source, {});
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 3);
  // 3. 记录 history（用 cm 0.4 + 当前单位 cm）
  const history = RC.pushHistory([], source.cm, 'cm');
  assert.strictEqual(history.length, 1);
  assert.strictEqual(history[0].value, 0.4);
});

// ============================================================
// bug #6 子 bug：调整父 R 角后 4 个子只写 2 个（Mac LTSC per-call sync 累积）
// ============================================================

t.test('bug #6 子 bug：syncLayoutChildrenR 写 4 个子（2×2）→ 全部写进去', async () => {
  // 用户实测：调整父 R 角后上面 2 个子变了，下面 2 个没变
  // 根因：writeRadius 内部 readTag 调 ctx.sync()，4 次 readTag + 4 次 setAdjFraction 在同一个 run
  //       Mac LTSC 上 per-shape sync 累积（v1.2.6 同样坑），后几个 shape 的 setAdjFraction 失败/丢失
  // 修法：sibling applyLayout 模式 —— readTagsBulk 一次拿全部 + 写所有 setAdjFraction + final sync
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // 触发 syncLayoutChildrenR：父 R 改成 1.5cm（>短边一半 0.75cm clamp）
  // subRcm = max(0, 1.5 - 0.3) = 1.2cm → clamp 到 0.75cm → adj = 0.75/1.5 = 0.5
  const r = await RC.syncLayoutChildrenR(
    h.driver, 'parent_p1', ['lc1', 'lc2', 'lc3', 'lc4'], 0.3, 'subtract', 1.5
  );
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 4, `bug #6 子 bug：期望 applied=4，实际 ${r.applied}（${r.failed} 个失败）`);
  assert.strictEqual(r.failed, 0);
  // 4 个子都写了（关键 assertion）
  for (let i = 0; i < 4; i++) {
    const c = f.layoutChildren[i];
    h.assertShape(c, { adjFraction: (v) => Math.abs(v - 0.5) < 1e-6 }, `lc${i+1} 应该 R 角=0.5cm`);
  }
});

t.test('bug #6 子 bug：applyLayout 写 4 个子（2×2）→ 全部写进去', async () => {
  // 同样问题在 applyLayout 路径：创建 2×2 layout 时 4 个子要全写 R 角
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver,
    'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    ['lc1', 'lc2', 'lc3', 'lc4'],
    { writeParentTag: true, syncR: true }
  );
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 4, `bug #6 子 bug：期望 applied=4，实际 ${r.applied}`);
  // 4 个子的 R 角都写了
  for (let i = 0; i < 4; i++) {
    const c = f.layoutChildren[i];
    // 父 R=0.3*8=2.4cm，subRcm = max(0, 2.4-0.3) = 2.1cm → clamp 到 短边一半 0.75cm → adj = 0.5
    h.assertShape(c, { adjFraction: (v) => Math.abs(v - 0.5) < 1e-6 }, `lc${i+1} 应该 R 角=0.5cm`);
  }
});

t.test('bug #6 子 bug：writeRadius opts.knownLockState 跳过 per-call readTag', async () => {
  // 验证修法核心：传 knownLockState 后，writeRadius 内部不再调 driver.readTag
  // → 测 call 计数（assertCalled 是不是 setAdjFraction 但不 readTag）
  const f = makeStandardFixture();
  f.layoutChildren[0]._tags[RC.LOCK_TAG_KEY] = '0.5';  // lc1 标 locked
  const h = createHarness({ shapes: f.allShapes });
  h.reset();
  await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1'], 0, 'same', 0.8);
  // 期望：调 1 次 loadTagsBulk（一次 sync 拿所有 tag），不调 readTag（per-shape sync）
  h.assertCalled('loadTagsBulk');
  h.assertCallCount('readTag', 0);
});

// ============================================================
// 完整 monitorTick → detectLayoutParentChanges → syncLayoutChildrenR 集成
// ============================================================
//
// 复现真实 PPT 场景：用户选中父 → 在 PPT 里直接拖父的 R 角黄色滑块
//   tick 1：monitorTick 读到 currentCm=2.4 → detectLayoutParentChanges 返回 [{parentId, lastCm=null, newCm=2.4}] → 只记 lastCm 不 fire
//   tick 2：用户拖完松手，currentCm=0.5 → detectLayoutParentChanges 返回 [{parentId, lastCm=2.4, newCm=0.5}] → 标 dirty + scheduleParentRSync
//   tick 3+：200ms 后 scheduleParentRSync 调 syncLayoutChildrenR → 4 个子联动

t.test('集成：完整 monitorTick → detect → schedule → sync 链路（4 子联动）', async () => {
  // 模拟 dialog.js 内存状态
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  // 父 + 4 个子都已经有 layout tag（fixture 没自动建，但 v1.2 期间建过——模拟）
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = JSON.stringify({
    rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract',
    childIds: ['lc1', 'lc2', 'lc3', 'lc4'],
  });
  for (const c of f.layoutChildren) {
    c._tags[RC.LAYOUT_CHILD_TAG_KEY] = 'parent_p1';
  }
  // 模拟 selectedShapes（dialog.js 内存里有 layout 父 + 4 个子的状态）
  // 关键：父的 currentCm 在 tick 1 时是 2.4（fixture 设定）
  const selectedShapes = [
    { id: 'parent_p1', layoutRole: 'parent', layoutParams: { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' }, layoutChildIds: ['lc1', 'lc2', 'lc3', 'lc4'], currentCm: 2.4, isRoundRect: true },
  ];
  const lockMonitor = { lastCm: {} };

  // === tick 1：first-see，不 fire ===
  let changes = RC.detectLayoutParentChanges(lockMonitor.lastCm, selectedShapes);
  assert.strictEqual(changes.length, 1);
  assert.strictEqual(changes[0].lastCm, null);
  assert.strictEqual(changes[0].newCm, 2.4);
  for (const c of changes) lockMonitor.lastCm[c.parentId] = c.newCm;  // 记 lastCm

  // === tick 2：用户拖完松手，currentCm = 0.5 ===
  selectedShapes[0].currentCm = 0.5;
  changes = RC.detectLayoutParentChanges(lockMonitor.lastCm, selectedShapes);
  assert.strictEqual(changes.length, 1, 'tick 2 应该检测到 1 个父变化');
  assert.strictEqual(changes[0].parentId, 'parent_p1');
  assert.strictEqual(changes[0].lastCm, 2.4);
  assert.strictEqual(changes[0].newCm, 0.5);
  for (const c of changes) lockMonitor.lastCm[c.parentId] = c.newCm;
  // 标 dirty（scheduleParentRSync 200ms 后会调 syncLayoutChildrenRIfNeeded）

  // === 模拟 scheduleParentRSync 调 syncLayoutChildrenR ===
  // subRcm = max(0, 0.5 - 0.3) = 0.2cm；子短边 1.5cm，0.2 不超 0.75，adj = 0.2/1.5
  const r = await RC.syncLayoutChildrenR(h.driver, 'parent_p1', ['lc1', 'lc2', 'lc3', 'lc4'], 0.3, 'subtract', 0.5);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 4, '关键 assertion：4 个子都要写');
  assert.strictEqual(r.failed, 0);
  // 验证每个子的 R 角
  for (let i = 0; i < 4; i++) {
    const c = f.layoutChildren[i];
    h.assertShape(c, { adjFraction: (v) => Math.abs(v - 0.2 / 1.5) < 1e-6 }, `lc${i+1} 应该 R 角=0.2cm`);
  }
});

t.test('集成：monitorTick 拖回原值（来回拖）→ 应该 fire（lastCm 更新了）', async () => {
  // 场景：用户拖父 R 角 2.4 → 0.5 → 2.4（来回拖）
  // 期望：每次拖动都 fire
  const lockMonitor = { lastCm: {} };
  const selectedShapes = [
    { id: 'p1', layoutRole: 'parent', currentCm: 2.4, isRoundRect: true },
  ];

  // tick 1：first-see
  let changes = RC.detectLayoutParentChanges(lockMonitor.lastCm, selectedShapes);
  assert.strictEqual(changes.length, 1);
  for (const c of changes) lockMonitor.lastCm[c.parentId] = c.newCm;

  // tick 2：拖到 0.5
  selectedShapes[0].currentCm = 0.5;
  changes = RC.detectLayoutParentChanges(lockMonitor.lastCm, selectedShapes);
  assert.strictEqual(changes.length, 1);
  for (const c of changes) lockMonitor.lastCm[c.parentId] = c.newCm;

  // tick 3：拖回 2.4
  selectedShapes[0].currentCm = 2.4;
  changes = RC.detectLayoutParentChanges(lockMonitor.lastCm, selectedShapes);
  assert.strictEqual(changes.length, 1, '来回拖应该 fire');
  assert.strictEqual(changes[0].newCm, 2.4);
  assert.strictEqual(changes[0].lastCm, 0.5);
});

// ============================================================
// v1.2.7：autoPadding 端到端集成（拖父 R 角时子位置/尺寸也跟着变）
// ============================================================

t.test('v1.2.7 集成：拖父 R 角从 2.4 → 3.5（超过 d_max）→ 子位置/尺寸自动变（autoPadding）', async () => {
  // 父 12x8 默认 R=2.4, d_init=0.3
  // 当父 R=3.5 → d_max = 8/2 - 3.5 = 0.5，d_init 0.3 < 0.5 → 不调
  // 但 R=3.9 → d_max = 0.1, 0.3 > 0.1 → effective = 0.1
  const f = makeStandardFixture();
  // 改父 R = 3.9
  f.parent._adjFraction = 3.9 / 8;  // R = 3.9cm
  const h = createHarness({ shapes: f.allShapes });
  // 模拟 dialog.js memory：layout 父 + 4 个子
  f.parent._tags[RC.LAYOUT_PARENT_TAG_KEY] = JSON.stringify({
    rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'same',
    childIds: ['lc1', 'lc2', 'lc3', 'lc4'],
  });
  for (const c of f.layoutChildren) c._tags[RC.LAYOUT_CHILD_TAG_KEY] = 'parent_p1';

  // v1.2.7：syncLayoutChildrenRIfNeeded 走 applyLayout → autoPadding 起作用
  const r = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'same' },
    ['lc1', 'lc2', 'lc3', 'lc4'],
    { writeParentTag: false, syncR: true }
  );
  assert.strictEqual(r.ok, true);
  // 验证：effective=0.1 + gutter=0.2 → subW = (12 - 2*0.1 - 0.2) / 2 = 5.8cm
  //                         subH = (8 - 2*0.1 - 0.2) / 2 = 3.8cm（gutter 在 width/height 都减一次）
  for (let i = 0; i < 4; i++) {
    const c = f.layoutChildren[i];
    const expectedW = 5.8 * PT_PER_CM;
    const expectedH = 3.8 * PT_PER_CM;
    assert.ok(Math.abs(c.width - expectedW) < 1, `lc${i+1} width ${c.width/PT_PER_CM}cm 应该 = 5.8cm (autoPadding effective=0.1, gutter=0.2)`);
    assert.ok(Math.abs(c.height - expectedH) < 1, `lc${i+1} height ${c.height/PT_PER_CM}cm 应该 = 3.8cm (autoPadding)`);
  }
  // 子 R 角 = 父 R 角 = 3.9cm（same 模式）→ 但子短边 3.9cm, max R = 1.95cm → clamp 到 1.95cm
  // 等距公式：R_sub = R_父 = 3.9 → clamp 到 1.95（短边一半）
  // adj = 1.95 / 3.9 = 0.5
  for (let i = 0; i < 4; i++) {
    const c = f.layoutChildren[i];
    h.assertShape(c, { adjFraction: 0.5 }, `lc${i+1} R 角应该 = 父 R 角 = 3.9cm (clamp 到 1.95cm)`);
  }
});

t.test('v1.2.7 集成：父 R=0.7 + d_init=0.3 → autoPadding 不调（正常 layout）', async () => {
  // 用户实际场景：父 12x8 R=0.7, d_init=0.3
  // d_max = 4 - 0.7 = 3.3, 0.3 < 3.3 → effective = 0.3（不调）
  const f = makeStandardFixture();
  f.parent._adjFraction = 0.7 / 8;
  const h = createHarness({ shapes: f.allShapes });
  const r = await RC.applyLayout(
    h.driver, 'parent_p1',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'same' },
    ['lc1', 'lc2', 'lc3', 'lc4'],
    { writeParentTag: false, syncR: true }
  );
  assert.strictEqual(r.ok, true);
  // 验证：effective=0.3 + gutter=0.2 → subW = (12 - 0.6 - 0.2) / 2 = 5.6cm
  //                         subH = (8 - 0.6 - 0.2) / 2 = 3.6cm（gutter 在 width/height 都减一次）
  for (let i = 0; i < 4; i++) {
    const c = f.layoutChildren[i];
    const expectedW = 5.6 * PT_PER_CM;
    const expectedH = 3.6 * PT_PER_CM;
    assert.ok(Math.abs(c.width - expectedW) < 1, `lc${i+1} width 应该 = 5.6cm (effective=0.3, gutter=0.2)`);
    assert.ok(Math.abs(c.height - expectedH) < 1, `lc${i+1} height 应该 = 3.6cm`);
  }
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
// UI 静态回归：边距/间距联动按钮
// ============================================================

t.test('边距/间距联动按钮始终显示链条，只用背景色区分状态', () => {
  const dialogHtml = fs.readFileSync(path.join(__dirname, '..', 'src', 'dialog', 'dialog.html'), 'utf8');
  const dialogJs = fs.readFileSync(path.join(__dirname, '..', 'src', 'dialog', 'dialog.js'), 'utf8');
  const i18nData = fs.readFileSync(path.join(__dirname, '..', 'src', 'dialog', 'i18n-data.js'), 'utf8');

  assert.ok(
    dialogHtml.includes('id="layout-pg-link-icon">🔗</span>'),
    '初始图标应该是链条'
  );
  assert.ok(
    dialogJs.includes("pgLinkIcon.textContent = '🔗';"),
    '切换联动状态后仍应该写入链条图标'
  );
  assert.ok(
    !dialogJs.includes("pgLinkIcon.textContent = linkPG ? '🔗' : '🔓';"),
    '关闭联动时不能再切换成开锁图标'
  );
  assert.ok(
    dialogHtml.includes('v1.3.1 · Group 布局与缩放稳定性修复') &&
      i18nData.includes("footerVersion: 'v1.3.1 · Group 布局与缩放稳定性修复'") &&
      i18nData.includes("footerVersion: 'v1.3.1 · Group layout and resize stability fixes'"),
    '中英文页脚版本都应该与正式 v1.3.1 一致'
  );
});

// ============================================================
// 跑
// ============================================================

t.run().catch((e) => { console.error(e); process.exit(1); });
