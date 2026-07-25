/*
 * test-radius-core.js — R 角调整 v1.3+ 纯算法测试
 *
 * 跑：node test/test-radius-core.js
 * 跑（npm）：npm test
 *
 * 测试目标（纯函数，无 driver / Office.js 依赖）：
 *   1. 布局 math（computeLayout）— 各种 grid + 边界
 *   2. 单位换算（valueToCm / cmToValue）— cm ↔ %
 *   3. R 角联动公式（computeLinkedSubR）— subtract / same / off
 *   4. clamp + adj 转换（clampRadius / cmToAdj / computeFinalRadius）
 *   5. 业务规则（shouldRejectWriteRadius / shouldRejectOnApply / shouldRejectLayoutApply / syncFixedValueIfLocked）
 *
 * 不测：
 *   - writeRadius / applyLayout / syncLayoutChildrenR — 走 driver 集成测试
 *   - writeRadiusToShapePure / applyLayoutPure — v1.3 已删（业务只走 driver 路径）
 *
 * 用新框架（fixtures + test-harness）:
 *   - 用 createTestRunner() 而不是手写 test() 框架
 *   - 共享 PT_PER_CM / cm helper
 */

const path = require('path');
const assert = require('assert');
const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
const { createTestRunner, PT_PER_CM } = require('./test-harness');

const t = createTestRunner();
const toPt = (c) => c * PT_PER_CM;  // cm → pt
const toCm = (pt) => pt / PT_PER_CM;  // pt → cm（纯算法测试用，fixtures.js 的 cm 方向相反）

// ============================================================
// 1. computeLayout — 布局 math
// ============================================================

t.test('computeLayout 2×2 基础：subW/subH 跟 4 个位置都对', () => {
  const r = RC.computeLayout({ left: 0, top: 0, width: toPt(12), height: toPt(8) }, 2, 2, 0.5, 0.3);
  assert.strictEqual(r.feasible, true);
  assert.strictEqual(r.positions.length, 4);
  assert.ok(Math.abs(toCm(r.subW) - 5.35) < 0.01);
  assert.ok(Math.abs(toCm(r.subH) - 3.35) < 0.01);
  // 4 个位置：(padding, padding) / (padding+subW+gutter, padding) / (padding, padding+subH+gutter) / 右下
  assert.ok(Math.abs(toCm(r.positions[0].left) - 0.5) < 0.01);
  assert.ok(Math.abs(toCm(r.positions[0].top) - 0.5) < 0.01);
  assert.ok(Math.abs(toCm(r.positions[1].left) - 0.5 - 5.35 - 0.3) < 0.01);
  assert.ok(Math.abs(toCm(r.positions[2].top) - 0.5 - 3.35 - 0.3) < 0.01);
});

t.test('computeLayout 1×3 横向：第 0 个在左，最后一个右边沿贴父右 padding', () => {
  const r = RC.computeLayout({ left: 0, top: 0, width: toPt(12), height: toPt(4) }, 1, 3, 0.3, 0.2);
  assert.strictEqual(r.positions.length, 3);
  assert.ok(Math.abs(toCm(r.subW) - (12 - 0.6 - 0.4) / 3) < 0.01);
  assert.ok(Math.abs(toCm(r.positions[0].left) - 0.3) < 0.01);
  // 最后位置：left + subW = 12 - padding
  const last = r.positions[2];
  assert.ok(Math.abs(toCm(last.left + last.w) - (12 - 0.3)) < 0.01);
});

t.test('computeLayout 3×2 非零起点：pos #0 起点 = parent.left+padding', () => {
  const r = RC.computeLayout({ left: toPt(5), top: toPt(3), width: toPt(10), height: toPt(6) }, 3, 2, 0.5, 0.3);
  assert.strictEqual(r.positions[0].idx, 0);
  assert.ok(Math.abs(toCm(r.positions[0].left) - 5.5) < 0.01);
  assert.ok(Math.abs(toCm(r.positions[0].top) - 3.5) < 0.01);
  assert.strictEqual(r.positions[5].idx, 5);
});

t.test('computeLayout 1×1：subW = 父 width - 2padding', () => {
  const r = RC.computeLayout({ left: 0, top: 0, width: toPt(10), height: toPt(6) }, 1, 1, 0.5, 0.3);
  assert.strictEqual(r.positions.length, 1);
  assert.ok(Math.abs(toCm(r.subW) - 9) < 0.01);
  assert.ok(Math.abs(toCm(r.subH) - 5) < 0.01);
});

t.test('computeLayout 不可行：padding 太大 → feasible=false + reason', () => {
  const r = RC.computeLayout({ left: 0, top: 0, width: toPt(2), height: toPt(2) }, 2, 2, 1, 0.1);
  assert.strictEqual(r.feasible, false);
  assert.strictEqual(r.positions.length, 0);
  assert.ok(r.reason.length > 0);
});

// ============================================================
// 2. valueToCm / cmToValue — 单位换算
// ============================================================

t.test('valueToCm cm 直通', () => assert.strictEqual(RC.valueToCm(1.5, 'cm', 0), 1.5));
t.test('valueToCm % 转 cm：50% of 4cm = 2cm', () => assert.ok(Math.abs(RC.valueToCm(50, '%', 4) - 2) < 0.01));
t.test('valueToCm % 转 cm：20% of 5cm = 1cm', () => assert.ok(Math.abs(RC.valueToCm(20, '%', 5) - 1) < 0.01));
t.test('cmToValue % 转回：2cm of 4cm = 50%', () => assert.ok(Math.abs(RC.cmToValue(2, '%', 4) - 50) < 0.01));
t.test('cmToValue refMinSide=0 防御 → 0', () => assert.strictEqual(RC.cmToValue(2, '%', 0), 0));
t.test('cmToValue cm 直通', () => assert.strictEqual(RC.cmToValue(1.5, 'cm', 0), 1.5));

// ============================================================
// 3. computeLinkedSubR — 联动公式
// ============================================================

t.test('computeLinkedSubR subtract：parentR=1.5, padding=0.3 → 1.2', () => {
  assert.ok(Math.abs(RC.computeLinkedSubR(1.5, 0.3, 'subtract') - 1.2) < 0.01);
});
t.test('computeLinkedSubR subtract：parentR < padding → 0（不能负）', () => {
  assert.strictEqual(RC.computeLinkedSubR(1.0, 1.5, 'subtract'), 0);
});
t.test('computeLinkedSubR subtract：parentR=0 → 0', () => {
  assert.strictEqual(RC.computeLinkedSubR(0, 0.3, 'subtract'), 0);
});
t.test('computeLinkedSubR same：parentR=1.5, padding=10 → 1.5（不 clamp）', () => {
  assert.strictEqual(RC.computeLinkedSubR(1.5, 10, 'same'), 1.5);
});
t.test('computeLinkedSubR off：任意 → 0', () => {
  assert.strictEqual(RC.computeLinkedSubR(1.5, 0.3, 'off'), 0);
  assert.strictEqual(RC.computeLinkedSubR(0, 0, 'off'), 0);
});

// ============================================================
// 4. clampRadius / cmToAdj / computeFinalRadius
// ============================================================

t.test('clampRadius 不超短边一半：target=1.5, minSide=4 → 1.5', () => {
  assert.strictEqual(RC.clampRadius(1.5, 4), 1.5);
});
t.test('clampRadius 超短边一半：target=3, minSide=4 → 2（=4/2）', () => {
  assert.strictEqual(RC.clampRadius(3, 4), 2);
});
t.test('clampRadius 负值：target=-0.5 → -0.5（不静默处理，让外面判断）', () => {
  assert.strictEqual(RC.clampRadius(-0.5, 4), -0.5);
});
t.test('clampRadius minSide=0：不 clamp（target 直通）', () => {
  assert.strictEqual(RC.clampRadius(1, 0), 1);
});

t.test('cmToAdj：R=1cm / minSide=4cm = 0.25', () => {
  assert.ok(Math.abs(RC.cmToAdj(1, 4) - 0.25) < 0.001);
});
t.test('cmToAdj 短边一半 = 0.5', () => {
  assert.ok(Math.abs(RC.cmToAdj(2, 4) - 0.5) < 0.001);
});
t.test('cmToAdj minSide=0 防御 → 0', () => {
  assert.strictEqual(RC.cmToAdj(1, 0), 0);
});

t.test('computeFinalRadius off → skipped=true, finalCm=0', () => {
  const r = RC.computeFinalRadius(4, 'off', 1.5, 0.5);
  assert.strictEqual(r.skipped, true);
  assert.strictEqual(r.finalCm, 0);
});
t.test('computeFinalRadius subtract：parentR=1.5, padding=0.5 → finalCm=1.0, adj=0.25', () => {
  const r = RC.computeFinalRadius(4, 'subtract', 1.5, 0.5);
  assert.strictEqual(r.skipped, false);
  assert.ok(Math.abs(r.finalCm - 1.0) < 0.01);
  assert.ok(Math.abs(r.adj - 0.25) < 0.001);
});
t.test('computeFinalRadius same：parentR=1.5 → finalCm=1.5, adj=0.375', () => {
  const r = RC.computeFinalRadius(4, 'same', 1.5, 0.5);
  assert.ok(Math.abs(r.finalCm - 1.5) < 0.01);
  assert.ok(Math.abs(r.adj - 0.375) < 0.001);
});
t.test('computeFinalRadius clamp：parentR=5, minSide=2 → finalCm=1（不能超 1）', () => {
  const r = RC.computeFinalRadius(2, 'subtract', 5, 0);
  assert.ok(Math.abs(r.finalCm - 1) < 0.01);
});
t.test('computeFinalRadius subtract parentR < padding：parentR=0.3, padding=1 → finalCm=0', () => {
  const r = RC.computeFinalRadius(4, 'subtract', 0.3, 1.0);
  assert.strictEqual(r.finalCm, 0);
});

// ============================================================
// 5. 业务规则（纯函数：被 dialog.js 第一道防线 + radius-core 内部用）
// ============================================================

t.test('shouldRejectWriteRadius 普通 → allow=true', () => {
  const r = RC.shouldRejectWriteRadius({ isStrict: false, isRoundRect: true, minSideCm: 4 });
  assert.strictEqual(r.allow, true);
});
t.test('shouldRejectWriteRadius strict → 拒绝 reason=strict（最高优先级）', () => {
  const r = RC.shouldRejectWriteRadius({ isStrict: true, isRoundRect: true, minSideCm: 4 });
  assert.strictEqual(r.allow, false);
  assert.strictEqual(r.reason, 'strict');
});
t.test('shouldRejectWriteRadius 非圆角 → reason=not-roundRect', () => {
  const r = RC.shouldRejectWriteRadius({ isStrict: false, isRoundRect: false, minSideCm: 4 });
  assert.strictEqual(r.reason, 'not-roundRect');
});
t.test('shouldRejectWriteRadius 0 尺寸 → reason=no-size', () => {
  const r = RC.shouldRejectWriteRadius({ isStrict: false, isRoundRect: true, minSideCm: 0 });
  assert.strictEqual(r.reason, 'no-size');
});

t.test('shouldRejectOnApply 空选区 → 不拒绝', () => {
  const r = RC.shouldRejectOnApply([]);
  assert.deepStrictEqual(r, { shouldReject: false, strictCount: 0 });
});
t.test('shouldRejectOnApply 5 个全普通 → 不拒绝', () => {
  const shapes = Array(5).fill({ isRoundRect: true, isStrict: false });
  const r = RC.shouldRejectOnApply(shapes);
  assert.deepStrictEqual(r, { shouldReject: false, strictCount: 0 });
});
t.test('shouldRejectOnApply 含 1 个 strict → 全部拒绝', () => {
  const r = RC.shouldRejectOnApply([
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: true },
    { isRoundRect: true, isStrict: false },
  ]);
  assert.strictEqual(r.shouldReject, true);
  assert.strictEqual(r.strictCount, 1);
});
t.test('shouldRejectOnApply 非 roundRect 的 strict 不算', () => {
  // {isRoundRect:false, isStrict:true} 不是圆角矩形，strict 标签没意义
  const r = RC.shouldRejectOnApply([
    { isRoundRect: false, isStrict: true },
    { isRoundRect: true, isStrict: false },
  ]);
  assert.deepStrictEqual(r, { shouldReject: false, strictCount: 0 });
});

t.test('shouldRejectLayoutApply 2×2 含 strict 子 → 整个拒绝', () => {
  const shapes = [
    { id: 'p', layoutRole: 'parent', isStrict: false },
    { id: 'c1', layoutRole: null, isStrict: false },
    { id: 'c2', layoutRole: null, isStrict: true },  // strict
    { id: 'c3', layoutRole: null, isStrict: false },
    { id: 'c4', layoutRole: null, isStrict: false },
  ];
  const r = RC.shouldRejectLayoutApply(shapes, 'p', ['c1', 'c2', 'c3', 'c4']);
  assert.strictEqual(r.shouldReject, true);
  assert.strictEqual(r.strictShapes.length, 1);
  assert.strictEqual(r.strictShapes[0].id, 'c2');
});
t.test('shouldRejectLayoutApply 父 strict 不影响（strict 标签只看子）', () => {
  const shapes = [
    { id: 'p', layoutRole: 'parent', isStrict: true },  // 父 strict
    { id: 'c1', layoutRole: null, isStrict: false },
    { id: 'c2', layoutRole: null, isStrict: false },
  ];
  const r = RC.shouldRejectLayoutApply(shapes, 'p', ['c1', 'c2']);
  assert.strictEqual(r.shouldReject, false);
});

t.test('syncFixedValueIfLocked unlocked → 不同步', () => {
  assert.deepStrictEqual(
    RC.syncFixedValueIfLocked({ isLocked: false, lockedCm: 0 }, 1.5),
    { newLockedCm: 0, synced: false }
  );
});
t.test('syncFixedValueIfLocked locked → 同步到 newCm', () => {
  assert.deepStrictEqual(
    RC.syncFixedValueIfLocked({ isLocked: true, lockedCm: 1.0 }, 1.5),
    { newLockedCm: 1.5, synced: true }
  );
});
t.test('syncFixedValueIfLocked locked 写 0 → 同步到 0（user 把 R 角拉到 0 也要同步）', () => {
  assert.deepStrictEqual(
    RC.syncFixedValueIfLocked({ isLocked: true, lockedCm: 1.0 }, 0),
    { newLockedCm: 0, synced: true }
  );
});

// ============================================================
// 集成场景（纯函数组合）
// ============================================================

t.test('集成：2×2 layout + subtract 联动公式 — 4 个子 subR 跟公式一致', () => {
  // 父 12×8cm, R=1.62cm, padding=0.5
  const parentBox = { left: 0, top: 0, width: toPt(12), height: toPt(8) };
  const parentRcm = 1.62;
  const padding = 0.5;

  const layout = RC.computeLayout(parentBox, 2, 2, padding, 0.3);
  assert.strictEqual(layout.feasible, true);

  for (let k = 0; k < 4; k++) {
    const pos = layout.positions[k];
    const childMinSideCm = Math.min(toCm(pos.w), toCm(pos.h));
    const final = RC.computeFinalRadius(childMinSideCm, 'subtract', parentRcm, padding);
    // 公式：subR = max(0, parentR - padding) = 1.12
    assert.ok(Math.abs(final.finalCm - 1.12) < 0.01, `子 #${k} subR=${final.finalCm} 应为 1.12`);
    // clamp 验证：不能超子短边一半
    assert.ok(final.finalCm <= childMinSideCm / 2 + 0.01);
  }
});

t.test('集成：1×3 layout + same 联动公式 — 3 个子 subR 都 = parentR', () => {
  const parentBox = { left: 0, top: 0, width: toPt(15), height: toPt(5) };
  const parentRcm = 1.0;
  const layout = RC.computeLayout(parentBox, 1, 3, 0.3, 0.2);

  for (let k = 0; k < 3; k++) {
    const pos = layout.positions[k];
    const childMinSideCm = Math.min(toCm(pos.w), toCm(pos.h));
    const final = RC.computeFinalRadius(childMinSideCm, 'same', parentRcm, 0.3);
    assert.strictEqual(final.finalCm, 1.0, `子 #${k} same 模式 subR=parentR=1.0`);
  }
});

t.test('集成：apply 拒绝场景 — strict 优先于其它', () => {
  // onApply 选区含 strict → 全部拒绝
  const r1 = RC.shouldRejectOnApply([
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: true },
  ]);
  assert.strictEqual(r1.shouldReject, true);
  // layout apply 含 strict 子 → 拒绝
  const r2 = RC.shouldRejectLayoutApply(
    [{ id: 'p', layoutRole: 'parent', isStrict: false }, { id: 'c1', isStrict: false }, { id: 'c2', isStrict: true }],
    'p', ['c1', 'c2']
  );
  assert.strictEqual(r2.shouldReject, true);
  // 全解锁 → 不拒绝
  const r3 = RC.shouldRejectOnApply([
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: false },
  ]);
  assert.strictEqual(r3.shouldReject, false);
});

t.test('集成：lock 同步 fixed value — 写 R 角后 lockedCm 跟 newCm 同步', () => {
  // 模拟 writeRadius 内部：写完 R 角 + 同步 fixed value
  const before = { isLocked: true, lockedCm: 0.5 };
  const newSubRcm = 1.12;
  const sync = RC.syncFixedValueIfLocked(before, newSubRcm);
  assert.strictEqual(sync.newLockedCm, 1.12);
  assert.strictEqual(sync.synced, true);

  const before2 = { isLocked: false, lockedCm: 0 };
  const sync2 = RC.syncFixedValueIfLocked(before2, newSubRcm);
  assert.strictEqual(sync2.synced, false);
});

t.test('集成：用户样例 — 父 R=1, padding=0.5, subtract → subR=0.5', () => {
  // v1.0 用户原始需求：外层 12×8cm R=1cm，内层缩进 0.5cm → 内层 R=0.5cm
  assert.strictEqual(RC.computeLinkedSubR(1.0, 0.5, 'subtract'), 0.5);
});

// ============================================================
// 跑
// ============================================================

t.run().catch((e) => { console.error(e); process.exit(1); });
