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
// 6. detectLayoutParentChanges — monitorTick → 联动 trigger 检测
// ============================================================

t.test('detectLayoutParentChanges: layout 父 R 角变了 → 返回 changed 列表', () => {
  // 模拟：monitorTick 看到 layout 父 currentCm 从 1.0 变成 1.5
  const knownCm = { p1: 1.0 };
  const sel = [{ id: 'p1', layoutRole: 'parent', currentCm: 1.5 }];
  const changes = RC.detectLayoutParentChanges(knownCm, sel);
  assert.strictEqual(changes.length, 1);
  assert.deepStrictEqual(changes[0], { parentId: 'p1', lastCm: 1.0, newCm: 1.5 });
});

t.test('detectLayoutParentChanges: 父 R 角没变（容差内） → 返回空', () => {
  // monitorTick 浮点抖动：当前 1.0001 vs 上次 1.0（容差 1e-3）→ 不算变
  const knownCm = { p1: 1.0 };
  const sel = [{ id: 'p1', layoutRole: 'parent', currentCm: 1.0001 }];
  const changes = RC.detectLayoutParentChanges(knownCm, sel);
  assert.strictEqual(changes.length, 0);
});

t.test('detectLayoutParentChanges: 首次见到（lastCm null） → 算变了（caller 决定要不要同步）', () => {
  const knownCm = {};  // 之前没记过
  const sel = [{ id: 'p1', layoutRole: 'parent', currentCm: 1.0 }];
  const changes = RC.detectLayoutParentChanges(knownCm, sel);
  assert.strictEqual(changes.length, 1);
  assert.deepStrictEqual(changes[0], { parentId: 'p1', lastCm: null, newCm: 1.0 });
});

t.test('detectLayoutParentChanges: 父 currentCm 是 null → 跳过（还没读到 R 角）', () => {
  const knownCm = { p1: 1.0 };
  const sel = [{ id: 'p1', layoutRole: 'parent', currentCm: null }];
  const changes = RC.detectLayoutParentChanges(knownCm, sel);
  assert.strictEqual(changes.length, 0);
});

t.test('detectLayoutParentChanges: 非 layout 父（layoutRole != "parent"） → 跳过', () => {
  // 子或普通形状 R 角变了不算（联动只针对父 → 子）
  const knownCm = { c1: 0.5 };
  const sel = [
    { id: 'c1', layoutRole: 'child', currentCm: 0.6 },
    { id: 'r1', layoutRole: null, currentCm: 0.7 },
  ];
  const changes = RC.detectLayoutParentChanges(knownCm, sel);
  assert.strictEqual(changes.length, 0);
});

t.test('detectLayoutParentChanges: 多个 layout 父混合 → 只返回有变化的', () => {
  const knownCm = { p1: 1.0, p2: 2.0, p3: 3.0 };
  const sel = [
    { id: 'p1', layoutRole: 'parent', currentCm: 1.0 },  // 没变
    { id: 'p2', layoutRole: 'parent', currentCm: 2.5 },  // 变了
    { id: 'p3', layoutRole: 'parent', currentCm: 3.0 },  // 没变
  ];
  const changes = RC.detectLayoutParentChanges(knownCm, sel);
  assert.strictEqual(changes.length, 1);
  assert.strictEqual(changes[0].parentId, 'p2');
  assert.deepStrictEqual(changes[0], { parentId: 'p2', lastCm: 2.0, newCm: 2.5 });
});

t.test('detectLayoutParentChanges: NaN / Infinity → 跳过（v1.3.5 invalid-adj 防御）', () => {
  const knownCm = { p1: 1.0, p2: 1.0 };
  const sel = [
    { id: 'p1', layoutRole: 'parent', currentCm: NaN },
    { id: 'p2', layoutRole: 'parent', currentCm: Infinity },
  ];
  const changes = RC.detectLayoutParentChanges(knownCm, sel);
  assert.strictEqual(changes.length, 0);
});

t.test('detectLayoutParentChanges: 空选区 → 返回空（不 throw）', () => {
  const changes = RC.detectLayoutParentChanges({}, []);
  assert.deepStrictEqual(changes, []);
  // 非数组（防呆）
  const changes2 = RC.detectLayoutParentChanges({}, null);
  assert.deepStrictEqual(changes2, []);
});

t.test('detectLayoutParentChanges: bug #6 场景 — 拖父 R 角触发联动', () => {
  // 复现 #6 bug：用户在 PPT 里直接拖父的 R 角黄色滑块
  // 1. 父初始 currentCm = 1.0（首次见到，caller 决定不主动同步）
  // 2. 用户拖到 1.5
  // 3. monitorTick 拿到 detectLayoutParentChanges 结果 → 触发 syncLayoutChildrenRIfNeeded
  //    → 子 R 角从 0.5 变成 max(0, 1.5 - padding) = 1.0
  const knownCm = {};
  // step 1: 首次见到
  const sel1 = [{ id: 'p1', layoutRole: 'parent', currentCm: 1.0 }];
  assert.strictEqual(RC.detectLayoutParentChanges(knownCm, sel1).length, 1);
  // step 2: caller 把 1.0 记到 knownCm
  knownCm.p1 = 1.0;
  // step 3: 用户拖到 1.5
  const sel2 = [{ id: 'p1', layoutRole: 'parent', currentCm: 1.5 }];
  const changes = RC.detectLayoutParentChanges(knownCm, sel2);
  assert.strictEqual(changes.length, 1);
  assert.strictEqual(changes[0].newCm, 1.5);
  // 验证联动公式：父 R 1.5, padding 0.5, subtract → subR = 1.0
  assert.strictEqual(RC.computeLinkedSubR(1.5, 0.5, 'subtract'), 1.0);
});

// ============================================================
// 7. parseLayoutParentTagValue / detectStaleChildrenInLayout — 纯函数
// ============================================================

t.test('parseLayoutParentTagValue: 完整合法 JSON', () => {
  const tag = JSON.stringify({ rows: 2, cols: 3, padding: 0.5, gutter: 0.3, linkRMode: 'subtract', childIds: ['a', 'b', 'c'] });
  const r = RC.parseLayoutParentTagValue(tag);
  assert.deepStrictEqual(r, {
    rows: 2, cols: 3, padding: 0.5, gutter: 0.3, linkRMode: 'subtract',
    childIds: ['a', 'b', 'c'],
  });
});

t.test('parseLayoutParentTagValue: 缺字段 → null（不 throw）', () => {
  assert.strictEqual(RC.parseLayoutParentTagValue(null), null);
  assert.strictEqual(RC.parseLayoutParentTagValue(''), null);
  assert.strictEqual(RC.parseLayoutParentTagValue('garbage'), null);
  assert.strictEqual(RC.parseLayoutParentTagValue('{"rows":2}'), null);  // 缺 cols + childIds
  assert.strictEqual(RC.parseLayoutParentTagValue('{"rows":2,"cols":2,"childIds":"not array"}'), null);
});

t.test('parseLayoutParentTagValue: 旧版 linkR 兼容（v1.2 前 boolean）', () => {
  // 旧版可能存的是 linkR: false → 转 'off'；linkR: true/缺省 → 'subtract'
  const r1 = RC.parseLayoutParentTagValue(JSON.stringify({ rows: 2, cols: 2, padding: 0.5, gutter: 0.3, linkR: false, childIds: [] }));
  assert.strictEqual(r1.linkRMode, 'off');
  const r2 = RC.parseLayoutParentTagValue(JSON.stringify({ rows: 2, cols: 2, padding: 0.5, gutter: 0.3, childIds: [] }));
  assert.strictEqual(r2.linkRMode, 'subtract');
});

t.test('parseLayoutParentTagValue: childIds 过滤掉非 string', () => {
  // 旧数据可能混着 number / null
  const r = RC.parseLayoutParentTagValue(JSON.stringify({ rows: 2, cols: 2, childIds: ['a', 1, null, 'b', ''] }));
  assert.deepStrictEqual(r.childIds, ['a', 'b']);
});

t.test('parseLayoutParentTagValue: padding/gutter 缺省 / NaN → 0', () => {
  const r = RC.parseLayoutParentTagValue(JSON.stringify({ rows: 2, cols: 2, padding: NaN, childIds: [] }));
  assert.strictEqual(r.padding, 0);
  const r2 = RC.parseLayoutParentTagValue(JSON.stringify({ rows: 2, cols: 2, gutter: 'bad', childIds: [] }));
  assert.strictEqual(r2.gutter, 0);
});

t.test('detectStaleChildrenInLayout: 全 valid → validChildIds 全，staleChildIds 空', () => {
  const parsed = { rows: 2, cols: 2, childIds: ['a', 'b', 'c', 'd'] };
  const ids = new Set(['a', 'b', 'c', 'd', 'e']);
  const r = RC.detectStaleChildrenInLayout(parsed, ids);
  assert.deepStrictEqual(r.validChildIds, ['a', 'b', 'c', 'd']);
  assert.deepStrictEqual(r.staleChildIds, []);
});

t.test('detectStaleChildrenInLayout: 部分 stale（中间页删除 / 跨 slide）→ 分类', () => {
  // 'a' 和 'c' 已被删（不在选区），'b' 和 'd' 还在
  const parsed = { rows: 2, cols: 2, childIds: ['a', 'b', 'c', 'd'] };
  const ids = new Set(['b', 'd', 'e']);
  const r = RC.detectStaleChildrenInLayout(parsed, ids);
  assert.deepStrictEqual(r.validChildIds, ['b', 'd']);
  assert.deepStrictEqual(r.staleChildIds, ['a', 'c']);
});

t.test('detectStaleChildrenInLayout: null / undefined → 空数组（不 throw）', () => {
  assert.deepStrictEqual(RC.detectStaleChildrenInLayout(null, new Set()), { validChildIds: [], staleChildIds: [] });
  assert.deepStrictEqual(RC.detectStaleChildrenInLayout(undefined, new Set()), { validChildIds: [], staleChildIds: [] });
  assert.deepStrictEqual(RC.detectStaleChildrenInLayout({}, new Set()), { validChildIds: [], staleChildIds: [] });
});

// ============================================================
// 8. pushHistory — 纯函数
// ============================================================

t.test('pushHistory 空 history → 推 1 条', () => {
  const r = RC.pushHistory([], 0.5, 'cm');
  assert.strictEqual(r.length, 1);
  assert.strictEqual(r[0].value, 0.5);
  assert.strictEqual(r[0].unit, 'cm');
  assert.ok(Number.isFinite(r[0].ts));
});

t.test('pushHistory 重复 value+unit → 去重 + 移到最前（保留新 ts）', () => {
  // v1.0 行为：unshift 新对象（带新 ts）→ 旧的同 value 记录被丢弃，新记录排到最前
  const h0 = [
    { value: 0.5, unit: 'cm', ts: 1 },
    { value: 0.3, unit: 'cm', ts: 2 },
    { value: 0.1, unit: 'cm', ts: 3 },
  ];
  const r = RC.pushHistory(h0, 0.5, 'cm');
  assert.strictEqual(r.length, 3);
  assert.strictEqual(r[0].value, 0.5);
  assert.ok(Number.isFinite(r[0].ts) && r[0].ts >= 1);  // 新 ts（≥ 原 ts）
  // 其他顺序保持
  assert.strictEqual(r[1].value, 0.3);
  assert.strictEqual(r[2].value, 0.1);
});

t.test('pushHistory 跨 unit 重复不算重复（cm 0.5 和 % 0.5 是不同的）', () => {
  const h0 = [{ value: 0.5, unit: 'cm', ts: 1 }];
  const r = RC.pushHistory(h0, 0.5, '%');
  assert.strictEqual(r.length, 2);
  assert.strictEqual(r[0].unit, '%');
  assert.strictEqual(r[1].unit, 'cm');
});

t.test('pushHistory 限 5 条上限', () => {
  // v1.0 行为：unshift 0.6 后变 6 条，slice(0, 5) 保留前 5 条 = [0.6, 0.1, 0.2, 0.3, 0.4]
  //         → 0.5（原来最新）被挤掉（不是 0.1）
  const h0 = [
    { value: 0.1, unit: 'cm', ts: 1 },
    { value: 0.2, unit: 'cm', ts: 2 },
    { value: 0.3, unit: 'cm', ts: 3 },
    { value: 0.4, unit: 'cm', ts: 4 },
    { value: 0.5, unit: 'cm', ts: 5 },
  ];
  const r = RC.pushHistory(h0, 0.6, 'cm');
  assert.strictEqual(r.length, 5);
  assert.strictEqual(r[0].value, 0.6);
  // v1.0 行为：0.5（次新）被挤掉
  assert.ok(r.every((h) => h.value !== 0.5));
});

t.test('pushHistory 不可变（原 history 数组不被改）', () => {
  const h0 = [{ value: 0.5, unit: 'cm', ts: 1 }];
  const h0Copy = JSON.parse(JSON.stringify(h0));
  RC.pushHistory(h0, 0.3, 'cm');
  assert.deepStrictEqual(h0, h0Copy);
});

t.test('pushHistory null/undefined history → 当作空数组处理', () => {
  const r1 = RC.pushHistory(null, 0.5, 'cm');
  assert.strictEqual(r1.length, 1);
  const r2 = RC.pushHistory(undefined, 0.3, 'cm');
  assert.strictEqual(r2.length, 1);
});

// ============================================================
// 跑
// ============================================================

t.run().catch((e) => { console.error(e); process.exit(1); });
