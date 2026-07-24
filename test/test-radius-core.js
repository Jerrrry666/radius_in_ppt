/*
 * test-radius-core.js — R 角调整 v1.2 核心算法测试
 *
 * 跑：node test/test-radius-core.js
 * 跑（npm）：npm test
 *
 * 测试目标：
 *   1. 布局 math（computeLayout）
 *   2. 单位换算（valueToCm / cmToValue）
 *   3. R 角联动公式（subtract / same / off）
 *   4. clamp + adj 转换
 *   5. strict/lock 行为规则
 *   6. onApply / layout apply 拒绝逻辑
 *   7. 端到端集成场景
 */

const core = require('../src/lib/radius-core.js');

let passed = 0;
let failed = 0;
let currentSuite = '';

function suite(name) {
  currentSuite = name;
  console.log('\n=== ' + name + ' ===');
}
function test(name, actual, expected, tol) {
  tol = tol || 1e-9;
  let ok;
  if (Array.isArray(expected)) {
    ok = Array.isArray(actual) && actual.length === expected.length;
    if (ok) {
      for (let i = 0; i < actual.length; i++) {
        if (typeof expected[i] === 'number') {
          if (Math.abs(actual[i] - expected[i]) > tol) { ok = false; break; }
        } else if (actual[i] !== expected[i]) { ok = false; break; }
      }
    }
  } else if (typeof expected === 'object' && expected !== null) {
    ok = actual !== null && typeof actual === 'object';
    if (ok) {
      for (const k in expected) {
        if (typeof expected[k] === 'number') {
          if (Math.abs(actual[k] - expected[k]) > tol) { ok = false; break; }
        } else if (actual[k] !== expected[k]) { ok = false; break; }
      }
    }
  } else if (typeof expected === 'number') {
    ok = typeof actual === 'number' && Math.abs(actual - expected) <= tol;
  } else {
    ok = actual === expected;
  }
  if (ok) {
    console.log('  ✓ ' + name);
    passed++;
  } else {
    console.log('  ✗ ' + name);
    console.log('    actual:   ' + JSON.stringify(actual));
    console.log('    expected: ' + JSON.stringify(expected));
    failed++;
  }
}
function approx(a, b, tol) { return Math.abs(a - b) <= (tol || 1e-4); }
function pt(cm) { return cm * core.PT_PER_CM; }
function cm(pt) { return pt / core.PT_PER_CM; }

// =====================================================
// 1. computeLayout（布局 math）
// =====================================================
suite('computeLayout — 基础 2×2');
{
  const parent = { left: 0, top: 0, width: pt(12), height: pt(8) };
  const r = core.computeLayout(parent, 2, 2, 0.5, 0.3);
  test('feasible', r.feasible, true);
  test('subW cm', cm(r.subW), 5.35, 0.01);  // (12 - 1 - 0.3) / 2 = 5.35
  test('subH cm', cm(r.subH), 3.35, 0.01);  // (8 - 1 - 0.3) / 2 = 3.35
  test('positions count', r.positions.length, 4);
  // 第 0 个位置：(padding, padding)
  test('pos #0 left cm', cm(r.positions[0].left), 0.5, 0.01);
  test('pos #0 top cm', cm(r.positions[0].top), 0.5, 0.01);
  // 第 1 个位置：(padding + subW + gutter, padding)
  test('pos #1 left cm', cm(r.positions[1].left), 0.5 + 5.35 + 0.3, 0.01);
  // 第 2 个位置：(padding, padding + subH + gutter)
  test('pos #2 top cm', cm(r.positions[2].top), 0.5 + 3.35 + 0.3, 0.01);
  // 第 3 个位置：右下角
  test('pos #3 left cm', cm(r.positions[3].left), 0.5 + 5.35 + 0.3, 0.01);
  test('pos #3 top cm', cm(r.positions[3].top), 0.5 + 3.35 + 0.3, 0.01);
}

suite('computeLayout — 1×3 横向');
{
  const parent = { left: 0, top: 0, width: pt(12), height: pt(4) };
  const r = core.computeLayout(parent, 1, 3, 0.3, 0.2);
  test('feasible', r.feasible, true);
  test('subW cm', cm(r.subW), (12 - 0.6 - 0.4) / 3, 0.01);  // (12 - 2*0.3 - 2*0.2) / 3
  test('subH cm', cm(r.subH), 4 - 0.6, 0.01);              // 4 - 2*0.3
  test('positions count', r.positions.length, 3);
  // 第 0 个：左
  test('pos #0 left cm', cm(r.positions[0].left), 0.3, 0.01);
  // 第 2 个：右（用容差避免浮点精度）
  test('pos #2 left cm', cm(r.positions[2].left), 12 - 0.3 - cm(r.subW), 0.05);
  // 验证：最后子形状的右边沿 = 父 left + 父 width - padding（右）
  // 即 lastPos.left + subW = 父 width - padding
  const lastPos = r.positions[r.positions.length - 1];
  test('最后 left + subW = 父 width - padding (右)', cm(lastPos.left + lastPos.w), 12 - 0.3, 0.01);
}

suite('computeLayout — 3×2 + 非零起点');
{
  const parent = { left: pt(5), top: pt(3), width: pt(10), height: pt(6) };
  const r = core.computeLayout(parent, 3, 2, 0.5, 0.3);
  test('feasible', r.feasible, true);
  test('subW cm', cm(r.subW), (10 - 1 - 0.3) / 2, 0.01);
  test('subH cm', cm(r.subH), (6 - 1 - 0.6) / 3, 0.01);
  // 第 0 个：(5 + 0.5, 3 + 0.5) cm
  test('pos #0 left cm', cm(r.positions[0].left), 5.5, 0.01);
  test('pos #0 top cm', cm(r.positions[0].top), 3.5, 0.01);
  // 第 5 个：右下角
  test('pos #5 idx', r.positions[5].idx, 5);
}

suite('computeLayout — 不可行（padding 太大）');
{
  const parent = { left: 0, top: 0, width: pt(2), height: pt(2) };
  const r = core.computeLayout(parent, 2, 2, 1, 0.1);
  test('feasible', r.feasible, false);
  test('positions empty', r.positions.length, 0);
  test('reason 非空', r.reason.length > 0, true);
}

suite('computeLayout — 单行单列 (1×1)');
{
  const parent = { left: 0, top: 0, width: pt(10), height: pt(6) };
  const r = core.computeLayout(parent, 1, 1, 0.5, 0.3);
  test('feasible', r.feasible, true);
  test('subW cm', cm(r.subW), 9, 0.01);  // 10 - 1
  test('subH cm', cm(r.subH), 5, 0.01);  // 6 - 1
  test('positions count', r.positions.length, 1);
  test('pos #0 (0.5, 0.5) cm', cm(r.positions[0].left), 0.5, 0.01);
}

// =====================================================
// 2. 单位换算
// =====================================================
suite('valueToCm / cmToValue');
{
  test('cm→cm 1.5', core.valueToCm(1.5, 'cm', 0), 1.5);
  test('cm→cm 0', core.valueToCm(0, 'cm', 0), 0);
  test('%→cm 50% of 4cm', core.valueToCm(50, '%', 4), 2, 0.01);
  test('%→cm 20% of 5cm', core.valueToCm(20, '%', 5), 1, 0.01);
  test('cm→% of 4cm', core.cmToValue(2, '%', 4), 50, 0.01);
  test('cm→% of 0', core.cmToValue(2, '%', 0), 0);  // 边界：refMinSide=0 → 0
  test('cm→cm (pass through)', core.cmToValue(1.5, 'cm', 0), 1.5);
}

// =====================================================
// 3. R 角联动公式
// =====================================================
suite('computeLinkedSubR — subtract (v1.0 公式)');
{
  test('parentR=1.5, padding=0.3 → 1.2', core.computeLinkedSubR(1.5, 0.3, 'subtract'), 1.2, 0.01);
  test('parentR=1.0, padding=1.5 → 0 (clamp ≥ 0)', core.computeLinkedSubR(1.0, 1.5, 'subtract'), 0, 0.01);
  test('parentR=0.5, padding=0.3 → 0.2', core.computeLinkedSubR(0.5, 0.3, 'subtract'), 0.2, 0.01);
  test('parentR=0 → 0', core.computeLinkedSubR(0, 0.3, 'subtract'), 0, 0.01);
  test('parentR=2, padding=0 → 2', core.computeLinkedSubR(2, 0, 'subtract'), 2, 0.01);
}
suite('computeLinkedSubR — same (严格 45° 等宽)');
{
  test('parentR=1.5 → 1.5 (无 padding 减)', core.computeLinkedSubR(1.5, 0.3, 'same'), 1.5, 0.01);
  test('parentR=0 → 0', core.computeLinkedSubR(0, 0.3, 'same'), 0, 0.01);
  test('parentR=1, padding=10 → 1 (不 clamp)', core.computeLinkedSubR(1, 10, 'same'), 1, 0.01);
}
suite('computeLinkedSubR — off');
{
  test('任意 → 0', core.computeLinkedSubR(1.5, 0.3, 'off'), 0);
  test('0/0 → 0', core.computeLinkedSubR(0, 0, 'off'), 0);
}

// =====================================================
// 4. clamp + adj 转换
// =====================================================
suite('clampRadius');
{
  test('target 1.5, minSide 4 → 1.5 (不超短边一半)', core.clampRadius(1.5, 4), 1.5);
  test('target 3.0, minSide 4 → 2.0 (clamp 到 短边一半)', core.clampRadius(3.0, 4), 2.0);
  test('target -0.5, minSide 4 → -0.5 (允许负值，让外面处理)', core.clampRadius(-0.5, 4), -0.5);
  test('target 1, minSide 0 → 1 (无 minSide 时不 clamp)', core.clampRadius(1, 0), 1);
  test('target 0 → 0', core.clampRadius(0, 4), 0);
}
suite('cmToAdj');
{
  test('R=1cm, minSide=4cm → adj=0.25', core.cmToAdj(1, 4), 0.25, 0.001);
  test('R=0, minSide=4 → adj=0', core.cmToAdj(0, 4), 0);
  test('R=2cm, minSide=4cm → adj=0.5 (短边一半)', core.cmToAdj(2, 4), 0.5, 0.001);
  test('R=1, minSide=0 → adj=0', core.cmToAdj(1, 0), 0);
}
suite('computeFinalRadius');
{
  // linkRMode = 'off' → skipped
  const r1 = core.computeFinalRadius(4, 'off', 1.5, 0.5);
  test('off → skipped=true', r1.skipped, true);
  test('off → finalCm=0', r1.finalCm, 0);

  // linkRMode = 'subtract'（默认公式）
  const r2 = core.computeFinalRadius(4, 'subtract', 1.5, 0.5);
  test('subtract → skipped=false', r2.skipped, false);
  test('subtract, parentR=1.5, padding=0.5 → finalCm=1.0', r2.finalCm, 1.0, 0.01);
  test('subtract → adj=0.25 (1/4)', r2.adj, 0.25, 0.001);

  // linkRMode = 'same'
  const r3 = core.computeFinalRadius(4, 'same', 1.5, 0.5);
  test('same, parentR=1.5 → finalCm=1.5', r3.finalCm, 1.5, 0.01);
  test('same → adj=0.375 (1.5/4)', r3.adj, 0.375, 0.001);

  // clamp：parentR 超子短边一半
  const r4 = core.computeFinalRadius(2, 'subtract', 5, 0);  // parentR=5, minSide=2 → 最多 1
  test('clamp: parentR=5, minSide=2 → finalCm=1', r4.finalCm, 1, 0.01);

  // subtract：parentR < padding → 0
  const r5 = core.computeFinalRadius(4, 'subtract', 0.3, 1.0);
  test('subtract, parentR=0.3 < padding=1.0 → finalCm=0', r5.finalCm, 0, 0.01);
}

// =====================================================
// 5. strict/lock 行为规则
// =====================================================
suite('shouldRejectWriteRadius');
{
  test('普通 roundRect, 非 strict → allow', core.shouldRejectWriteRadius({ isStrict: false, isRoundRect: true, minSideCm: 4 }), { allow: true });
  test('strict → reject', core.shouldRejectWriteRadius({ isStrict: true, isRoundRect: true, minSideCm: 4 }).allow, false);
  test('strict 拒绝 reason', core.shouldRejectWriteRadius({ isStrict: true, isRoundRect: true, minSideCm: 4 }).reason, 'strict');
  test('非 roundRect → reject (not-roundRect)', core.shouldRejectWriteRadius({ isStrict: false, isRoundRect: false, minSideCm: 4 }).reason, 'not-roundRect');
  test('minSide=0 → reject (no-size)', core.shouldRejectWriteRadius({ isStrict: false, isRoundRect: true, minSideCm: 0 }).reason, 'no-size');
}
suite('shouldRejectOnApply');
{
  test('空选区 → 不拒绝', core.shouldRejectOnApply([]), { shouldReject: false, strictCount: 0 });
  test('5 个普通 → 不拒绝', core.shouldRejectOnApply([
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: false },
  ]), { shouldReject: false, strictCount: 0 });
  test('1 个 strict → 全部拒绝', core.shouldRejectOnApply([
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: true },
    { isRoundRect: true, isStrict: false },
  ]), { shouldReject: true, strictCount: 1 });
  test('非 roundRect 严格不算', core.shouldRejectOnApply([
    { isRoundRect: false, isStrict: true },  // 不是 roundRect，不算 strict
    { isRoundRect: true, isStrict: false },
  ]), { shouldReject: false, strictCount: 0 });
}
suite('shouldRejectLayoutApply');
{
  const shapes = [
    { id: 'p', layoutRole: 'parent', isStrict: false },
    { id: 'c1', layoutRole: null, isStrict: false },
    { id: 'c2', layoutRole: null, isStrict: true },  // strict
    { id: 'c3', layoutRole: null, isStrict: false },
    { id: 'c4', layoutRole: null, isStrict: false },
  ];
  const r = core.shouldRejectLayoutApply(shapes, 'p', ['c1', 'c2', 'c3', 'c4']);
  test('2×2 含 strict 子 → 拒绝', r.shouldReject, true);
  test('strict 子数', r.strictShapes.length, 1);
  test('strict 子是 c2', r.strictShapes[0].id, 'c2');

  // 父是 strict 不影响
  const shapes2 = [
    { id: 'p', layoutRole: 'parent', isStrict: true },
    { id: 'c1', layoutRole: null, isStrict: false },
    { id: 'c2', layoutRole: null, isStrict: false },
  ];
  const r2 = core.shouldRejectLayoutApply(shapes2, 'p', ['c1', 'c2']);
  test('父 strict 不影响', r2.shouldReject, false);
}
suite('syncFixedValueIfLocked');
{
  test('unlocked → 不同步', core.syncFixedValueIfLocked({ isLocked: false, lockedCm: 0 }, 1.5), { newLockedCm: 0, synced: false });
  test('locked, 写 1.5 → 同步 1.5', core.syncFixedValueIfLocked({ isLocked: true, lockedCm: 1.0 }, 1.5), { newLockedCm: 1.5, synced: true });
  test('locked, 写 0 → 同步 0', core.syncFixedValueIfLocked({ isLocked: true, lockedCm: 1.0 }, 0), { newLockedCm: 0, synced: true });
}

// =====================================================
// 6. 集成场景
// =====================================================
suite('集成场景 — 1 大 + 4 小 (2×2) layout apply');
{
  // 父 12×8cm R=1.62cm，padding 0.5，gutter 0.3
  const parentBox = { left: 0, top: 0, width: pt(12), height: pt(8) };
  const parentRcm = 1.62;
  const padding = 0.5;
  const gutter = 0.3;
  const rows = 2, cols = 2;

  const layout = core.computeLayout(parentBox, rows, cols, padding, gutter);
  test('feasible', layout.feasible, true);

  // 4 个子
  for (let k = 0; k < 4; k++) {
    const pos = layout.positions[k];
    const subWcm = cm(pos.w);
    const subHcm = cm(pos.h);
    const childMinSideCm = Math.min(subWcm, subHcm);
    const final = core.computeFinalRadius(childMinSideCm, 'subtract', parentRcm, padding);
    const expectedSubR = Math.max(0, parentRcm - padding);  // 1.12
    test(`子 #${k} subR=1.12cm (subtract 公式)`, final.finalCm, expectedSubR, 0.01);
    // 验证：子 R 不能超过子短边一半
    test(`子 #${k} subR 不超子短边一半`, final.finalCm <= childMinSideCm / 2 + 0.01, true);
  }
}

suite('集成场景 — 1 大 + 3 小 (1×3) layout apply');
{
  const parentBox = { left: 0, top: 0, width: pt(15), height: pt(5) };
  const parentRcm = 1.0;
  const padding = 0.3;
  const gutter = 0.2;
  const rows = 1, cols = 3;

  const layout = core.computeLayout(parentBox, rows, cols, padding, gutter);
  test('feasible', layout.feasible, true);
  test('positions count', layout.positions.length, 3);
  // 每个子
  for (let k = 0; k < 3; k++) {
    const pos = layout.positions[k];
    const subWcm = cm(pos.w);
    const subHcm = cm(pos.h);
    const childMinSideCm = Math.min(subWcm, subHcm);
    const final = core.computeFinalRadius(childMinSideCm, 'same', parentRcm, padding);
    test(`1×3 子 #${k} same 模式: subR=parentR=1.0`, final.finalCm, 1.0, 0.01);
  }
}

suite('集成场景 — apply 拒绝场景（防误触优先）');
{
  // 场景 1: onApply 有 strict
  const r1 = core.shouldRejectOnApply([
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: true },
  ]);
  test('onApply 选区含 strict → 全部拒绝', r1.shouldReject, true);

  // 场景 2: layout apply 有 strict 子
  const r2 = core.shouldRejectLayoutApply(
    [
      { id: 'p', layoutRole: 'parent', isStrict: false },
      { id: 'c1', isStrict: false },
      { id: 'c2', isStrict: true },
    ],
    'p',
    ['c1', 'c2']
  );
  test('layout 含 strict 子 → 拒绝整个 apply', r2.shouldReject, true);

  // 场景 3: 全部解锁 → 不拒绝
  const r3 = core.shouldRejectOnApply([
    { isRoundRect: true, isStrict: false },
    { isRoundRect: true, isStrict: false },
  ]);
  test('全解锁 → 不拒绝', r3.shouldReject, false);
}

suite('集成场景 — lock 同步 fixed value');
{
  // 场景：lock 形状被 R 角联动公式更新 → fixed value 同步
  const before = { id: 'c1', isRoundRect: true, isStrict: false, isLocked: true, lockedCm: 0.5 };
  // 联动公式：subR = 1.12cm（减去 0.5 padding）
  const newSubRcm = 1.12;
  // writeRadiusToShape 内部：写 R 角 + 同步 fixed value
  const sync = core.syncFixedValueIfLocked(before, newSubRcm);
  test('locked 形状：fixed value 同步到新 R 角', sync.newLockedCm, 1.12, 0.01);
  test('synced=true', sync.synced, true);

  // 场景：unlock 形状 → 不动 lockedCm
  const before2 = { id: 'c2', isRoundRect: true, isStrict: false, isLocked: false, lockedCm: 0 };
  const sync2 = core.syncFixedValueIfLocked(before2, newSubRcm);
  test('unlocked 形状：fixed value 不动', sync2.synced, false);
}

suite('集成场景 — 用户给的样例（2×2 嵌套等距缩进）');
{
  // 用户原始需求：外层 12×8cm R=1cm，内层缩进 0.5cm → 内层 R 应该是 0.5cm
  // 父 R = 1.0, padding = 0.5 → subR = max(0, 1.0 - 0.5) = 0.5
  const parentR = 1.0;
  const padding = 0.5;
  const subR = core.computeLinkedSubR(parentR, padding, 'subtract');
  test('用户样例：父 R=1, padding=0.5 → subR=0.5', subR, 0.5, 0.01);
}

// =====================================================
// 输出
// =====================================================
console.log('\n' + '='.repeat(50));
console.log('结果: ' + passed + ' passed, ' + failed + ' failed');
console.log('='.repeat(50));
process.exit(failed > 0 ? 1 : 0);
