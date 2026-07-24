async function __main() {
/*
 * test-mock-harness.js — 用 mock shape 测 writeRadiusToShape + applyLayout 完整流程
 *
 * 跑：node test/test-mock-harness.js
 *
 * 目的：
 *   - 模拟 PowerPoint.js 的 shape proxy（adjustments / tags / position）
 *   - 在 mock 环境里跑 writeRadiusToShape 完整逻辑（读 tag → 拦截 → 写 R 角 → 同步 fixed value）
 *   - 在 mock 环境里跑 applyLayoutToChildren 端到端（建布局）
 *   - 验证各种场景下 mock shape 的最终状态对不对
 *
 * 局限：
 *   - 不能测真实的 PowerPoint.js API 行为（tag 持久化、cross-page 等）
 *   - 那些只能靠真实 PPT 测试 + debug panel 日志
 */

const core = require('../src/lib/radius-core.js');

let passed = 0;
let failed = 0;

function suite(name) {
  console.log('\n=== ' + name + ' ===');
}
async function test(name, actual, expected, tol) {
  tol = tol || 1e-9;
  let ok;
  if (typeof expected === 'object' && expected !== null && !Array.isArray(expected)) {
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
  } else if (Array.isArray(expected)) {
    ok = Array.isArray(actual) && actual.length === expected.length;
    if (ok) {
      for (let i = 0; i < actual.length; i++) {
        if (typeof expected[i] === 'number') {
          if (Math.abs(actual[i] - expected[i]) > tol) { ok = false; break; }
        } else if (actual[i] !== expected[i]) { ok = false; break; }
      }
    }
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
const pt = (cm) => cm * core.PT_PER_CM;
const cm = (pt) => pt / core.PT_PER_CM;

// =====================================================
// Mock shape factory
// =====================================================

/**
 * 创建一个 mock shape，符合 writeRadiusToShapePure 协议
 * @param {Object} init - { width(cm), height(cm), adjValue(0~1), tags }
 * @returns {Object} mock shape
 */
function makeMockShape(init) {
  init = init || {};
  const w = init.width != null ? pt(init.width) : pt(4);
  const h = init.height != null ? pt(init.height) : pt(2);
  // adj 内部存 _value 字段
  const adj = {
    _value: init.adjValue || 0,
    get count() { return init.adjCount != null ? init.adjCount : 1; },
    get: function(idx) {
      return { get value() { return adj._value; } };
    },
    set: function(idx, v) { adj._value = v; },
  };
  return {
    id: init.id || 'shape_' + Math.random().toString(36).slice(2, 8),
    width: w,
    height: h,
    left: init.left != null ? init.left : 0,
    top: init.top != null ? init.top : 0,
    adjustments: adj,
    tags: init.tags ? { ...init.tags } : {},
  };
}

function makeMockSlide(shapes) {
  const map = {};
  for (const s of shapes) map[s.id] = s;
  return { shapes: map };
}

// =====================================================
// writeRadiusToShapePure 完整流程
// =====================================================
suite('writeRadiusToShapePure — 基础写入');
{
  const shape = makeMockShape({ width: 4, height: 2, adjValue: 0.1 });
  let r = await core.writeRadiusToShapePure(shape, 1.5, {});
  await test('ok=true', r.ok, true);
  // 4×2cm shape, minSide=2, max R = 1 → 1.5 被 clamp 到 1.0
  await test('newCm=1.0 (clamp 到 minSide/2)', r.newCm, 1.0, 0.01);
  await test('wasLocked=false', r.wasLocked, false);
  // adj = 1.0 / 2 = 0.5
  await test('adj = 0.5 (1/2)', shape.adjustments._value, 0.5, 0.001);
}

suite('writeRadiusToShapePure — 不超短边一半');
{
  // 4×4cm 圆角矩形，minSide=4，max R = 2
  const shape = makeMockShape({ width: 4, height: 4, adjValue: 0 });
  let r = await core.writeRadiusToShapePure(shape, 1.0, {});
  await test('newCm=1.0', r.newCm, 1.0, 0.01);
  await test('adj = 0.25 (1/4)', shape.adjustments._value, 0.25, 0.001);
}

suite('writeRadiusToShapePure — clamp 超短边一半');
{
  // 2cm 矩形，写 5cm → clamp 到 1
  const shape = makeMockShape({ width: 2, height: 2, adjValue: 0 });
  let r = await core.writeRadiusToShapePure(shape, 5, {});
  await test('newCm=1.0 (clamp 到 minSide/2)', r.newCm, 1.0, 0.01);
  await test('adj=0.5 (短边一半)', shape.adjustments._value, 0.5, 0.001);
}

suite('writeRadiusToShapePure — strict 拦截');
{
  const shape = makeMockShape({
    width: 4, height: 4, adjValue: 0,
    tags: { [core.LOCK_STRICT_TAG_KEY]: '1' },
  });
  let r = await core.writeRadiusToShapePure(shape, 1.5, {});
  await test('ok=false', r.ok, false);
  await test('reason=strict', r.reason, 'strict');
  await test('isStrict=true', r.isStrict, true);
  await test('adj 没变', shape.adjustments._value, 0);
  await test('lock tag 没写', shape.tags[core.LOCK_TAG_KEY], undefined);
}

suite('writeRadiusToShapePure — locked 同步 fixed value');
{
  const shape = makeMockShape({
    width: 4, height: 4, adjValue: 0,
    tags: { [core.LOCK_TAG_KEY]: '0.5' },  // 原 fixed value 0.5
  });
  let r = await core.writeRadiusToShapePure(shape, 1.5, {});
  await test('ok=true', r.ok, true);
  await test('newCm=1.5', r.newCm, 1.5, 0.01);
  await test('wasLocked=true', r.wasLocked, true);
  await test('lock tag 同步到 1.5', shape.tags[core.LOCK_TAG_KEY], '1.5');
  await test('strict tag 没动', shape.tags[core.LOCK_STRICT_TAG_KEY], undefined);
}

suite('writeRadiusToShapePure — locked + strict 同时存在');
{
  // 这种情况理论上 onApply 会提前拒绝，但 writeRadiusToShapePure 内部也会拦截
  const shape = makeMockShape({
    width: 4, height: 4, adjValue: 0,
    tags: { [core.LOCK_TAG_KEY]: '0.5', [core.LOCK_STRICT_TAG_KEY]: '1' },
  });
  let r = await core.writeRadiusToShapePure(shape, 1.5, {});
  await test('ok=false (strict 优先)', r.ok, false);
  await test('reason=strict', r.reason, 'strict');
  await test('lock tag 没变', shape.tags[core.LOCK_TAG_KEY], '0.5');
  await test('adj 没变', shape.adjustments._value, 0);
}

suite('writeRadiusToShapePure — layoutParentId 写子 tag');
{
  const shape = makeMockShape({ width: 4, height: 4, adjValue: 0 });
  let r = await core.writeRadiusToShapePure(shape, 1.0, { layoutParentId: 'parent_1' });
  await test('ok=true', r.ok, true);
  await test('子 tag 写入了', shape.tags[core.LAYOUT_CHILD_TAG_KEY], 'parent_1');
}

suite('writeRadiusToShapePure — not-roundRect (count=0)');
{
  const shape = makeMockShape({ width: 4, height: 4, adjValue: 0, adjCount: 0 });
  let r = await core.writeRadiusToShapePure(shape, 1.5, {});
  await test('ok=false', r.ok, false);
  await test('reason=not-roundRect', r.reason, 'not-roundRect');
}

suite('writeRadiusToShapePure — no-size (width=0)');
{
  const shape = makeMockShape({ width: 0, height: 0, adjValue: 0 });
  let r = await core.writeRadiusToShapePure(shape, 1.5, {});
  await test('ok=false', r.ok, false);
  await test('reason=no-size', r.reason, 'no-size');
}

// =====================================================
// applyLayoutPure 端到端
// =====================================================

suite('applyLayoutPure — 1 大 + 4 小 (2×2)');
{
  // 父 12×8cm R=1.62 (adj = 1.62/8 = 0.2025)
  const parent = makeMockShape({
    id: 'p', width: 12, height: 8, adjValue: 1.62 / 8, left: pt(5), top: pt(3),
  });
  // 4 个子 (2.5×1.5cm 初始)
  const c1 = makeMockShape({ id: 'c1', width: 2.5, height: 1.5, adjValue: 0.2 });
  const c2 = makeMockShape({ id: 'c2', width: 2.5, height: 1.5, adjValue: 0.2 });
  const c3 = makeMockShape({ id: 'c3', width: 2.5, height: 1.5, adjValue: 0.2 });
  const c4 = makeMockShape({ id: 'c4', width: 2.5, height: 1.5, adjValue: 0.2 });
  const slide = makeMockSlide([parent, c1, c2, c3, c4]);

  const r = await core.applyLayoutPure(slide, 'p',
    { rows: 2, cols: 2, padding: 0.5, gutter: 0.3, linkRMode: 'subtract' },
    ['c1', 'c2', 'c3', 'c4'],
    { writeParentTag: true, syncR: true }
  );
  await test('ok=true', r.ok, true);
  await test('applied=4', r.applied, 4);
  await test('failed=0', r.failed, 0);

  // 4 个子的位置/尺寸应该按公式
  // subW = (12 - 1 - 0.3) / 2 = 5.35
  // subH = (8 - 1 - 0.3) / 2 = 3.35
  await test('c1 subW cm', cm(c1.width), 5.35, 0.01);
  await test('c1 subH cm', cm(c1.height), 3.35, 0.01);
  await test('c1 left cm (5 + 0.5)', cm(c1.left), 5.5, 0.01);
  await test('c1 top cm (3 + 0.5)', cm(c1.top), 3.5, 0.01);

  // R 角公式：subtract, parentR=1.62, padding=0.5 → subR = 1.12
  // c1 minSide = min(5.35, 3.35) = 3.35，max R = 1.675 → 不 clamp
  // adj = 1.12 / 3.35
  const c1ExpectedAdj = 1.12 / 3.35;
  await test('c1 adj ≈ 0.334 (1.12/3.35)', c1.adjustments._value, c1ExpectedAdj, 0.001);
  await test('c2 adj 同 c1', c2.adjustments._value, c1ExpectedAdj, 0.001);

  // 子 tag 写入了
  await test('c1 child tag', c1.tags[core.LAYOUT_CHILD_TAG_KEY], 'p');
  await test('c4 child tag', c4.tags[core.LAYOUT_CHILD_TAG_KEY], 'p');

  // 父 tag 写入了
  await test('父 tag 存在', !!parent.tags[core.LAYOUT_PARENT_TAG_KEY], true);
  const parentTag = JSON.parse(parent.tags[core.LAYOUT_PARENT_TAG_KEY]);
  await test('父 tag rows=2', parentTag.rows, 2);
  await test('父 tag cols=2', parentTag.cols, 2);
  await test('父 tag linkRMode=subtract', parentTag.linkRMode, 'subtract');
  await test('父 tag childIds.length=4', parentTag.childIds.length, 4);
}

suite('applyLayoutPure — 1 大 + 3 小 (1×3) same 模式');
{
  const parent = makeMockShape({
    id: 'p', width: 15, height: 5, adjValue: 1.0 / 5,
  });
  const c1 = makeMockShape({ id: 'c1', width: 1, height: 1, adjValue: 0 });
  const c2 = makeMockShape({ id: 'c2', width: 1, height: 1, adjValue: 0 });
  const c3 = makeMockShape({ id: 'c3', width: 1, height: 1, adjValue: 0 });
  const slide = makeMockSlide([parent, c1, c2, c3]);

  const r = await core.applyLayoutPure(slide, 'p',
    { rows: 1, cols: 3, padding: 0.3, gutter: 0.2, linkRMode: 'same' },
    ['c1', 'c2', 'c3'],
    { writeParentTag: true, syncR: true }
  );
  await test('ok=true', r.ok, true);
  await test('applied=3', r.applied, 3);

  // 1×3, 父 15×5, padding 0.3, gutter 0.2
  // subW = (15 - 0.6 - 0.4) / 3 = 14/3 ≈ 4.667
  // subH = 5 - 0.6 = 4.4
  await test('c1 subW cm', cm(c1.width), 14 / 3, 0.01);
  await test('c1 subH cm', cm(c1.height), 4.4, 0.01);

  // same 模式：subR = parentR = 1.0
  // c1 minSide = min(4.667, 4.4) = 4.4，max R = 2.2 → 不 clamp
  // adj = 1.0 / 4.4 ≈ 0.227
  const c1ExpectedAdj = 1.0 / 4.4;
  await test('c1 adj ≈ 0.227 (1.0/4.4)', c1.adjustments._value, c1ExpectedAdj, 0.001);
}

suite('applyLayoutPure — 含 strict 子 → 整个 apply 拒绝');
{
  const parent = makeMockShape({ id: 'p', width: 10, height: 10, adjValue: 0.1 });
  const c1 = makeMockShape({ id: 'c1', width: 1, height: 1, adjValue: 0 });
  const c2 = makeMockShape({
    id: 'c2', width: 1, height: 1, adjValue: 0,
    tags: { [core.LOCK_STRICT_TAG_KEY]: '1' },  // strict
  });
  const slide = makeMockSlide([parent, c1, c2]);

  const r = await core.applyLayoutPure(slide, 'p',
    { rows: 1, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    ['c1', 'c2'],
    { writeParentTag: true, syncR: true }
  );
  await test('ok=false', r.ok, false);
  await test('strictCount=1', r.strictCount, 1);

  // 重要：c1 也不应该被改（位置/尺寸/R 角都没动）
  await test('c1 width 没变', cm(c1.width), 1, 0.01);
  await test('c1 height 没变', cm(c1.height), 1, 0.01);
  await test('c1 adj 没变', c1.adjustments._value, 0);
  await test('c2 width 没变', cm(c2.width), 1, 0.01);
  await test('c2 adj 没变', c2.adjustments._value, 0);
  // 父 tag 也没写
  await test('父 tag 没写', parent.tags[core.LAYOUT_PARENT_TAG_KEY], undefined);
}

suite('applyLayoutPure — locked 子同步 fixed value');
{
  const parent = makeMockShape({ id: 'p', width: 10, height: 10, adjValue: 0.1 });
  const c1 = makeMockShape({
    id: 'c1', width: 1, height: 1, adjValue: 0,
    tags: { [core.LOCK_TAG_KEY]: '0.3' },  // 原 fixed value 0.3
  });
  const c2 = makeMockShape({ id: 'c2', width: 1, height: 1, adjValue: 0 });
  const slide = makeMockSlide([parent, c1, c2]);

  const r = await core.applyLayoutPure(slide, 'p',
    { rows: 1, cols: 2, padding: 0.5, gutter: 0.3, linkRMode: 'subtract' },
    ['c1', 'c2'],
    { writeParentTag: true, syncR: true }
  );
  await test('ok=true', r.ok, true);
  await test('applied=2', r.applied, 2);
  await test('lockedCount=1', r.lockedCount, 1);

  // c1: subtract, parentR = 0.1 * 10 = 1.0, padding=0.5 → subR = 0.5
  // 但 c1 minSide = min(subW, subH) = min((10-1-0.3)/2, 10-1) = min(4.35, 9) = 4.35
  // 不 clamp，subR = 0.5
  // lock tag 同步到 0.5
  await test('c1 lock tag 同步到 0.5', c1.tags[core.LOCK_TAG_KEY], '0.5');
}

suite('applyLayoutPure — 子数量不足');
{
  const parent = makeMockShape({ id: 'p', width: 10, height: 10, adjValue: 0.1 });
  const c1 = makeMockShape({ id: 'c1', width: 1, height: 1, adjValue: 0 });
  const slide = makeMockSlide([parent, c1]);

  const r = await core.applyLayoutPure(slide, 'p',
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    ['c1'],  // 只 1 个，need 4
    { writeParentTag: true, syncR: true }
  );
  await test('ok=false', r.ok, false);
  await test('warn 非空', r.warn.length > 0, true);
}

suite('applyLayoutPure — 父不是 roundRect');
{
  // adjCount=0 (父不是 roundRect)
  const parent = makeMockShape({ id: 'p', width: 10, height: 10, adjCount: 0, adjValue: 0 });
  const c1 = makeMockShape({ id: 'c1', width: 1, height: 1, adjValue: 0 });
  const c2 = makeMockShape({ id: 'c2', width: 1, height: 1, adjValue: 0 });
  const slide = makeMockSlide([parent, c1, c2]);

  const r = await core.applyLayoutPure(slide, 'p',
    { rows: 1, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    ['c1', 'c2'],
    { writeParentTag: true, syncR: true }
  );
  // 父 R = 0
  await test('ok=true', r.ok, true);
  // subR = max(0, 0 - 0.3) = 0
  await test('c1 adj=0', c1.adjustments._value, 0);
}

suite('applyLayoutPure — 跨页隔离（父不在当前 slide）');
{
  const c1 = makeMockShape({ id: 'c1', width: 1, height: 1, adjValue: 0 });
  const slide = makeMockSlide([c1]);  // 没有 p

  const r = await core.applyLayoutPure(slide, 'p',
    { rows: 1, cols: 1, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    ['c1'],
    { writeParentTag: true, syncR: true }
  );
  await test('ok=false', r.ok, false);
  await test('warn 提到 parent', r.warn.indexOf('parent') >= 0, true);
}

suite('applyLayoutPure — linkRMode=off (不联动)');
{
  const parent = makeMockShape({ id: 'p', width: 10, height: 10, adjValue: 0.1 });
  const c1 = makeMockShape({ id: 'c1', width: 1, height: 1, adjValue: 0.3 });  // 原 adj 0.3
  const slide = makeMockSlide([parent, c1]);

  const r = await core.applyLayoutPure(slide, 'p',
    { rows: 1, cols: 1, padding: 0.3, gutter: 0.2, linkRMode: 'off' },
    ['c1'],
    { writeParentTag: true, syncR: true }
  );
  await test('ok=true', r.ok, true);
  // 位置/尺寸应该改
  await test('c1 subW 改了', cm(c1.width) > 1, true);
  // R 角没动（off 模式）
  await test('c1 adj 保持 0.3', c1.adjustments._value, 0.3, 0.001);
}

// =====================================================
// 输出
// =====================================================
console.log('\n' + '='.repeat(50));
console.log('结果: ' + passed + ' passed, ' + failed + ' failed');
console.log('='.repeat(50));
process.exit(failed > 0 ? 1 : 0);

}
__main().catch(e => { console.error(e); process.exit(1); });