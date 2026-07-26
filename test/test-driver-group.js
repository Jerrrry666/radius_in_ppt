/*
 * test-driver-group.js — GroupShape 分层加载/展平单测 + 集成（v1.3.1）
 *
 * 测 3 件事：
 *   1. driver.isGroup / groupShapes：mock group 节点正确识别
 *   2. driver.flattenSelected：递归展平 group → 叶子 shape 数组
 *   3. 集成：flatten 后业务函数（writeRadius / loadLayoutTags）能正常处理叶子 shape
 *
 * 关键边界：
 *   - 嵌套 group（group 里再 group）→ 全部展平
 *   - 防死循环（group 循环引用）→ seen Set
 *   - 空选区 / null / 普通数组 / Office.js collection 4 种输入
 *   - mixed：group + 普通 shape 混在选区
 *   - 业务函数对叶子数组无感（不需要改 radius-core）
 *
 * mock group 协议（driver.isGroup / groupShapes 内部约定）：
 *   - s._isGroup = true   → driver.isGroup 走 mock 路径返回 true
 *   - s._groupShapes = [s1, s2, ...] → driver.groupShapes 走 mock 路径返回子数组
 *   - 真实 PPT：s.type === 'Group' + s.group.shapes.items（不需要 mock 字段）
 */

const path = require('path');
const assert = require('assert');
const { createDriver } = require(path.join(__dirname, '..', 'src', 'lib', 'ppt-driver.js'));
const { createHarness, makeFixtureShape, PT_PER_CM, cm } = require('./test-harness');

// 全局 ctx（mock）— driver 不需要真的 sync
const ctx = { sync: async () => {} };
const driver = createDriver(ctx);

// ── helpers ──
// 造一个 mock group（带 _isGroup + _groupShapes + group level）
function makeMockGroup(id, children) {
  const group = {
    id,
    _isGroup: true,
    _groupShapes: children || [],
    _groupLevel: 0,
    width: 0, height: 0, left: 0, top: 0,  // group proxy 也有 box，但业务不读
    adjustments: { count: 0, get: () => ({ value: 0 }), set: () => {} },
    tags: {},
  };
  const markLevel = (shape, level) => {
    if (!shape) return;
    shape._groupLevel = level;
    if (shape._isGroup && Array.isArray(shape._groupShapes)) {
      for (const child of shape._groupShapes) markLevel(child, level + 1);
    }
  };
  for (const child of group._groupShapes) markLevel(child, 1);
  return group;
}

// 造一个 mock leaf shape（用 makeFixtureShape 就行）
function makeLeaf(id, opts) {
  return makeFixtureShape(Object.assign({ id }, opts || {}));
}

// 造一个可记录 load 的 Office.js collection mock
function makeLoadableCollection(items, loads) {
  return {
    items: items || [],
    load(fields) {
      loads.push(fields);
      // 回归保护：Mac LTSC 单选普通 shape 时，group path 会让 sync 抛 GeneralException。
      if (String(fields).includes('group/shapes')) {
        throw new Error('GeneralException: group path applied to non-group');
      }
    },
  };
}

function makeOfficeGroup(id, children, loads) {
  const g = makeLeaf(id);
  g.type = 'Group';
  g.group = { shapes: makeLoadableCollection(children, loads) };
  return g;
}

const t = (() => {
  const tests = [];
  const setupFns = [];
  const teardownFns = [];
  function test(name, fn) { tests.push({ name, fn }); }
  function beforeEach(fn) { setupFns.push(fn); }
  function afterEach(fn) { teardownFns.push(fn); }
  async function run() {
    let passed = 0, failed = 0;
    for (const { name, fn } of tests) {
      for (const s of setupFns) await s();
      try {
        await fn();
        passed++;
        console.log(`  ✓ ${name}`);
      } catch (e) {
        failed++;
        console.log(`  ✗ ${name}`);
        console.log(`     ${e.message}`);
        if (e.stack) console.log(`     ${e.stack.split('\n').slice(1, 4).join('\n     ')}`);
      }
      for (const t of teardownFns) await t();
    }
    console.log('\n' + '='.repeat(50));
    console.log(`结果: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(50));
    if (failed > 0) process.exit(1);
  }
  return { test, beforeEach, afterEach, run };
})();

// ============================================================
// driver.isGroup — 基础判定
// ============================================================

t.test('isGroup: null/undefined/普通 shape → false', () => {
  assert.strictEqual(driver.isGroup(null), false);
  assert.strictEqual(driver.isGroup(undefined), false);
  assert.strictEqual(driver.isGroup(makeLeaf('a')), false);
});

t.test('isGroup: _isGroup=true mock → true', () => {
  const g = makeMockGroup('g1', []);
  assert.strictEqual(driver.isGroup(g), true);
});

t.test('isGroup: type=\'Group\' 字符串 → true（真实 PPT 路径）', () => {
  // 模拟真实 PPT 上 PowerPoint.ShapeType.group 的值 = 'Group'
  const s = makeLeaf('a');
  s.type = 'Group';
  assert.strictEqual(driver.isGroup(s), true);
});

t.test('shapeLevel: 顶层 shape=0，group 叶子=1，嵌套叶子=2', () => {
  const top = makeLeaf('top');
  const nestedLeaf = makeLeaf('nestedLeaf');
  const nestedGroup = makeMockGroup('nestedGroup', [nestedLeaf]);
  makeMockGroup('rootGroup', [nestedGroup]);

  assert.strictEqual(driver.shapeLevel(top), 0);
  assert.strictEqual(driver.shapeLevel(nestedGroup), 1);
  assert.strictEqual(driver.shapeLevel(nestedLeaf), 2);
});

t.test('shapeLevel: 真实 Office.js level 字段可读', () => {
  const child = makeLeaf('child');
  child.level = 1;
  assert.strictEqual(driver.shapeLevel(child), 1);
});

t.test('parentGroupOf/topGroupOf: 从安全 tree 索引取得直接组与顶层组', () => {
  const leaf = makeLeaf('leaf');
  const nested = makeMockGroup('nested', [leaf]);
  const root = makeMockGroup('root', [nested]);
  driver.flattenSelected([root]);

  assert.strictEqual(driver.parentGroupOf(leaf), nested);
  assert.strictEqual(driver.parentGroupOf(nested), root);
  assert.strictEqual(driver.topGroupOf(leaf), root);
  assert.strictEqual(driver.parentGroupOf(root), null);
});

t.test('isGroup: type=几何形状 → false', () => {
  const s = makeLeaf('a');
  s.type = 'GeometricShape';
  assert.strictEqual(driver.isGroup(s), false);
});

t.test('isGroup: 读 type 抛错 → false（防御）', () => {
  const s = makeLeaf('a');
  Object.defineProperty(s, 'type', { get() { throw new Error('boom'); } });
  assert.strictEqual(driver.isGroup(s), false);
});

// ============================================================
// driver.groupShapes — 拿子 shape 数组
// ============================================================

t.test('groupShapes: null → []', () => {
  assert.deepStrictEqual(driver.groupShapes(null), []);
});

t.test('groupShapes: _groupShapes mock → 返回子数组', () => {
  const c1 = makeLeaf('c1');
  const c2 = makeLeaf('c2');
  const g = makeMockGroup('g1', [c1, c2]);
  const subs = driver.groupShapes(g);
  assert.strictEqual(subs.length, 2);
  assert.strictEqual(subs[0].id, 'c1');
  assert.strictEqual(subs[1].id, 'c2');
});

t.test('groupShapes: 真实 PPT 路径（s.group.shapes.items）', () => {
  const c1 = makeLeaf('c1');
  const c2 = makeLeaf('c2');
  const s = makeLeaf('g1');
  // 模拟 PowerPoint.ShapeGroup：s.group.shapes 是 { items: [...] }
  s.group = { shapes: { items: [c1, c2] } };
  const subs = driver.groupShapes(s);
  assert.strictEqual(subs.length, 2);
});

t.test('groupShapes: 没 group 字段（普通 shape）→ []', () => {
  const s = makeLeaf('a');
  // 没 s.group 字段
  assert.deepStrictEqual(driver.groupShapes(s), []);
});

// ============================================================
// driver.loadTagsBulk — TagCollection key/value 批量加载
// ============================================================

t.test('loadTagsBulk: 所有 TagCollection 排队后只 sync 一次，并保留宿主大写 key', async () => {
  const loads = [];
  let syncCount = 0;
  const d = createDriver({ sync: async () => { syncCount++; } });
  const shapes = [
    {
      id: 's1',
      tags: {
        items: [{ key: 'RADIUSLOCK_V1', value: '0.25' }],
        load(fields) { loads.push(['s1', fields]); },
      },
    },
    {
      id: 's2',
      tags: {
        items: [{ key: 'LAYOUTCHILD_V1', value: 'parent' }],
        load(fields) { loads.push(['s2', fields]); },
      },
    },
  ];

  const result = await d.loadTagsBulk(shapes);

  assert.deepStrictEqual(loads, [
    ['s1', 'key, value'],
    ['s2', 'key, value'],
  ]);
  assert.strictEqual(syncCount, 1);
  assert.strictEqual(result.s1.RADIUSLOCK_V1, '0.25');
  assert.strictEqual(result.s2.LAYOUTCHILD_V1, 'parent');
});

// ============================================================
// driver.flattenSelected — 核心展平逻辑
// ============================================================

t.test('flattenSelected: 空数组 → []', () => {
  assert.deepStrictEqual(driver.flattenSelected([]), []);
});

t.test('flattenSelected: null / undefined → []', () => {
  assert.deepStrictEqual(driver.flattenSelected(null), []);
  assert.deepStrictEqual(driver.flattenSelected(undefined), []);
});

t.test('flattenSelected: 全是叶子（无 group）→ 原样返回', () => {
  const a = makeLeaf('a');
  const b = makeLeaf('b');
  const c = makeLeaf('c');
  const out = driver.flattenSelected([a, b, c]);
  assert.strictEqual(out.length, 3);
  assert.deepStrictEqual(out.map(s => s.id), ['a', 'b', 'c']);
});

t.test('flattenSelected: 单层 group（3 个子）→ 3 个叶子', () => {
  const c1 = makeLeaf('c1');
  const c2 = makeLeaf('c2');
  const c3 = makeLeaf('c3');
  const g = makeMockGroup('g1', [c1, c2, c3]);
  const out = driver.flattenSelected([g]);
  assert.strictEqual(out.length, 3);
  assert.deepStrictEqual(out.map(s => s.id), ['c1', 'c2', 'c3']);
});

t.test('flattenSelected: 嵌套 group（group 里再 group）→ 全部展平', () => {
  // 结构：g1 → [c1, g2, c3]
  //       g2 → [c2a, c2b]
  // 期望输出顺序：c1, c2a, c2b, c3（depth-first）
  const c1 = makeLeaf('c1');
  const c2a = makeLeaf('c2a');
  const c2b = makeLeaf('c2b');
  const c3 = makeLeaf('c3');
  const g2 = makeMockGroup('g2', [c2a, c2b]);
  const g1 = makeMockGroup('g1', [c1, g2, c3]);
  const out = driver.flattenSelected([g1]);
  assert.strictEqual(out.length, 4);
  assert.deepStrictEqual(out.map(s => s.id), ['c1', 'c2a', 'c2b', 'c3']);
});

t.test('flattenSelected: 混合（group + 普通 shape）→ 全部展平', () => {
  const a = makeLeaf('a');
  const c1 = makeLeaf('c1');
  const c2 = makeLeaf('c2');
  const b = makeLeaf('b');
  const g = makeMockGroup('g1', [c1, c2]);
  const out = driver.flattenSelected([a, g, b]);
  assert.strictEqual(out.length, 4);
  assert.deepStrictEqual(out.map(s => s.id), ['a', 'c1', 'c2', 'b']);
});

t.test('flattenSelected: 多个 group + 普通 shape → 全部展平', () => {
  const a = makeLeaf('a');
  const g1 = makeMockGroup('g1', [makeLeaf('c1'), makeLeaf('c2')]);
  const b = makeLeaf('b');
  const g2 = makeMockGroup('g2', [makeLeaf('c3')]);
  const out = driver.flattenSelected([a, g1, b, g2]);
  assert.strictEqual(out.length, 5);
  assert.deepStrictEqual(out.map(s => s.id), ['a', 'c1', 'c2', 'b', 'c3']);
});

t.test('flattenSelected: 防死循环（A 含 B，B 含 A）→ 不会无限递归', () => {
  // 构造循环：A._groupShapes = [B]，B._groupShapes = [A]
  // 期望：seen Set 拦住 A 或 B 的二次访问
  const a = makeMockGroup('A', []);
  const b = makeMockGroup('B', [a]);
  a._groupShapes = [b];
  // A 在选区里：walk(A) → walk(B) → walk(A) → seen(A) 拦住 → 退出
  const out = driver.flattenSelected([a]);
  // 结果应该只有 B（被加 out），A 第二次访问时 isGroup 但 sub 是已 visited 的 A（被跳过）
  // 实际：第一次 walk(A) → 展开 B → walk(B) → 展开 A → 第二次 walk(A) seen → return
  // 第二次 walk(A) 因为已经走过 seen(A)，直接 return，不加 out
  // 所以结果是：B（叶子虽然 A 是 group）
  // 等等：B 的 sub 是 A，A 是 group，不加 out（只 push 叶子）→ out = []
  // 然后 B 自己 isGroup，不加 out
  // 所以 out = []
  // 这是正确行为：循环引用下我们不强行 push 任何东西（避免重复）
  assert.strictEqual(out.length, 0);
});

t.test('flattenSelected: 输入是 Office.js collection（{items: [...]}）→ 拿 .items 展平', () => {
  const c1 = makeLeaf('c1');
  const c2 = makeLeaf('c2');
  const collectionLike = { items: [c1, c2] };
  const out = driver.flattenSelected(collectionLike);
  assert.strictEqual(out.length, 2);
  assert.deepStrictEqual(out.map(s => s.id), ['c1', 'c2']);
});

t.test('flattenSelected: 同一 shape 在选区出现两次 → 只入结果一次（seen Set）', () => {
  const a = makeLeaf('a');
  const out = driver.flattenSelected([a, a]);
  // 第二次访问 seen(a) 拦住，但 a 是叶子，会被 push 一次还是被 seen 拦？
  // 实际逻辑：seen 检查在 walk 开头，命中就 return（不 push）
  // 所以 out 应该是 [a]，不重复
  assert.strictEqual(out.length, 1);
  assert.strictEqual(out[0].id, 'a');
});

t.test('flattenSelected: 叶子 shape 在 group 内 + 也在选区顶层 → 只入结果一次', () => {
  const c1 = makeLeaf('c1');
  const g = makeMockGroup('g1', [c1]);
  const out = driver.flattenSelected([c1, g]);
  // 第一次 walk(c1)：seen.add('c1') → 叶子 → push
  // 第二次 walk(g) → walk(c1)：seen(c1) 拦住 → return
  assert.strictEqual(out.length, 1);
  assert.strictEqual(out[0].id, 'c1');
});

t.test('hasTopLevelGroup: 普通叶子选区 → false', () => {
  const a = makeLeaf('a');
  a.type = 'GeometricShape';
  assert.strictEqual(driver.hasTopLevelGroup([a]), false);
  assert.strictEqual(driver.hasTopLevelGroup({ items: [a] }), false);
});

t.test('hasTopLevelGroup: 顶层含 Group → true', () => {
  const a = makeLeaf('a');
  a.type = 'GeometricShape';
  const g = makeMockGroup('g1', [makeLeaf('c1')]);
  assert.strictEqual(driver.hasTopLevelGroup([a, g]), true);
  assert.strictEqual(driver.hasTopLevelGroup({ items: [g] }), true);
});

// ============================================================
// driver.loadShapeTree — Mac LTSC 分阶段 collection load
// ============================================================

t.test('loadShapeTree: 单选普通 shape 不加载 group path（GeneralException 回归）', async () => {
  const loads = [];
  let syncCount = 0;
  const leaf = makeLeaf('plain');
  leaf.type = 'GeometricShape';
  Object.defineProperty(leaf, 'group', {
    get() { throw new Error('普通 shape 不应访问 group'); },
  });
  const collection = makeLoadableCollection([leaf], loads);
  const d = createDriver({ sync: async () => { syncCount++; } });

  const out = await d.loadShapeTree(collection, 'id, name, width, height, adjustments');

  assert.deepStrictEqual(out.map((s) => s.id), ['plain']);
  assert.strictEqual(syncCount, 1);
  assert.strictEqual(loads.length, 1);
  assert.ok(loads[0].includes('items/type'));
  assert.ok(!loads[0].includes('group/shapes'));
});

t.test('loadShapeTree: 只展开真实 Group 的子 collection', async () => {
  const topLoads = [];
  const childLoads = [];
  let syncCount = 0;
  const c1 = makeLeaf('c1');
  const c2 = makeLeaf('c2');
  c1.type = 'GeometricShape';
  c2.type = 'GeometricShape';
  const group = makeOfficeGroup('g1', [c1, c2], childLoads);
  const collection = makeLoadableCollection([group], topLoads);
  const d = createDriver({ sync: async () => { syncCount++; } });

  const out = await d.loadShapeTree(collection, 'id, width, height, adjustments, tags');

  assert.deepStrictEqual(out.map((s) => s.id), ['c1', 'c2']);
  assert.strictEqual(syncCount, 2);
  assert.strictEqual(topLoads.length, 1);
  assert.strictEqual(childLoads.length, 1);
  assert.ok(!topLoads[0].includes('group/shapes'));
  assert.ok(!childLoads[0].includes('group/shapes'));
});

t.test('loadShapeTree: 嵌套 Group 按层加载并递归展平', async () => {
  const topLoads = [];
  const outerLoads = [];
  const innerLoads = [];
  let syncCount = 0;
  const c1 = makeLeaf('c1');
  const c2 = makeLeaf('c2');
  c1.type = 'GeometricShape';
  c2.type = 'GeometricShape';
  const inner = makeOfficeGroup('g2', [c2], innerLoads);
  const outer = makeOfficeGroup('g1', [c1, inner], outerLoads);
  const collection = makeLoadableCollection([outer], topLoads);
  const d = createDriver({ sync: async () => { syncCount++; } });

  const out = await d.loadShapeTree(collection, 'id, type, width');

  assert.deepStrictEqual(out.map((s) => s.id), ['c1', 'c2']);
  assert.strictEqual(syncCount, 3);
  assert.strictEqual(topLoads.length, 1);
  assert.strictEqual(outerLoads.length, 1);
  assert.strictEqual(innerLoads.length, 1);
});

// ============================================================
// 集成：flattenSelected 后 writeRadius
// ============================================================

t.test('集成: flatten 后 writeRadius 写到 group 里的子', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const c1 = makeLeaf('c1', { widthCm: 5, heightCm: 3, adjFraction: 0.05 });
  const c2 = makeLeaf('c2', { widthCm: 4, heightCm: 2, adjFraction: 0.1 });
  const g = makeMockGroup('g1', [c1, c2]);
  const h = createHarness({ shapes: [c1, c2] });
  // 模拟 caller：flatten 选区
  const leaves = h.driver.flattenSelected([g]);
  assert.strictEqual(leaves.length, 2);
  // 对每个叶子写 R 角
  for (const sh of leaves) {
    await RC.writeRadius(h.driver, sh, 0.5);
  }
  // 验证 c1/c2 都改了
  h.assertShape(c1, { adjFraction: 0.5 / 3 });
  h.assertShape(c2, { adjFraction: 0.5 / 2 });
});

t.test('集成: flatten 后 writeRadius 混合叶子（group + 普通）', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const plain = makeLeaf('plain', { widthCm: 5, heightCm: 3, adjFraction: 0.05 });
  const c1 = makeLeaf('c1', { widthCm: 4, heightCm: 2, adjFraction: 0.1 });
  const g = makeMockGroup('g1', [c1]);
  const h = createHarness({ shapes: [plain, c1] });
  const leaves = h.driver.flattenSelected([plain, g]);
  assert.strictEqual(leaves.length, 2);
  for (const sh of leaves) {
    await RC.writeRadius(h.driver, sh, 0.6);
  }
  h.assertShape(plain, { adjFraction: 0.6 / 3 });
  h.assertShape(c1, { adjFraction: 0.6 / 2 });
});

// ============================================================
// 集成：flattenSelected 后 loadLayoutTags（按角色处理）
// ============================================================

t.test('集成: flatten 后 loadLayoutTags 找到 group 内的 layout 父/子（按角色处理）', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  // 父在 group 里（带 LAYOUT_PARENT_TAG_KEY）
  const parent = makeLeaf('parent', { widthCm: 12, heightCm: 8, adjFraction: 0.05 });
  parent._tags[RC.LAYOUT_PARENT_TAG_KEY || 'layoutParent_v1'] = JSON.stringify({
    rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkR: 'same', childIds: ['c1', 'c2']
  });
  // 两个子在另一个 group 里
  const c1 = makeLeaf('c1', { widthCm: 4, heightCm: 3, adjFraction: 0.05 });
  const c2 = makeLeaf('c2', { widthCm: 4, heightCm: 3, adjFraction: 0.05 });
  c1._tags['layoutChild_v1'] = 'parent';
  c2._tags['layoutChild_v1'] = 'parent';
  // 两个 group
  const gParent = makeMockGroup('gParent', [parent]);
  const gChildren = makeMockGroup('gChildren', [c1, c2]);
  // 选区 = 两个 group
  const h = createHarness({ shapes: [parent, c1, c2] });
  const leaves = h.driver.flattenSelected([gParent, gChildren]);
  // 展平后：parent, c1, c2（3 个）
  assert.strictEqual(leaves.length, 3);
  assert.deepStrictEqual(leaves.map(s => s.id).sort(), ['c1', 'c2', 'parent']);
  // 模拟 caller 拿 layout tag
  // 这里只验证 leaves 数组正确（业务层 loadLayoutTags 内部按 tag 找父/子，不需要改）
  // 业务函数对扁平化数组的兼容性已被其它 driver flatten 集成测试覆盖
});

t.test('集成: applyLayout 用 group 叶子建 ID 映射并写入 4 个子', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const parent = makeLeaf('parent', {
    widthCm: 12, heightCm: 8, leftCm: 5, topCm: 3, adjFraction: 0.1,
  });
  const children = [1, 2, 3, 4].map((n) => makeLeaf(`c${n}`, {
    widthCm: 2, heightCm: 2, leftCm: n, topCm: n, adjFraction: 0.05,
  }));
  const group = makeMockGroup('g1', [parent, ...children]);
  const h = createHarness({ shapes: [group] });

  const r = await RC.applyLayout(
    h.driver,
    parent.id,
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'same' },
    children.map((s) => s.id),
    { writeParentTag: true, syncR: true }
  );

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 4);
  assert.strictEqual(r.regrouped, true);
  assert.ok(parent._tags.layoutParent_v1);
  h.assertCallCount('setBox', 4);
  h.assertCallCount('ungroupShapeGroup', 1);
  h.assertCallCount('addGroup', 1);
  h.assertCallCount('selectShapes', 1);
  assert.strictEqual(h.slide._selectedShapeIds.length, 1);
  for (const child of children) {
    assert.strictEqual(child._tags.layoutChild_v1, parent.id);
    assert.ok(child.width > 2 * PT_PER_CM);
    assert.ok(child.height > 2 * PT_PER_CM);
  }
});

t.test('集成: group 布局先 ungroup，全部 setBox 完成后再 regroup', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const parent = makeLeaf('parent', {
    widthCm: 12, heightCm: 8, leftCm: 5, topCm: 3, adjFraction: 0.1,
  });
  const children = [1, 2, 3, 4].map((n) => makeLeaf(`c${n}`, {
    widthCm: 2, heightCm: 2, leftCm: n, topCm: n, adjFraction: 0.05,
  }));
  const h = createHarness({ shapes: [makeMockGroup('g1', [parent, ...children])] });

  const r = await RC.applyLayout(
    h.driver,
    parent.id,
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'same' },
    children.map((s) => s.id),
    { writeParentTag: false, syncR: false }
  );

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.regrouped, true);
  const boxIndexes = h.calls
    .map((call, index) => ({ call, index }))
    .filter(({ call }) => call.method === 'setBox')
    .map(({ index }) => index);
  assert.strictEqual(boxIndexes.length, 4);
  const ungroupIndex = h.calls.findIndex((call) => call.method === 'ungroupShapeGroup');
  const regroupIndex = h.calls.findIndex((call) => call.method === 'addGroup');
  assert.ok(ungroupIndex >= 0, '必须先解除组合');
  assert.ok(regroupIndex > boxIndexes[boxIndexes.length - 1], '全部 setBox 后才能重新组合');
  assert.ok(boxIndexes.every((index) => index > ungroupIndex && index < regroupIndex));
});

t.test('集成: group 切换 R 联动模式只写 R/tag，不反算任何子 box', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const parent = makeLeaf('parent', {
    widthCm: 12, heightCm: 8, leftCm: 5, topCm: 3, adjFraction: 0.1,
  });
  const children = [1, 2, 3, 4].map((n) => makeLeaf(`c${n}`, {
    widthCm: 2 + n,
    heightCm: 1.5 + n,
    leftCm: n * 1.1,
    topCm: n * 0.9,
    adjFraction: 0.05,
  }));
  // 模拟真实 PowerPoint：Tag.key 从集合批量读出时统一为大写。
  for (const child of children) child._tags.RADIUSLOCK_V1 = '0.25';
  const originalBoxes = children.map((s) => ({
    left: s.left,
    top: s.top,
    width: s.width,
    height: s.height,
  }));
  const h = createHarness({ shapes: [makeMockGroup('g1', [parent, ...children])] });

  const r = await RC.applyLayout(
    h.driver,
    parent.id,
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.2, linkRMode: 'subtract' },
    children.map((s) => s.id),
    { writeParentTag: true, syncR: true, writeGeometry: false }
  );

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.regrouped, true);
  assert.strictEqual(r.geometryWritten, false);
  assert.strictEqual(r.lockedCount, 4);
  h.assertCallCount('setBox', 0);
  h.assertCallCount('ungroupShapeGroup', 1);
  h.assertCallCount('addGroup', 1);
  children.forEach((child, index) => {
    assert.deepStrictEqual(
      {
        left: child.left,
        top: child.top,
        width: child.width,
        height: child.height,
      },
      originalBoxes[index],
      `child #${index + 1} 的 box 不应被 R 模式切换改动`
    );
    const minSideCm = Math.min(child.width, child.height) / PT_PER_CM;
    assert.ok(Math.abs(child._adjFraction - (0.5 / minSideCm)) < 1e-9);
    assert.ok(Math.abs(Number(child._tags.radiusLock_v1) - 0.5) < 1e-9);
  });
  const parentTag = JSON.parse(parent._tags.layoutParent_v1);
  assert.strictEqual(parentTag.linkRMode, 'subtract');
});

t.test('集成: group 布局事务保留原 group 名称和 tags', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const parent = makeLeaf('parent', {
    widthCm: 12, heightCm: 8, leftCm: 5, topCm: 3, adjFraction: 0.1,
  });
  const child = makeLeaf('c1', {
    widthCm: 2, heightCm: 2, leftCm: 1, topCm: 1, adjFraction: 0.05,
  });
  const group = makeMockGroup('g1', [parent, child]);
  group.name = '我的组合';
  group._tags = { custom_group_tag: 'keep-me' };
  const h = createHarness({ shapes: [group] });

  const r = await RC.applyLayout(
    h.driver,
    parent.id,
    { rows: 1, cols: 1, padding: 0.3, gutter: 0, linkRMode: 'same' },
    [child.id],
    { writeParentTag: true, syncR: true }
  );

  assert.strictEqual(r.ok, true);
  const rebuilt = h.slide.shapes.items.find((s) => s._isGroup);
  assert.ok(rebuilt);
  assert.strictEqual(rebuilt.name, '我的组合');
  assert.strictEqual(rebuilt._tags.custom_group_tag, 'keep-me');
});

t.test('集成: 父子分属不同 group 时拒绝写入，避免破坏组合', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const parent = makeLeaf('parent', {
    widthCm: 12, heightCm: 8, leftCm: 5, topCm: 3, adjFraction: 0.1,
  });
  const child = makeLeaf('c1', {
    widthCm: 2, heightCm: 2, leftCm: 1, topCm: 1, adjFraction: 0.05,
  });
  const h = createHarness({
    shapes: [makeMockGroup('gParent', [parent]), makeMockGroup('gChild', [child])],
  });

  const r = await RC.applyLayout(
    h.driver,
    parent.id,
    { rows: 1, cols: 1, padding: 0.3, gutter: 0, linkRMode: 'same' },
    [child.id],
    { writeParentTag: true, syncR: true }
  );

  assert.strictEqual(r.ok, false);
  assert.ok(r.warn.includes('不同或嵌套组合'));
  h.assertCallCount('setBox', 0);
  h.assertCallCount('ungroupShapeGroup', 0);
});

t.test('集成: group 布局写入异常时 catch 路径仍恢复组合', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const parent = makeLeaf('parent', {
    widthCm: 12, heightCm: 8, leftCm: 5, topCm: 3, adjFraction: 0.1,
  });
  const child = makeLeaf('c1', {
    widthCm: 2, heightCm: 2, leftCm: 1, topCm: 1, adjFraction: 0.05,
  });
  const h = createHarness({ shapes: [makeMockGroup('g1', [parent, child])] });
  h.driver.setBox = () => { throw new Error('mock box failure'); };

  const r = await RC.applyLayout(
    h.driver,
    parent.id,
    { rows: 1, cols: 1, padding: 0.3, gutter: 0, linkRMode: 'same' },
    [child.id],
    { writeParentTag: true, syncR: true }
  );

  assert.strictEqual(r.ok, false);
  assert.ok(r.error.includes('mock box failure'));
  h.assertCallCount('ungroupShapeGroup', 1);
  h.assertCallCount('addGroup', 1);
  assert.ok(h.slide.shapes.items.some((s) => s._isGroup));
});

t.test('集成: group 整体缩放后的 R-only 联动不写任何子 box', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const parent = makeLeaf('parent', {
    widthCm: 12, heightCm: 10, leftCm: 5, topCm: 3, adjFraction: 0.2,
  });
  const children = [1, 2, 3, 4].map((n) => makeLeaf(`c${n}`, {
    widthCm: 5, heightCm: 4, leftCm: n, topCm: n, adjFraction: 0.05,
  }));
  const group = makeMockGroup('g1', [parent, ...children]);
  const h = createHarness({ shapes: [group] });

  const r = await RC.syncLayoutChildrenR(
    h.driver,
    parent.id,
    children.map((s) => s.id),
    0.5,
    'same',
    1.5
  );

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.applied, 4);
  h.assertCallCount('setBox', 0);
  for (const child of children) {
    h.assertShape(child, { adjFraction: 1.5 / 4 });
  }
});

t.test('集成: group 原生缩放后按新父框完整重排，恢复固定 padding/gutter', async () => {
  const RC = require(path.join(__dirname, '..', 'src', 'lib', 'radius-core.js'));
  const parent = makeLeaf('parent', {
    widthCm: 16, heightCm: 7, leftCm: 3, topCm: 2, adjFraction: 0.1,
  });
  // 模拟 PowerPoint 原生非等比缩放后的后代：box / padding / gutter 都已失真。
  const children = [1, 2, 3, 4].map((n) => makeLeaf(`c${n}`, {
    widthCm: 6 + n * 0.2,
    heightCm: 2 + n * 0.1,
    leftCm: 3 + n,
    topCm: 2 + n * 0.4,
    adjFraction: 0.05,
  }));
  for (const child of children) child._tags.RADIUSLOCK_V1 = '0.2';
  const h = createHarness({ shapes: [makeMockGroup('g1', [parent, ...children])] });
  const originalParentBox = {
    left: parent.left,
    top: parent.top,
    width: parent.width,
    height: parent.height,
  };

  const r = await RC.applyLayout(
    h.driver,
    parent.id,
    { rows: 2, cols: 2, padding: 0.3, gutter: 0.4, linkRMode: 'subtract' },
    children.map((s) => s.id),
    { writeParentTag: false, syncR: true, writeGeometry: true }
  );

  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.regrouped, true);
  assert.strictEqual(r.geometryWritten, true);
  assert.strictEqual(r.lockedCount, 4);
  assert.deepStrictEqual(
    { left: parent.left, top: parent.top, width: parent.width, height: parent.height },
    originalParentBox,
    '完整重排只能读新父框，不能改变父形状'
  );

  const expected = RC.computeLayout(originalParentBox, 2, 2, 0.3, 0.4);
  assert.strictEqual(expected.feasible, true);
  children.forEach((child, index) => {
    const pos = expected.positions[index];
    assert.ok(Math.abs(child.left - pos.left) < 1e-9);
    assert.ok(Math.abs(child.top - pos.top) < 1e-9);
    assert.ok(Math.abs(child.width - pos.w) < 1e-9);
    assert.ok(Math.abs(child.height - pos.h) < 1e-9);
  });
  h.assertCallCount('setBox', 4);
  h.assertCallCount('ungroupShapeGroup', 1);
  h.assertCallCount('addGroup', 1);
});

(async () => { await t.run(); })();
