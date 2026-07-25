/*
 * fixtures.js — 标准 R 角矩形 fixture（test 共享）
 *
 * 提供一组预设好状态的 R 角矩形，覆盖主要业务场景：
 *   - 普通 R 角矩形（多个尺寸）
 *   - 已设置 R 角的形状（adjFraction != 0）
 *   - locked（固定 R 角值）
 *   - strict（防误触）
 *   - locked + strict 同时
 *   - clamp 边界（短边 30pt + R 角 = 0.5cm）
 *   - 0 尺寸
 *   - 非圆角矩形（adjCount=0）
 *   - layout 父（用于 applyLayout 测试）
 *
 * 用法：
 *   const { makeStandardFixture } = require('./fixtures');
 *   const { shapes, parent } = makeStandardFixture();
 *   const harness = createHarness({ shapes: Object.values(shapes).concat([parent]) });
 *   // ...调业务方法，验证 driver 反应
 *
 * 设计原则：
 *   - 用 cm（业务单位）声明尺寸，内部换算成 pt 喂给 mock shape
 *   - 字段命名跟业务一致（adjFraction 而非 _adjFraction）
 *   - 不在 fixture 里包含 driver 行为（harness 负责）
 */

const PT_PER_CM = 28.3464567;
const cm = (c) => c * PT_PER_CM;

/**
 * 创建一个 mock shape（fixture 专用，简洁版）
 * @param {Object} init
 *   - id: string
 *   - widthCm: number (cm)
 *   - heightCm: number (cm)
 *   - leftCm: number (cm, 默认 0)
 *   - topCm: number (cm, 默认 0)
 *   - adjFraction: number (0~1, 默认 0)
 *   - isRoundRect: boolean (默认 true)
 *   - tags: object (默认 {})
 * @returns {Object} mock shape
 */
function makeFixtureShape(init) {
  init = init || {};
  const isRR = init.isRoundRect !== false;
  const tags = Object.assign({}, init.tags || {});
  const shape = {
    id: init.id,
    width: cm(init.widthCm != null ? init.widthCm : 4),
    height: cm(init.heightCm != null ? init.heightCm : 2),
    left: cm(init.leftCm != null ? init.leftCm : 0),
    top: cm(init.topCm != null ? init.topCm : 0),
    _adjFraction: init.adjFraction != null ? init.adjFraction : 0,
    _tags: tags,
  };
  // adjustments 协议（跟 ppt-driver 兼容）
  shape.adjustments = {
    count: isRR ? 1 : 0,
    get() { return { value: shape._adjFraction }; },
    set(_, v) { shape._adjFraction = v; },
  };
  return shape;
}

/**
 * 标准 fixture：5+ R 角矩形 + 边界 + layout 父
 *
 * 返回的 shape 全部以 cm 声明尺寸，命名清晰可读。
 *
 * IDs 命名规则：
 *   - rN_xxx  → 普通 R 角矩形
 *   - rN_xxx_locked → 已 locked
 *   - rN_xxx_strict → strict
 *   - rN_xxx_both → locked + strict
 *   - rN_xxx_edge → clamp 边界
 *   - rN_xxx_zero → 0 尺寸
 *   - rectN → 非圆角矩形
 *   - parent_p1 → layout 父
 *
 * @returns {Object} { shapes: {id: shape}, parent: shape, layoutChildren: [shape] }
 */
function makeStandardFixture() {
  const shapes = {};

  // ── 普通 R 角矩形（5 个不同尺寸，覆盖大部分场景）──
  shapes.r1_basic = makeFixtureShape({
    id: 'r1_basic',
    widthCm: 5, heightCm: 3,
    adjFraction: 0,
  });
  shapes.r2_medium = makeFixtureShape({
    id: 'r2_medium',
    widthCm: 8, heightCm: 4,
    adjFraction: 0.1,  // 已有 R 角：0.1 * minSide(4) = 0.4cm
  });
  shapes.r3_large = makeFixtureShape({
    id: 'r3_large',
    widthCm: 12, heightCm: 8,
    adjFraction: 0.2,  // 已有 R 角：0.2 * minSide(8) = 1.6cm
  });
  shapes.r4_tiny = makeFixtureShape({
    id: 'r4_tiny',
    widthCm: 2, heightCm: 1.5,
    adjFraction: 0,
  });
  shapes.r5_wide = makeFixtureShape({
    id: 'r5_wide',
    widthCm: 20, heightCm: 5,
    adjFraction: 0.3,
  });

  // ── locked（固定 R 角值）──
  shapes.r6_locked = makeFixtureShape({
    id: 'r6_locked',
    widthCm: 6, heightCm: 4,
    adjFraction: 0.2,  // 当前 R = 0.8cm
    tags: { radiusLock_v1: '0.8' },
  });

  // ── strict（防误触）──
  shapes.r7_strict = makeFixtureShape({
    id: 'r7_strict',
    widthCm: 5, heightCm: 5,
    adjFraction: 0,
    tags: { radiusLockStrict_v1: '1' },
  });

  // ── locked + strict 同时（理论 onApply 会提前拒绝，writeRadius 也拦截）──
  shapes.r8_lockedStrict = makeFixtureShape({
    id: 'r8_lockedStrict',
    widthCm: 4, heightCm: 4,
    adjFraction: 0.15,
    tags: { radiusLock_v1: '0.6', radiusLockStrict_v1: '1' },
  });

  // ── clamp 边界：短边 30pt（约 1.06cm），R 角 = 0.5cm（接近短边一半）──
  // 30pt = 1.058cm，max R = 0.529cm
  // 当前 adj = 0.47 → R = 0.47 * 1.058 = 0.497cm
  shapes.r9_clampEdge = makeFixtureShape({
    id: 'r9_clampEdge',
    widthCm: 5, heightCm: 1.058,  // 短边 30pt
    adjFraction: 0.47,
  });

  // ── 0 尺寸（异常）──
  shapes.r10_zeroSize = makeFixtureShape({
    id: 'r10_zeroSize',
    widthCm: 0, heightCm: 0,
    adjFraction: 0,
  });

  // ── layout 父：12×8cm，R 角 0.3（约 0.3 * 8 = 2.4cm R 角）──
  const parent = makeFixtureShape({
    id: 'parent_p1',
    widthCm: 12, heightCm: 8,
    leftCm: 5, topCm: 3,
    adjFraction: 0.3,
  });

  // ── layout 子：4 个普通 R 角矩形（2×2 用）──
  const layoutChildren = [
    makeFixtureShape({ id: 'lc1', widthCm: 2.5, heightCm: 1.5, adjFraction: 0 }),
    makeFixtureShape({ id: 'lc2', widthCm: 2.5, heightCm: 1.5, adjFraction: 0 }),
    makeFixtureShape({ id: 'lc3', widthCm: 2.5, heightCm: 1.5, adjFraction: 0 }),
    makeFixtureShape({ id: 'lc4', widthCm: 2.5, heightCm: 1.5, adjFraction: 0 }),
  ];

  // ── 非圆角矩形（普通矩形）──
  const rect1 = makeFixtureShape({
    id: 'rect1_plain',
    widthCm: 6, heightCm: 3,
    isRoundRect: false,
  });

  return {
    shapes,
    parent,
    layoutChildren,
    rect1,
    // 全部 shape 一维数组（方便喂给 harness）
    allShapes: Object.values(shapes).concat([parent, ...layoutChildren, rect1]),
  };
}

module.exports = {
  makeStandardFixture,
  makeFixtureShape,
  PT_PER_CM,
  cm,
};
