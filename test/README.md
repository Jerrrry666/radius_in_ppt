# R 角调整 — 测试

## 跑测试

```bash
npm test
```

只跑一个：

```bash
node test/test-radius-core.js         # 46 个 — 纯算法
node test/test-features.js            # 49 个 — 功能（业务函数）
```

## 测试分层

| 层 | 文件 | 测什么 | 跑不跑 |
|---|---|---|---|
| **driver 层** | `ppt-driver.js` 16 个方法 | Mac LTSC Office.js 兼容性 | **不在 npm test 里**——在真实 PPT 跑"Driver 烟囱测试" |
| **纯算法** | `test-radius-core.js` | `computeLayout` / `valueToCm` / 业务规则（`shouldReject*` / `syncFixedValueIfLocked`） | npm test |
| **功能** | `test-features.js` | 业务函数（`writeRadius` / `applyLayout` / `syncLayoutChildrenR` / `readLockState` / `writeLockState` / `reapplyLock`）—— "模拟交互反馈" | npm test |

**v1.3 重整后**：功能测试用 `assertShape` 验最终状态，**不关心 driver 内部调了哪些方法**。

## 写新功能怎么测（v1.3 流程）

### 1. 拿标准 fixture

```js
const { createHarness, makeStandardFixture } = require('./test-harness');
const f = makeStandardFixture();
// 标准 5+ R 角矩形：
//   f.shapes.r1_basic          — 普通 R 角矩形（5×3cm）
//   f.shapes.r2_medium         — 已有 R 角（8×4cm, adj=0.1）
//   f.shapes.r3_large          — 大矩形（12×8cm, adj=0.2）
//   f.shapes.r4_tiny           — 小矩形（2×1.5cm）
//   f.shapes.r5_wide           — 宽矩形（20×5cm）
//   f.shapes.r6_locked         — locked, radiusLock_v1=0.8
//   f.shapes.r7_strict         — strict, radiusLockStrict_v1=1
//   f.shapes.r8_lockedStrict   — locked + strict 同时
//   f.shapes.r9_clampEdge      — clamp 边界（短边 1.058cm）
//   f.shapes.r10_zeroSize      — 0 尺寸
//   f.parent                   — layout 父（12×8cm, adj=0.3）
//   f.layoutChildren [lc1-lc4] — layout 子
//   f.rect1                    — 非圆角矩形
```

### 2. 建 harness

```js
const h = createHarness({ shapes: f.allShapes });
// h.driver   — driver 包装（功能测试不直接用，但 fixture 需要它作为 ctx）
// h.calls    — 所有 driver 方法调用记录（debug 用，不作为主断言）
// h.snapshot() — 当前所有 shape 状态
```

### 3. 调功能方法 + 验最终状态

```js
const RC = require('../src/lib/radius-core.js');

// 调功能
const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);

// 验证返回 + 形状最终状态
assert.strictEqual(r.ok, true);
assert.strictEqual(r.newCm, 0.5);
h.assertShape(f.shapes.r1_basic, {
  adjFraction: 0.5 / 3,        // R 角变成新值
  tags: {},                      // 没动
});
```

## 框架 API

### `createHarness({ shapes })`

返回：
- `driver` — driver 实例
- `calls` — `[{method, args, time}]` 数组（debug 用，不作为主断言）
- `shapes` — 输入的 shape 数组
- `snapshot()` — 当前所有 shape 的状态
- **`assertShape(shape, expected)`** — **主断言**。验证 shape 最终状态
- `dumpCalls()` / `reset()` — debug 工具

### `assertShape(shape, expected)`

`expected` 字段（都可省略）：
- `adjFraction` — 数字 / 谓词函数（(val) => bool）
- `tags` — `{key: value}` map，value=undefined 表示"该 key 不存在"，value=函数表示"满足谓词"
- `box` — `{left, top, width, height}`，数字或谓词

### `createTestRunner()`

```js
const t = createTestRunner();
t.test('name', async () => { /* ... */ });
t.test('name', () => { /* ... */ });
t.beforeEach(() => { /* 每个 test 前跑 */ });
t.afterEach(() => { /* 每个 test 后跑 */ });
await t.run();
```

## 覆盖范围

### 1. 纯算法（test-radius-core.js，46 个）

| 类别 | 数量 | 例子 |
| --- | --- | --- |
| `computeLayout` 布局 math | 5 | 2×2 / 1×3 / 3×2 / 1×1 / 不可行 |
| 单位换算 | 6 | cm↔cm / %↔cm / 边界 |
| 联动公式 `computeLinkedSubR` | 5 | subtract（3 边界）/ same（1）/ off（2）|
| `clampRadius` | 4 | 不超 / 超 / 负值 / minSide=0 |
| `cmToAdj` | 3 | 基础 / 短边一半 / minSide=0 |
| `computeFinalRadius` | 5 | off / subtract / same / clamp / parentR<padding |
| 业务规则 `shouldRejectWriteRadius` | 4 | 普通 / strict / 非圆角 / 0 尺寸 |
| 业务规则 `shouldRejectOnApply` | 4 | 空 / 全普通 / 含 strict / 非 roundRect strict 不算 |
| 业务规则 `shouldRejectLayoutApply` | 2 | 含 strict 子 / 父 strict 不影响 |
| 业务规则 `syncFixedValueIfLocked` | 3 | unlocked / locked / locked 写 0 |
| 集成场景 | 5 | 2×2 / 1×3 / 拒绝 / lock 同步 / 用户样例 |

### 2. 功能（test-features.js，49 个）

| 类别 | 数量 | 测的 |
| --- | --- | --- |
| writeRadius 基础 | 6 | 5+ fixture + clamp 边界 |
| writeRadius strict/locked | 3 | strict / locked / locked+strict |
| writeRadius 边界 | 6 | 0 尺寸 / 非圆角 / NaN / Infinity / layoutParentId / driver 异常 |
| 批量写 R 角 | 2 | 5 个全成功 / 5 个混合 |
| readLockState | 5 | 无 / lock / strict / 都有 / 非数字 |
| writeLockState | 5 | 写 lock / 删 lock / 写 strict / undefined 不动 / 同时删两个 |
| reapplyLock | 6 | 基础 / clamp / 非圆角 / 0 尺寸 / 负数 / 恢复 |
| applyLayout | 8 | 2x2 / off / same / 父不在 / 子不足 / stale / writeParentTag=false / infeasible |
| syncLayoutChildrenR | 7 | subtract / same / off / parentRcm=0 / stale / strict / 非圆角 |
| 自测场景 | 1 | 批量写 → 锁定 → 再写 → layout 联动 |

## 测试覆盖原则

按"**防误触 = 最高优先级**"原则（见 AGENTS.md 1.4）：

- ✅ 任何 R 角写入路径都不能 skip strict
- ✅ 任何 R 角写入路径都要检查 lock 状态
- ✅ locked 形状被 R 角写入 → 同步 fixed value
- ✅ 含 strict 的 layout apply → 整个拒绝（包括位置/尺寸）

## 添加新功能

写新功能的测试：

```js
const f = makeStandardFixture();
const h = createHarness({ shapes: f.allShapes });

const r = await yourNewFunction(h.driver, f.shapes.r1_basic, /* args */);

h.assertShape(f.shapes.r1_basic, {
  adjFraction: 0.5 / 3,  // R 角变成新值
  tags: { /* 期望的 tag 状态 */ },
});
```

复杂场景用 `createTestRunner()`：

```js
const t = createTestRunner();
t.test('批量写 5 个普通 R 角矩形', async () => {
  const f = makeStandardFixture();
  const h = createHarness({ shapes: f.allShapes });
  for (const s of Object.values(f.shapes).slice(0, 5)) {
    await RC.writeRadius(h.driver, s, 0.3);
  }
  h.assertShape(f.shapes.r1_basic, { adjFraction: 0.3 / 3 });
  // ... 其它 shape
});
await t.run();
```

## 真实 PPT 测试（**driver 层**）

**driver 单元测试不在 npm test 里**——在真实 PPT 内做：

- 跑 .app，在真实 PPT 里打开 task pane
- 点「🧪 Driver 烟囱测试」按钮 → 14/14 全过即 driver verified
- 改了 `ppt-driver.js` 后必跑这个（"只要没有新的交互操作，就不用再运行"——意思是新加 driver 方法才需要）

**功能测试不能替代的**：
- Mac LTSC PowerPoint.js bug（per-shape load adjustments 不 work / get(0) ClientResult）
- shape.tags 真实持久化（关 PPT → 重开是否还在）
- 跨 page 隔离 / lock monitor 真实反算

## driver 协议

`radius-core.writeRadius(driver, shape, ...)` 接受 `createDriver(ctx)` 返回的对象：

```js
driver = {
  // 加载 + 同步
  load(proxy, fields),      // proxy.load(fields)
  sync(),                   // ctx.sync()

  // Collection accessors
  selectedShapes(),         // ctx.presentation.getSelectedShapes()
  activeSlide(),            // ctx.presentation.getSelectedSlides().getItemAt(0)
  slideShapes(slide),       // slide.shapes

  // 读（假定已 load + sync）
  shapeId(s),               // s.id
  size(s),                  // { width, height }
  box(s),                   // { left, top, width, height }
  isRoundRect(s),           // s.adjustments.count > 0
  adjFraction(s),           // s.adjustments.get(0).value (0~1)
  loadAdjValue(s),          // s.adjustments.load('items/value')

  // 写（假定已 load）
  setBox(s, box),
  setAdjFraction(s, frac),

  // Tag 操作
  addTag(s, key, value),
  deleteTag(s, key),
  readTag(s, key),          // async
}
```

## mock shape 协议

```js
shape = {
  id: string,
  width: number (pt),
  height: number (pt),
  left: number (pt),
  top: number (pt),
  adjustments: {
    count: number,
    get(0): { value: number },  // 0~1 比例
    set(0, value): void,
  },
  tags: { [key]: value },
}
```

测试用 `makeFixtureShape({...})`（fixtures.js）创建 mock shape。
