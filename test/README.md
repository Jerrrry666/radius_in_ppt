# R 角调整 — 测试

## 跑测试

```bash
npm test
```

只跑一个：

```bash
node test/test-radius-core.js         # 46 个 — 纯算法
node test/test-driver-integration.js   # 54 个 — driver 集成（新框架）
```

## 测试结构

```
test/
├── README.md
├── fixtures.js                       # 标准 5+ R 角矩形 fixture
├── test-harness.js                   # driver harness + call tracker + test runner
├── test-radius-core.js               # 46 个 — 纯算法（computeLayout / 单位换算 / 业务规则）
└── test-driver-integration.js        # 54 个 — driver + 业务方法集成
```

**v1.3 重整后**：
- 删了 `test-mock-harness.js`（70 个）—— 跟 driver 集成完全重复
- 删了 `radius-core.writeRadiusToShapePure` / `applyLayoutPure` —— 业务只走 driver 路径
- 100 个测试，**全面不重复**

## 写新功能怎么测（v1.3 流程）

底层（ppt-driver.js）和图形交互（dialog.js）已经稳了。新功能不用每次都连真 PPT 测，按下面三步走：

### Step 1：拿标准 fixture

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

### Step 2：建 harness

```js
const h = createHarness({ shapes: f.allShapes });
// h.driver   — 包装过的 driver，方法调用全记录
// h.calls    — 所有 driver 方法调用 [{method, args, time}]
// h.shapes   — 所有 shape Map
// h.snapshot() — 当前所有 shape 状态
```

### Step 3：调业务方法 + 验证 driver 反应

```js
const RC = require('../src/lib/radius-core.js');

// 调业务方法
const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);

// 验证返回值
assert.strictEqual(r.ok, true);
assert.strictEqual(r.newCm, 0.5);

// 验证 driver 反应（"调用底层编辑工具，测试框架给出反应"）
h.assertCalled('setAdjFraction');              // setAdjFraction 被调
h.assertCalled('setAdjFraction', {             // 带参数验证
  with: ['r1_basic', 0.5 / 3]
});
h.assertNotCalled('addTag');                   // addTag 没被调
h.assertCallCount('readTag', 2);               // readTag 调 2 次
h.assertShape(f.shapes.r1_basic, {             // shape 状态对
  adjFraction: 0.5 / 3,
  tags: {},
});
```

## 框架 API 速查

### `createHarness({ shapes })`

返回：
- `driver` — 真实 createDriver + 全部 16 个方法被 recordCall 包装
- `calls` — `[{method, args, time}]` 数组
- `shapes` — 输入的 shape 数组
- `snapshot()` — 当前所有 shape 的 {id, width, height, left, top, adjFraction, tags}
- `assertCalled(method, { with })` — 验证 driver 方法被调
- `assertNotCalled(method)` — 验证 driver 方法没被调
- `assertCallCount(method, n)` — 验证 driver 方法被调 n 次
- `assertShape(shape, expected)` — 验证 shape 状态
- `dumpCalls(filter?)` — 打印所有 call（debug 用）
- `reset()` — 清空 calls

### `assertShape(shape, expected)`

`expected` 字段（都可省略）：
- `adjFraction` — 数字 / 谓词函数（(val) => bool）
- `tags` — `{key: value}` map，value=undefined 表示"该 key 不存在"
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

### 2. driver 集成（test-driver-integration.js，54 个）

**框架**：用 `fixtures.js` 标准 5+ R 角矩形 + `createHarness` 包装 driver。

测试风格：调业务方法 → 验证 driver 反应 + shape 状态。

| 类别 | 数量 | 验证 |
| --- | --- | --- |
| writeRadius 基础 | 5+ | 每个 fixture 一个用例（basic/medium/large/tiny/wide） |
| writeRadius strict/locked | 3 | strict / locked / locked+strict |
| writeRadius 边界 | 5 | 0 尺寸 / 非圆角 / NaN / Infinity / layoutParentId |
| writeRadius driver 异常 | 1 | reason=exception, error 含 message |
| 批量 5+ | 2 | 5 个全成功 / 5 个混合状态 |
| readLockState / writeLockState | 8 | 各种 tag 状态读写 |
| reapplyLock | 5 | 基础 / clamp / 0 尺寸 / 非圆角 / 负数 |
| applyLayout | 8 | 2x2 / off / same / 父不在 / 子不足 / stale / writeParentTag=false / infeasible |
| syncLayoutChildrenR | 7 | subtract / same / off / parentRcm=0 / stale / strict / 非圆角 |
| driver API 一致性 | 4 | 16 方法 / adjFraction defensive / size vs box / v1.2.2 回归 |
| 自测场景 | 4 | 5+ 形状组合 / clamp 边界 / reapplyLock / 多 slide |

## 测试覆盖原则

按"**防误触 = 最高优先级**"原则（见 AGENTS.md 1.4）：

- ✅ 任何 R 角写入路径都不能 skip strict
- ✅ 任何 R 角写入路径都要检查 lock 状态
- ✅ locked 形状被 R 角写入 → 同步 fixed value
- ✅ 含 strict 的 layout apply → 整个拒绝（包括位置/尺寸）

## 添加新测试

写新功能测试：

```js
// 1. 拿 fixture
const f = makeStandardFixture();

// 2. 建 harness
const h = createHarness({ shapes: f.allShapes });

// 3. 调新功能
const r = await yourNewFunction(h.driver, f.shapes.r1_basic, /* args */);

// 4. 验证 driver 反应 + shape 状态
h.assertCalled('setAdjFraction');
h.assertNotCalled('addTag');
h.assertShape(f.shapes.r1_basic, { adjFraction: 0.5 / 3 });
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
  h.assertCallCount('setAdjFraction', 5);
});
await t.run();
```

## 真实 PPT 测试（**需要人工**）

mock harness 测的是**逻辑流程**。**真实 PowerPoint.js API 行为**只能靠手工验证：

- 跑 .app，在真实 PPT 里建布局
- 看 task pane 底部「🔧 调试日志（点击展开）」
- 复制日志分析是否符合预期

**mock harness 不能测的**：
- Mac LTSC 的 PowerPoint.js bug（per-shape load adjustments 不 work、get(0) ClientResult 等）
- shape.tags 真实持久化（关 PPT → 重开是否还在）
- 跨 page 隔离是否真的生效
- lock monitor 是否真的反算 / 不同步

## driver 协议

`radius-core.writeRadius(driver, shape, ...)` 接受一个 `createDriver(ctx)` 返回的对象：

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

`writeRadiusToShapePure` 接受一个普通对象 `shape`（**v1.3 已删，业务只走 driver 路径**）：

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
