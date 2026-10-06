# R角调整 — 原生与历史对照测试

## Office.js历史对照测试

```bash
npm test
```

只跑一个：

```bash
node test/test-radius-core.js         # 115项 — 纯算法
node test/test-features.js            # 96项 — 功能（业务函数）
node test/test-driver-group.js        # 40项 — 分层树加载/组合事务
node test/test-regressions.js         # 36项 — 实际UI wiring、OOXML、错误/并发/HTTP回归
```

共287项；mock通过不等于Mac LTSC宿主API已验证，driver更新还需要真实PPT的14/14烟囱测试。

## 原生PPAM宿主回归

`test-native-package.py`验证格式/分发文件，不执行VBA。使用普通PPTX测试真实安装后的Ribbon，避免跨VBA项目的测试加载项影响宿主：

```bash
.venv-native/bin/pip install python-pptx
.venv-native/bin/python test/native-host-fixture.py build /tmp/RadiusNativeRegression.pptx
```

在PowerPoint打开该文稿。每页先激活画布、Esc、Cmd+A，使用「R角调整 · Native」执行：

| 页 | 操作及预期 |
| --- | --- |
| 1 | 0.50cm预设；父、子、普通圆角均写0.50，箭头不变，布局JSON保留 |
| 2 | 0.50cm；v1.4显示部分开启1/2并禁用所有R角写入，前面的普通形状也不变 |
| 3 | 开启防误触→写入按钮禁用→解除防误触→0.30cm |
| 4 | 20%→读回0.40cm；0cm→读回0；0.10cm→读回0.10；输入100cm→读回钳制值1.00cm |
| 5 | 0.30cm→开启防误触→写入按钮禁用→解除防误触→0.30cm；两层组合、名称、标签和几何保留 |
| 6 | 0.30cm；嵌套strict子在解组前拦截，全部ID/几何/标签/R角不变 |
| 7 | 显示请选择圆角矩形、按钮禁用；箭头不变 |

Cmd+S后运行独立保存状态校验，最终状态以表中最后一步为准：

```bash
.venv-native/bin/python test/native-host-fixture.py inspect /tmp/RadiusNativeRegression.pptx
.venv-native/bin/python test/native-host-fixture.py verify /tmp/RadiusNativeRegression.pptx
```

`verify`断言7页最终半径、原始标签、叶子ID、层级、名称和几何，独立于VBA源码。第4页中间值需实机读取，必要时用`inspect`保存阶段快照。2026-10-07在PowerPoint16.113.3通过；[本机保存状态](native-host-validation-20261007.json)不代表16.111或其他宿主版本已经验收。

## v1.4控件宿主协议

当前17项原生格式/安装测试；真实VBA算法自检10项。微调必须检查保存的实际半径，不能只看输入框变化。

| 页 | 控件操作与核对 |
| --- | --- |
| 1 | 混合多选，0.30cm预设→上箭头0.40→保存核对→上箭头0.50；3个圆角、箭头不变 |
| 2 | 选中全部，部分开启1/2；微调/预设/应用禁用，选区变化自动刷新 |
| 3 | 0cm，下箭头禁用→连续3次上箭头0.30→保护开启1/1→全部写入禁用→再次点toggle解除 |
| 4 | 20%应用→上箭头20.1%（R=0.402cm）→下箭头20%（R=0.40cm）→50%应用，上箭头禁用 |
| 5 | 完整缩放嵌套组，0.10cm→两次上箭头0.30→保护开启2/2→解除；每阶段保存检查几何/层级/标签 |
| 6 | strict嵌套子选区显示部分开启1/2，全部R角写入禁用，组与叶子不变 |
| 7 | 非圆角箭头，圆角数0，全部R角控件禁用 |

2026-10-07控件实测中第一个文稿在并发手动操作时发生组合整体移动，第5页另用全新样本重跑。独立核对明确记录两个保存文件来源：

```sh
.venv-native/bin/python test/native-host-fixture.py verify-controls /tmp/RadiusNativeV140Controls.pptx /tmp/RadiusNativeV140GroupControls.pptx
```

[控件验收记录](native-host-validation-v1.4-20261007.json)含保存状态及中间数值快照；第5页使用独立样本，其他页使用第一文稿。第一个文稿发生的整体移动不能算作插件的几何验证通过。

## 测试分层

| 层 | 文件 | 测什么 | 跑不跑 |
|---|---|---|---|
| **宿主验收** | `ppt-driver.js` + UI烟囱测试 | Mac LTSC Office.js 兼容性 | **不在 npm test 里**——在真实 PPT 跑"Driver 烟囱测试" |
| **纯算法** | `test-radius-core.js` | `computeLayout` / `valueToCm` / 业务规则（`shouldReject*` / `syncFixedValueIfLocked`） | npm test |
| **功能** | `test-features.js` | 业务函数（`writeRadius` / `applyLayout` / `syncLayoutChildrenR` / `readLockState` / `writeLockState` / `reapplyLock`）—— "模拟交互反馈" | npm test |
| **组合** | `test-driver-group.js` | 分层加载、展平、布局解组/重组及异常恢复 | npm test |
| **回归** | `test-regressions.js` | 实际dialog.js异步wiring、driver几何/tag读取、组结构恢复及HTTP错误 | npm test |

**v1.3 重整后**：功能测试用 `assertShape` 验最终状态，**不关心 driver 内部调了哪些方法**。

## 历史Office.js新功能测试流程（v1.3）

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
