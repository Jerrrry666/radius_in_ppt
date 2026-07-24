# R 角调整 — 测试

## 跑测试

```bash
# 跑所有测试
npm test

# 单独跑
node test/test-radius-core.js
node test/test-mock-harness.js
node test/test-driver-integration.js
```

## 测试结构

```
test/
├── README.md
├── test-radius-core.js         # 纯算法测试（103 个）
├── test-mock-harness.js        # Mock PowerPoint 集成测试（70 个）
└── test-driver-integration.js  # driver + writeRadius(driver) 集成测试（36 个）
```

## 覆盖范围

### 1. 纯算法层（test-radius-core.js，103 个）

| 类别 | 数量 | 例子 |
| --- | --- | --- |
| `computeLayout` 布局 math | 24 | 2×2, 1×3, 3×2, 不可行, 1×1, 非零起点 |
| 单位换算 | 7 | cm↔cm, %↔cm |
| 联动公式 | 11 | subtract, same, off 各种边界 |
| clamp + adj | 14 | clamp 到短边一半, adj 转换, computeFinalRadius 端到端 |
| strict/lock 行为 | 13 | 单 shape, onApply, layout apply, lock 同步 |
| 集成场景 | 22 | 2×2 layout, 1×3 layout, 拒绝场景, lock 同步, 用户样例 |

### 2. Mock PowerPoint 集成（test-mock-harness.js，70 个）

模拟 PowerPoint.js 的 shape proxy，在 mock 环境里跑完整的 `writeRadiusToShape` + `applyLayoutToChildren` 流程。

| 类别 | 数量 | 验证 |
| --- | --- | --- |
| `writeRadiusToShapePure` 基础 | 8 | 写入、clamp、adj 转换 |
| strict 拦截 | 4 | strict 永远拦截，adj/lock tag 都不动 |
| locked 同步 fixed value | 4 | lock tag 同步到新 R 角 |
| `applyLayoutPure` 端到端 | 12 | 位置/尺寸/R 角/tag 全部写对 |
| strict 拒绝整个 apply | 6 | 含 strict 子 → 整个 apply 拒绝，**位置/尺寸/tag 都不动** |
| locked 子同步 | 4 | lock tag 同步到新 R 角 |
| 边界 | 12 | 父不是 roundRect、子不足、跨 slide、linkRMode=off |

### 3. driver 集成（test-driver-integration.js，36 个）— 2026-07-24 新增

mock `createDriver(ctx)` 注入到 `radius-core.writeRadius(driver, ...)`，验证 driver 协议 + 业务逻辑整体。

| 类别 | 数量 | 验证 |
| --- | --- | --- |
| 基础写 R 角 | 1 | ok/newCm/wasLocked 都对，adjFraction 转换正确 |
| strict 永远拦截 | 1 | ok=false, reason=strict, adjFraction 不变 |
| 普通矩形 | 1 | reason=not-roundRect |
| 0 尺寸 | 1 | reason=no-size |
| clamp | 1 | 目标 > 短边一半 → clamp |
| negative targetCm | 1 | newCm=0 |
| locked 同步 fixed value | 4 | ok + wasLocked + addTag radiusLock_v1 + _tags 更新 |
| locked + strict | 4 | strict 优先，wasLocked=true（读了但没写） |
| layoutParentId | 2 | ok + addTag layoutChild_v1 |
| 批量 5 个 | 6 | 各种 reason 分布正确，readTag 调用次数对 |
| driver 异常 | 3 | reason=exception，error 含异常 message |
| API 一致性 | 1 | createDriver 返回 14 个方法，方法名跟 mock 一致 |

### 3. 真实 PPT 测试（**需要人工**）

mock harness 测的是**逻辑流程**。**真实 PowerPoint.js API 行为**只能靠手工验证：

- 跑 .app，在真实 PPT 里建布局
- 看 task pane 底部「🔧 调试日志（点击展开）」
- 复制日志分析是否符合预期

**mock harness 不能测的**：
- Mac LTSC 的 PowerPoint.js bug（per-shape load adjustments 不 work、get(0) ClientResult 等）
- shape.tags 真实持久化（关 PPT → 重开是否还在）
- 跨 page 隔离是否真的生效
- lock monitor 是否真的反算 / 不同步

## 测试覆盖原则

按"**防误触 = 最高优先级**"原则（见 AGENTS.md 4.6）：

- ✅ 任何 R 角写入路径都不能 skip strict
- ✅ 任何 R 角写入路径都要检查 lock 状态
- ✅ locked 形状被 R 角写入 → 同步 fixed value
- ✅ 含 strict 的 layout apply → 整个拒绝（包括位置/尺寸）

## 添加新测试

写测试时：
- 纯数学用 `test-radius-core.js`
- 涉及 R 角写入流程（mock 整 shape 协议）用 `test-mock-harness.js`
- 涉及 driver 协议（mock driver）用 `test-driver-integration.js`
- 跑 `npm test` 验证全部通过

## driver 协议（2026-07-24 新增）

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
  box(s),                   // { left, top, width, height }
  isRoundRect(s),           // s.adjustments.count > 0
  adjFraction(s),           // s.adjustments.get(0).value (0~1)

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

`writeRadiusToShapePure` 接受一个普通对象 `shape`：

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

测试用 `makeMockShape({...})` 创建 mock shape。
