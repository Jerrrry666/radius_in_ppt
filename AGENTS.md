# R 角调整 — Project Memory

> AI agent / 协作者第一次接手这个项目时必读。Mac LTSC Office.js 的行为和
> Windows / Microsoft 365 上有大量差异，这里把踩过的坑都整理好了。

## 1. 项目定位

macOS PowerPoint **task pane 加载项**，让用户用 **厘米** 或 **百分比** 设置圆角矩形的 R 角（圆角半径）。

- 支持多选
- 支持「锁定 R 角绝对值」（按厘米值保持；改变形状大小时自动按比例调整）
- 最近 5 次输入历史（本次 session 内存）
- 最终形态：`.app`（约 440 KB），双击启动本地静态 server + 把 manifest 注册到 PowerPoint 加载项目录

参考：iSlide 的「设计」tab 体验。

## 2. 架构

### 2.1 部署架构（PPT ↔ task pane ↔ server）

**纯 Office.js + task pane**：

```
┌─────────────┐    HTTP localhost:3000     ┌──────────────────┐
│  PowerPoint │ ◀─────────────────────────▶ │  Office Task Pane │
│  (LTSC Mac) │   Office.js bridge         │  - dialog.html    │
└─────────────┘   ShowTaskpane (侧边栏)    │  - dialog.js      │
        │                                  │  - dialog.css     │
        │ in-memory                        └──────────────────┘
        │ Adjustments.set(0, val)                  ▲
        │ getSelectedShapes()                      │
        ▼                                          │
┌─────────────┐                            ┌──────────────────┐
│  .pptx doc  │   (文件存盘由 PPT 自己处理) │  static server   │
│  + shape    │                            │  tools/serve.js  │
│    tags     │  ◀── lock 跟着 .pptx 走 ──  │  ~60 行          │
└─────────────┘                            └──────────────────┘
```

**lock 持久化用 `shape.tags`**（OOXML `<p:tagLst>` 段），跟着形状走，save .pptx 后跨设备/换机器都保留。

**history 纯内存**（本次 session 内用户主动应用过的 R 角值，关掉 PPT 任务窗格就清空）。

### 2.2 代码架构（两层 + UI 层）—— **2026-07-24 决定**

v1.2 暴露出来的问题：bug 卡在"是 PowerPoint 怪还是我逻辑怪"上永远分不清，
最后发现是 Office.js 的坑（per-shape load adjustments 不 work、load('tags') + sync 在 Mac LTSC 抛异常）
混在业务逻辑里，调试要靠 console.log 反推根因。

**新结构**：

```
┌──────────────────────────────────────────────────────────────┐
│  dialog.js (UI 层)                                           │
│  - 事件绑定、渲染、toast、debug log                           │
│  - 极薄，每个 handler 5-10 行                                 │
└──────────────────────────────────────────────────────────────┘
                              │ 调用
                              ▼
┌──────────────────────────────────────────────────────────────┐
│  src/lib/radius-core.js (实现层)                             │
│  - 所有 feature：writeRadius / applyLayout / applyPipette    │
│  - 业务判断：strict 拦截、lock 联动、padding 公式、关联父子   │
│  - 第一参数必是 driver（与 Office.js 解耦）                   │
│  - 零 Office.js 调用 → 可 100% 单元测试                       │
└──────────────────────────────────────────────────────────────┘
                              │ 调用
                              ▼
┌──────────────────────────────────────────────────────────────┐
│  src/lib/ppt-driver.js (交互层)  ← 新文件                    │
│  - createDriver(ctx) 工厂，导出 box/adj/load/sync/tagValue/   │
│    addTag/deleteTag 等小方法                                  │
│  - 每个方法假定"已 load + sync 过"，方法本身不调 load          │
│  - 零业务逻辑（不知道 strict / lock / layout 是什么）          │
│  - PPT 验过一次没 bug 后不再修改                              │
└──────────────────────────────────────────────────────────────┘
                              │ 调
                              ▼
┌──────────────────────────────────────────────────────────────┐
│  Office.js + PowerPoint (LTSC Mac)                           │
└──────────────────────────────────────────────────────────────┘
```

**关键约束**：
- **driver 不知道任何业务概念**——它只懂"形状"、"标签 key"、"调整值 0~1 分数"。
  不知道 `LOCK_TAG_KEY` / `LAYOUT_PARENT_TAG_KEY` 这些 key 字符串，也不管 strict
  标记是 `'1'` 还是 `'true'`，那是实现层的事。
- **实现层不 import Office.js**——所有形状读 / 写 / load / sync 都走 driver 方法。
  想测 feature 就 mock 一个 driver 对象喂进去。
- **UI 层是搬运工**——`onClick → 开 driver → 调 feature 函数 → 渲染结果`。

**driver API 形态**（10 个左右方法）：
```js
function createDriver(ctx) {
  return {
    // 加载 + 同步
    load: (proxy, fields) => proxy.load(fields),  // 传 'items/...' 路径
    sync: () => ctx.sync(),

    // 读（假定已 load + sync）
    box: (s) => ({ left, top, width, height }),
    adjFraction: (s) => s.adjustments.count > 0 ? s.adjustments.get(0).value : 0,
    isRoundRect: (s) => s.adjustments.count > 0,
    shapeId: (s) => s.id,

    // 写
    setBox: (s, box) => { ... },        // left/top/width/height
    setAdjFraction: (s, frac) => s.adjustments.set(0, frac),
    addTag: (s, key, value) => s.tags.add(key, value),
    deleteTag: (s, key) => s.tags.delete(key),

    // 读 tag（async，需要 sync 走完才能拿值）
    readTag: async (s, key) => { ... },
  };
}
```

### 2.3 迁移顺序（一个 feature 一个 feature 走）

每步都跑测试 + PPT 验一次再进下一步：

1. **driver 基础 + R 角调整**（最小可用）
2. **lock / strict 联动**（writeRadius 加 strict 拦截 + lock 同步 fixed value）
3. **layout mode**（applyLayout / syncLayoutChildrenR）
4. **pipette + history**
5. **删 dialog.js 旧逻辑**

### 2.4 单元测试策略

- `test/test-radius-core.js`（已存在，103 个）：纯算法层测试（computeLayout、computeLinkedSubR 等）
- `test/test-mock-harness.js`（已存在，70 个）：mock PowerPoint run 上下文，端到端测 feature
- `test/test-driver-integration.js`（**待新增**）：mock driver 对象 + 实现层 feature 函数的集成测试
  - 不再需要 mock 整个 PowerPoint.run，只 mock 10 个 driver 方法
  - 覆盖率从纯算法扩到全部 feature 路径

### 2.5 当前状态（2026-07-25 v1.3.5）

- [x] 实现层纯算法 + mock PowerPoint 集成（173 个测试全过）
- [x] **driver 层抽出完成**（v1.2.2）—— `src/lib/ppt-driver.js` 16 方法，createDriver 工厂
- [x] **writeRadius 迁移完成**（v1.2.2）—— `onApply` 走 driver + radius-core
- [x] **applyLayout 迁移完成**（v1.3.0）—— `applyLayoutToChildren` 160 行 → 50 行
- [x] **syncLayoutChildrenR 迁移完成**（v1.3.2）—— 36 行 → 11 行
- [x] **onToggleLock 关闭路径 fix**（v1.3.5）—— `locks[id]=null` + `strict[id]=false` 显式 delete
- [x] **writeRadius Infinity/NaN 防御**（v1.3.5）—— `!Number.isFinite(targetCm)` 在 clamp 前 reject
- [x] **112 个单测全过**（v1.3.5）—— 103 算法 + 70 mock harness + 109 driver 集成（含 v1.3.5 加的 2 个回归）
- [x] **driver 烟囱测试 14/14**（v1.3.4）—— 16 方法全 PPT 实测通过
- [x] **7 场景 end-to-end PPT 验证**（v1.3.5）—— onApply ×3 + strict / lock / pipette+history / layout
- [ ] **Step 3c** layout tag 读写迁移 + stale state 检测
- [ ] **Step 4** pipette + history 迁移到 driver + radius-core
- [ ] **Step 5** dialog.js UI 层重构（lockMonitor 重写、调试 log 清掉、dialog.js 缩到 ~500 行）

### 2.6 交互层 verified 状态（v1.3.5 决定）

**driver + radius-core + dialog.js wiring 正式 verified**。未来 feature（Step 3c/4/5）走：
- 单元测试覆盖（mock driver + radius-core）
- 代码 review（PR / diff 检查）
- **不再 PPT 实测**（除非逻辑有重大变更）

**verified 依据**：
- driver 16 方法：烟囱测试 14/14 ✅
- radius-core 8 个 driver 版函数：109 个集成测试 ✅
- dialog.js 关键 wiring（onApply / onToggleLock / onReapply / applyLayout / syncR）：7 场景 PPT 验证通过（除 feature bug #6/#7 外）

**回归保护**：
- 任何对 `writeRadius` / `writeLockState` / `applyLayout` / `syncLayoutChildrenR` 的修改必须先跑 `npm test`（112 个单测）
- 任何对 `onToggleLock` / `onApply` 的 dialog.js 逻辑修改需手动跑 7 场景 PPT 验证（一次性，verified 后转单测覆盖）



## 3. 目录结构

```
.
├── manifest.xml                      # Office Add-in 清单（指向 localhost:3000）
├── src/
│   └── dialog/                       # task pane UI（dialog 是历史命名）
│       ├── dialog.html
│       ├── dialog.js                 # 核心逻辑（~600 行）
│       └── dialog.css
├── app/MacOS/RadiusInPpt             # bash 启动器
├── tools/
│   ├── serve.js                      # ~60 行静态文件 server
│   ├── build-app.sh                  # 打包成 .app
│   ├── build-dmg.sh                  # 可选：打包成 .dmg
│   └── sign-and-notarize.sh          # 可选：代码签名 + 公证
├── assets/                           # 图标
├── dist/                             # build 输出（git ignore）
├── test.pptx                         # 测试用文件
├── AGENTS.md                         # ← 本文件
└── changelogs/
    └── 2026-07-23.md                 # v1.0 发布日志
```

## 4. Mac LTSC Office.js 行为差异（重点！）

> 这些是 Mac Office LTSC Standard for Mac 2021（build 16.111 / 26071325）上的实测行为。
> Microsoft 365 / Windows 上的行为可能不一样。**所有 API 行为以 Mac LTSC 为准**。

### 4.1 `Adjustments.value` 单位是 0~1，不是 0~50000

OOXML 里 `<a:gd name="adj" fmla="val X"/>` 的 X ∈ [0, 50000]（对应 0%~50% 短边）。
但 **Mac LTSC Office.js** `shape.adjustments.get(0).value` 返回的是 **0~1 的小数比例**（OOXML 值 ÷ 50000）。

```js
// 读：currentCm = adj.value * minSideCm
const adj = sh.adjustments.get(0).value;  // 0~1 fraction
const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
const currentCm = adj * minSideCm;

// 写：newAdj = (targetCm / minSideCm)
// 不能 Math.round（round(0.067) = 0，所有非整数都被截成 0）
const newAdj = (targetCm / minSideCm) * ADJ_SCALE;
sh.adjustments.set(0, newAdj);
```

### 4.2 `get(0)` 返回 ClientResult 代理，不是 ClientObject

`sh.adjustments.get(0)` 返回 **ClientResult 代理**（没有 `.load()` 方法），直接 `.value` 拿值。

```js
// ❌ 错（会报 "adjItem.load is not a function"）
const adj = sh.adjustments.get(0);
adj.load('value');
await ctx.sync();
const v = adj.value;

// ✅ 对
const v = sh.adjustments.get(0).value;
```

**注意**：`shape.tags.getItem('key').load('value')` 是 work 的——tag 不是 ClientResult 代理。两者 API 行为不同。

### 4.2.1 集合层 load 'items/adjustments' 不填 .value（v1.2.5 实测踩坑）

```js
// ❌ 错：集合层 load adjustments + sync 后，.value 还没填
shapes.load('items/id, items/adjustments');
await ctx.sync();
for (const sh of shapes.items) {
  const v = sh.adjustments.get(0).value;  // 抛「尚未加载结果对象的值」
}

// ❌ 错：批量 per-shape load 后单 sync（v1.2.6 实测）—— 第二个 shape 的 getItem 抛 GeneralException
shapes.load('items/id, items/adjustments');
await ctx.sync();
for (const sh of shapes.items) {
  if (sh.adjustments.count > 0) sh.adjustments.load('items/value');
}
await ctx.sync();  // ← 这个 sync 失败：selectedShapes.getItem(...) 抛 GeneralException

// ❌ 错：set + sync 之后立即读旧 proxy（v1.2.7 实测）—— proxy 不会自动 reload value
sh.adjustments.set(0, 0.5);
await ctx.sync();
const adjProxy = sh.adjustments.get(0);  // 早 get 的旧 proxy
const v = adjProxy.value;  // 0（不是 0.5）—— proxy 没自动 reload

// ❌ 错：v1.2.8 也错的——capture proxy 在最前面，set+sync+load+sync+读 proxy
//   Mac LTSC proxy 是 snapshot 风格，set 不会更新旧 proxy
const adjProxy = sh.adjustments.get(0);
sh.adjustments.set(0, 0.5);
await ctx.sync();
sh.adjustments.load('items/value');
await ctx.sync();
const v = adjProxy.value;  // 还是 0——旧 proxy 怎么 reload 都不会更新

// ✅ 对：per-shape get(0) + per-shape sync（v1.0 模式，慢但正确）
shapes.load('items/id, items/adjustments');
await ctx.sync();
for (const sh of shapes.items) {
  if (sh.adjustments.count === 0) continue;
  // 关键：get(0) 先存变量，再 per-shape sync，再读那个变量
  const adjResult = sh.adjustments.get(0);
  await ctx.sync();
  const v = adjResult.value;  // ✅
}

// ✅ 对：set 之后读，每次读之前 fresh get(0)（v1.2.9 实测）
// 关键洞察：Mac LTSC 上 `sh.adjustments.get(0)` 返回的 proxy 是 snapshot 风格，
// 后面 set/load/sync 都不会更新这个旧 proxy。**每次读都要 fresh get(0)**。
sh.adjustments.set(0, 0.5);
await ctx.sync();
sh.adjustments.load('items/value');
await ctx.sync();
const v = sh.adjustments.get(0).value;  // ✅ 0.5（fresh get，新 proxy 读最新 value）
```

v1.2.5 烟囱测试时 lock monitor 暴露的 bug——每 10ms 轮询 4 个 shape 全部失败。
- driver 层加 `driver.adjFraction(s)` 内部 try/catch 返回 0（defensive，不 throw）
- driver 层加 `driver.loadAdjValue(s)` 辅助（单 shape 情况 OK，批量会炸）
- v1.2.7 monitor 改回 v1.0 模式：per-shape get(0) + per-shape sync
- v1.2.8 烟囱测试 setAdjFraction 暴露「proxy 不自动 reload」：v1.2.7 修法不够，**set+sync 后必须再 load('items/value') + sync**

**绝对禁止**：
- ❌ 批量 per-shape `sh.adjustments.load('items/value')` 排队然后单 sync（v1.2.6 实测炸）
- ❌ `setAdjFraction` 之后用旧 proxy 读（同 sync 内 proxy 不会自动 reload value）—— 必须 reload + 再 sync

### 4.3 写 .pptx 持久化用 `shape.tags`（Mac LTSC 唯一 work 的方案）

`customProperties` 和 `customXmlParts` 在 **task pane 和 dialog 上下文都不可用**（Mac LTSC）：

- `customProperties` 在 dialog 直接 undefined
- `customXmlParts` 在 task pane / dialog 都不存在（Mac LTSC 16.111 实测）

**workaround**：`shape.tags`（PowerPointApi 1.10+，Mac LTSC 支持）：

```js
// 写
PowerPoint.run(async (ctx) => {
  const sh = ctx.presentation.getSelectedShapes().getItemAt(0);
  sh.tags.add("myKey", "myValue");
  await ctx.sync();
});

// 读
const tag = sh.tags.getItem("myKey");
tag.load("value");
await ctx.sync();
const v = tag.value;

// 删
sh.tags.delete("myKey");
```

**限制**：tag 是每个形状自己的，跨形状需要遍历。存的是 key-value 字符串对。

### 4.4 task pane 上下文里 `shapes.load` 不自动填 adjustments 子项

**这是 v1.0 唯一一个 hotfix 的坑**（commit `d6bba1a`），`refreshSelection` 里第一次读 adjustments.value 时漏了显式 load。

```js
// ❌ 错：task pane 里 .value 报 "结果对象的值尚未加载"
shapes.load('items/adjustments');
await ctx.sync();
const v = sh.adjustments.get(0).value;  // ❌ 报错

// ✅ 对：显式 load 子项（每个 roundRect 都要做）
shapes.load('items/adjustments');
await ctx.sync();
for (const sh of shapes.items) {
  if (sh.adjustments.count > 0) {
    sh.adjustments.load('items/value');  // ← 显式 load，task pane 必加
    await ctx.sync();
    const v = sh.adjustments.get(0).value;  // ✅
  }
}
```

（dialog 上下文里这步可能不必要；task pane 必须显式 load。）

**写时不需要**：只有读 `.value` 时才需要这个显式 load；写时 `sh.adjustments.set(0, newVal)` 不需要。

### 4.4.1 v1.2 新坑：per-shape load `adjustments` 不 work，必须 collection-level load（2026-07-24 实测）

**Mac LTSC task pane 上下文**：`shape.adjustments.count` 在 per-shape load 之后**永远 = 0**。必须在 collection 级别（`sel.load('items/adjustments')` 或 `slide.load('shapes/items/adjustments, ...')`）才 work。

```js
// ❌ 错：v1.2 applyLayoutToChildren 用了这个，.count 永远 0
parentSh.load('left, top, width, height, adjustments');
await ctx.sync();
parentSh.adjustments.count  // 永远 0，即使父是 Rounded Rectangle

// ✅ 对：collection-level load 一次性 load 所有需要的字段
const activeSlide = ctx.presentation.getSelectedSlides().getItemAt(0);
activeSlide.load('shapes/items/id, shapes/items/left, shapes/items/width, shapes/items/height, shapes/items/adjustments');
await ctx.sync();
const parentSh = idToShape.get(parentId);
parentSh.adjustments.count  // ✅ 1 (圆角矩形) / 0 (普通矩形)
```

**v1.0 / v1.1 monitor 用的都是 collection-level load**，所以 work。v1.2 applyLayout 用了 per-shape load，导致 R 角联动算出 subR = 0（用户状态卡能看到父 R 角 1.62cm，但 apply 时读不到）。

**v1.2 commit `fix/v1.2-load-adjustments-collection` 已修**。

**这是 v1.0 唯一一个 hotfix 的坑**（commit `d6bba1a`），`refreshSelection` 里第一次读 adjustments.value 时漏了显式 load。

```js
// ❌ 错：task pane 里 .value 报 "结果对象的值尚未加载"
shapes.load('items/adjustments');
await ctx.sync();
const v = sh.adjustments.get(0).value;  // ❌ 报错

// ✅ 对：显式 load 子项（每个 roundRect 都要做）
shapes.load('items/adjustments');
await ctx.sync();
for (const sh of shapes.items) {
  if (sh.adjustments.count > 0) {
    sh.adjustments.load('items/value');  // ← 显式 load
    await ctx.sync();
    const v = sh.adjustments.get(0).value;  // ✅
  }
}
```

（dialog 上下文里这步可能不必要；task pane 必须显式 load。）

**写时不需要**：只有读 `.value` 时才需要这个显式 load；写时 `sh.adjustments.set(0, newVal)` 不需要。

### 4.5 `Adjustments.count` 是 primitive，能直接用

`sh.adjustments.count` 在 Mac LTSC task pane 里是 **number**（不是 ClientObject），不需要 load：

```js
const isRoundRect = sh.adjustments.count > 0;
```

### 4.5.1 不要在 `writeRadiusToShape` 里 `ctxShape.load('tags')` + `await ctx.sync()`（2026-07-24 实测踩坑）

v1.2 `writeRadiusToShape` 一开始加了 `ctxShape.load('tags'); await ctx.sync();` 想「保险」预 load tags，
结果在 Mac LTSC task pane 抛未捕获异常，被外层 catch 吞了返回 `reason='exception'`——所有 R 角写入
全部静默失败（位置/尺寸写成功，R 角写不进去）。用户报告「批量修改 R 角时无法识别是圆角矩形」。

**根因**：
- 所有 4 个 caller（applyLayoutToChildren / syncLayoutChildrenR / onApply / applyPipetteToSelection）
  都已经在 `PowerPoint.run` 外层 `sel.load('items/.../tags')` 或 `slide.load('shapes/items/.../tags')`，
  collection-level 已经 load 过 tags
- v1.0 working 代码（lock monitor、onApply、applyPipette）**从未**做这个额外的 per-shape `load('tags')`，
  直接 `sh.tags.getItem(KEY).load('value')` + `await ctx.sync()` 就 work
- 额外的 `ctxShape.load('tags')` + `await ctx.sync()` 在 Mac LTSC 上可能跟 collection-level load 冲突
  或 load 路径不被支持，抛 `GeneralException` / `InvalidArgument`

**正确写法**：
```js
async function writeRadiusToShape(ctxShape, targetCm, opts) {
  try {
    // 1. 读 lock + strict —— 直接 getItem + load value，不做 ctxShape.load('tags')
    let isLocked = false, isStrict = false;
    try {
      const lockTag = ctxShape.tags.getItem(LOCK_TAG_KEY);
      lockTag.load('value');
      await ctx.sync();
      if (lockTag.value && parseFloat(lockTag.value) > 0) isLocked = true;
    } catch (_) {}
    try {
      const strictTag = ctxShape.tags.getItem(LOCK_STRICT_TAG_KEY);
      strictTag.load('value');
      await ctx.sync();
      if (strictTag.value === '1') isStrict = true;
    } catch (_) {}
    ...
```

**绝对禁止**：
- ❌ `ctxShape.load('tags'); await ctx.sync();` 在 `writeRadiusToShape` 函数体内
- ❌ 任何在 caller 已经 collection-level load 过 tags 的情况下，再做 per-shape `load('tags')` 的双重 load

**调试技巧**：如果以后又出现 `reason='exception'` 但 caller 只 log `reason`，**第一时间**检查
`writeRadiusToShape` 的 catch 块有没有把 `e.message` 主动 `console.log` 出来——大概率就是某个被吞的
Office.js 异常。加 `console.log('[writeRadius] EXCEPTION:', msg, '| stack:', stack)` 立即显形。

### 4.5.2 driver.box vs driver.size — 契约要清楚（2026-07-24 实测踩坑）

driver 有两个返回 size 相关的读方法：

| 方法 | 返回 | caller 必须 load 的字段 |
| --- | --- | --- |
| `driver.size(s)` | `{width, height}` | `s.width, s.height` |
| `driver.box(s)`  | `{left, top, width, height}` | `s.left, s.top, s.width, s.height` |

**踩坑历史**：v1.2.2 第一次 PPT 测，`onApply` load 的是
`items/id, items/width, items/height, items/adjustments, items/tags`
**没 load `items/left, items/top`**。`writeRadius` 调 `driver.box(shape)` 访问 `s.left`
直接抛 `"属性"left"不可用。读取属性的值之前，请先对包含对象调用 load 方法"`，被 catch 吞了返回
`reason='exception'`。

**修法**：
- 业务函数按需选方法：写 R 角只需要 `minSideCm` → 用 `driver.size`（只要 width/height）
- layout apply 需要算子位置 → 用 `driver.box`（要 4 个字段），caller load 时记得加 `items/left, items/top`
- driver 注释里写清楚每个方法的 load 契约

**绝对禁止**：
- ❌ 业务函数无脑调 `driver.box` 然后在 caller 漏 load left/top → 永远走不到正常路径
- ❌ 业务函数调 `driver.box` 但 caller 只 load 了部分字段

**测试覆盖**：`test/test-driver-integration.js` 有专门的回归测试
「writeRadius 不需要 left/top 被 load」——故意把 left/top 改成访问就抛异常的 getter，
确认 writeRadius 走 driver.size 不读 left/top。### 4.6 选区 API

| API | Mac LTSC task pane |
| --- | --- |
| `ctx.presentation.getSelectedShapes()` | ✅ 工作（PowerPointApi 1.6+） |
| `sh.width` / `sh.height` | ✅ 工作，单位是 pt |
| `sh.id` / `sh.name` | ✅ 工作 |
| `Office.context.document.addHandlerAsync(DocumentSelectionChanged, ...)` | ✅ 工作（Common API） |
| shape change 事件（`ShapeResized` 等） | ❌ **不存在**——必须用 setInterval 轮询 |

### 4.7 **没有 shape-level change 事件**

Office.js PowerPoint **不提供** `ShapeResized` / `ShapeMoved` / `ShapePropertyChanged`。必须用 `setInterval` 轮询检测拖动完成。

**当前实现**：10ms 一次轮询，4 次连续无变化（≈40ms 稳定）= 视为用户松手 → 反算 adj 写回。拖拽中尺寸在变 → 跳过 apply，避免和拖动手感冲突。

## 5. 部署 / 路径问题

### 5.1 manifest 路径会被 PowerPoint 重启清空

`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` 在 PowerPoint 退出（Cmd+Q）时可能被回收。

**所以 `.app` 启动器必须每次都**：

```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
cp -f dist/RadiusInPpt.app/Contents/Resources/manifest.xml \
      ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml
```

**用户必须 Cmd+Q 完全退出 PowerPoint**，然后重新打开，新 manifest 才生效。只关窗口不够（macOS 不会真的退 Office）。

### 5.2 启动器要主动找 node

macOS launchd 的精简 PATH 不一定有 `/opt/homebrew/bin/node`，启动器（`app/MacOS/RadiusInPpt`）要主动找：

```bash
for p in /opt/homebrew/bin/node /usr/local/bin/node /opt/local/bin/node; do
  if [ -x "$p" ]; then NODE="$p"; break; fi
done
```

### 5.3 localhost 用 HTTP

Office Add-in 允许 `http://localhost` 走 HTTP（**不**需要 HTTPS / 证书）。manifest 里所有 URL 都是 `http://localhost:3000`。

### 5.4 ⚠️ iCloud Documents 下的 dist 重建

如果项目放在 `~/Documents/`（iCloud Drive 同步盘），**别外层 `mavis-trash dist`**：
- iCloud 会把 dist 移到自己的 Trash，30 天后才真正删除
- 期间 PowerPoint / Spotlight / 其他 macOS 服务还在引用旧路径
- 系统会反复弹 "无法完成此操作，因为需要下载'dist'" 让用户恢复

**正确做法**：直接 `bash tools/build-app.sh` 覆盖式重建（脚本内部 `find $DIST -mindepth 1 -maxdepth 1 -exec rm -rf {} +` 已经清掉 `.app` / `.dmg` / `dmg-staging` / `AppIcon.iconset`），不要在外层 trash dist。

如果已经被卡住、对话框反复弹：
1. 点"好"消掉
2. Cmd+Q PowerPoint
3. `touch /Users/ma/Documents/minimax/radius_in_ppt/dist` 强制 iCloud 重新拉本地
4. 重启 .app

## 6. Git 推送（带 token 走 HTTPS）

Mac 上 token 经常被 git 拒（认证对话框），用一次性 credential helper：

```bash
GH_TOKEN="ghp_xxxxxxxxxxxx"
git -c credential.helper="!f() { echo username=x-access-token; echo password=$GH_TOKEN; }; f" \
    push origin minimax
```

**commit message 不要用中文标点**——bash 解析会炸。用 ASCII 标点。

## 7. 关键 commit 历史（看时间序）

| commit | 说明 |
| --- | --- |
| `99487d5` | 最初 PowerPoint R 角调整加载项 v1.0 |
| `35a64df` | 打包成 macOS .app |
| `2a30609` | 启动脚本主动找 node |
| `d70f2df` | 回归 Office Add-in 路线（wef 路径 + bash 启动器） |
| `07a1ce4` | 改用 server 端解析 .pptx（绕开浏览器 JSZip） |
| `eb5b724` | **重构：纯 Office.js，删 server 端 PPTX 处理** |
| `3a92e17` | **ADJ_SCALE 改成 1**（Mac LTSC 返回 0~1 不是 0~50000） |
| `918934d` | 改用 OOXML CustomXmlPart 存锁（后来发现 Mac LTSC 也不 work） |
| `e4629d8` | **改 task pane**（从 dialog 改成侧边栏） |
| `76f9bd6` | **改用 shape.tags 存锁**（Mac LTSC 唯一能 work 的持久化） |
| `b19172b` | history 加文件扫描（后来删了，file scan 太脆弱） |
| `14d1f6c` | history 简化为纯内存 |
| `v1.0` | **v1.0 正式版**：删调试代码、删 shared/ 和 commands/ 目录、代码重整 |

## 8. 已知限制 / 未来工作

- [x] **锁定 R 角 cross-machine**：✅ 改用 shape.tags
- [x] **lock 之后改变形状大小**：✅ setInterval 10ms 轮询 + 4 次稳定检测
- [x] **打包 .dmg**：`tools/build-dmg.sh` 已实现
- [ ] **代码签名 + 公证**：`tools/sign-and-notarize.sh` 已写好，待用户有 Apple Developer 账号时启用
- [ ] **history 跨会话**：当前只活内存。如果需要跨 session 保留，得用 shape.tags 在一个隐藏形状上挂 JSON
- [ ] **多选混合 UI**（圆角矩形 + 普通矩形）：当前标记非圆角 + 跳过 apply

## 9. 调试技巧

### 9.1 验证 lock 真的跟文件走

```
1. 选个圆角矩形，点「锁定 R 角」
2. Cmd + S 保存 .pptx
3. Cmd + Q 完全退 PPT
4. 重新打开同一个 .pptx
5. 选中刚才那个圆角矩形
6. 状态卡「已锁定」应该显示 1
```

### 9.2 重置 .app 状态

```bash
pkill -f "tools/serve.js"
bash tools/build-app.sh
node tools/serve.js > /tmp/serve.log 2>&1 &
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
cp -f dist/RadiusInPpt.app/Contents/Resources/manifest.xml \
      ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml
# 然后用户 Cmd+Q + 重开 PowerPoint
```

### 9.3 看 server 日志

```bash
tail -f /tmp/serve.log
```

## 10. PowerPoint 版本

- 目标：**Office LTSC Standard for Mac 2021**（build 16.111 / 26071325）
- API 范围：PowerPointApi 1.1 ~ 1.10
- 已验证可用：`getSelectedShapes`（1.6）、`Adjustments.get/set`（1.10）、`shape.tags`（1.10）、`customXmlParts`（Common API，**不可用**）、`customProperties`（1.7，**不可用**）
- Microsoft 365 用户理论上也能跑，但有些行为可能跟 LTSC 不一样

### 4.6 v1.2 防误触设计原则：最高优先级 + 任何路径都不能跳过（2026-07-24 实测决定）

**原则**：strict tag = "1" 的形状，R 角写入**永远**被拦截。任何 R 角写入路径都不能跳过（layout apply / 样式刷 / 联动 hook 都不行），必须由用户手动关闭防误触。

**v1.1 行为**：
- onApply：选区里有任何 strict → 全部拒绝
- 样式刷：选区里有任何 strict → 全部拒绝

**v1.2 新增**：
- `writeRadiusToShape(ctxShape, targetCm, opts)` 统一函数：
  - 写 R 角前**必查** strict tag
  - 命中 → 返回 `{ok: false, reason: 'strict'}`，不写 PPT
  - **不允许 skipStrict 选项**（已经移除，函数签名里没有这个字段）
- applyLayoutToChildren：进入 PowerPoint.run **之前**先扫选区内存（`selectedShapes[i].strictLocked`），有 strict → 拒绝整个 apply（不写位置/尺寸，不写 R 角，不写 tag）
  - 写 R 角时（writeRadiusToShape 内部）作为**第二道防线**再检查（实时读 tag，防 race）
- renderLayoutPanel：「进入组合时」判断 — 选区里有 strict → 「建布局」按钮禁用 + hint 提示「🔒 N 个启用了防误触，请先关闭」
- renderLayoutSetupList：strict 形状行**红框 + 🔒 标记** + title 提示
- syncLayoutChildrenR：strict 永远拦截（命中就跳过，layout 联动时 R 角不更新该子）

**绝对禁止**：
- 不要在 writeRadiusToShape 加 `skipStrict` 选项
- 不要在 applyLayoutToChildren / applyPipetteToSelection / syncLayoutChildrenR 加 bypass 逻辑
- 防误触是用户主动选择开启的，程序不能"贴心地"自动覆盖

**两道防线模式**（applyLayoutToChildren 是范例）：
1. 内存层：PowerPoint.run 之前检查 `selectedShapes[i].strictLocked`
2. PPT 层：PowerPoint.run 内 writeRadiusToShape 检查 tag（防 race / 防用户中途切状态）
