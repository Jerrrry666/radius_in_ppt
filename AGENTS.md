# R 角调整 — Project Memory

> AI agent / 协作者第一次接手这个项目时必读。
> 精简版 rules：详细变更 → [LOG.md](./LOG.md) / [changelogs/](./changelogs/) / [plans/](./plans/)

## 0. 项目定位

macOS PowerPoint **task pane 加载项**，让用户用 **厘米** 或 **百分比** 设置圆角矩形的 R 角（圆角半径）。支持多选 / 锁定 R 角绝对值 / 防误触 / 预设库 / 样式刷 / 5 次历史 / v1.2 布局模式。

- **目标平台**：Office LTSC Standard for Mac 2021（build 16.111 / 26071325）
- **API 范围**：PowerPointApi 1.1 ~ 1.10
- **最终形态**：`.app`（约 440 KB），双击启动本地静态 server + 把 manifest 注册到 PowerPoint 加载项目录

---

## 1. 工作铁律 (RULES)

> 这是硬约束，违反任何一条都会出 bug。

### 1.0 ⭐ Git commit / push 规则（用户授权）

**build 可以自动跑（本地操作，无副作用）；commit / push / build-and-deploy.sh 仍需用户指令。**

按操作风险分级：

| 操作 | AI 能否自主 | 理由 |
| --- | --- | --- |
| `npm test` / 单测 | ✅ | 纯本地、零副作用 |
| `bash tools/build-app.sh` / `build-dmg.sh` | ✅ | 纯本地、覆盖式重建 dist/，无远程操作 |
| `git add` / `git commit` / `git push` | ❌ 需用户明确指令 | 改写本地/远程历史 |
| `bash tools/build-and-deploy.sh` | ❌ 需用户明确指令 | **包含** commit + push + 移动 tag，复合操作 |
| Push git tag / 移动 tag | ❌ 需用户明确指令 | tag 改写公共历史 |

**默认行为**（改完代码后）：
1. 跑 `npm test` 验证
2. 写 changelog（v1.x.md）
3. **自动跑** `bash tools/build-app.sh`（让用户可以马上 PPT 实测）
4. 停下来，等用户指令：
   - "commit" → `git add . && git commit`（v1.x changelog / version bump 一并 commit）
   - "push" → `git push origin main && git push origin vX.Y`
   - "commit 并 push" / "commit + push" / "commit and push" → 两步连做
   - "帮我部署" / "deploy" / "打包部署" → 一次性跑 `build-and-deploy.sh`

**v1.2.14 教训**：用户反馈"位置错了" → 已经 commit + push 完 → 撤回要 force-push 改写公共历史（麻烦 + 风险）。**早问一句"现在要 commit/push 吗"省事**。build 不算 commit/push（v1.3.1 修订：build 是本地操作，可以自动跑让用户马上实测）。

### 1.1 ⭐ 版本号规则

格式 `v{MAJOR}.{MINOR}.{PATCH}`：

| 操作 | AI 能否自主 | 例子 |
| --- | --- | --- |
| Bump PATCH（第 3 位） | ✅ 可以 | v1.3.0 → v1.3.1 |
| Bump MINOR（第 2 位） | ❌ 需用户显式指令 | v1.2.x → v1.3.x |
| Bump MAJOR（第 1 位） | ❌ 需用户显式指令 | v1.x.x → v2.x.x |
| Push git tag | ✅ 可以（不动 manifest 版本号）| git push origin v1.2 |

**触发判断时机**：
- 改完代码要 bump → 看这次改动的"语义"是什么
  - 修 bug / 边界防御 / 内部重构 → PATCH
  - 加新 feature / 大改架构 / 破坏性变更 → MINOR（问用户）
  - 跨大版本不兼容 / 整体重写 → MAJOR（问用户）
- 拿不准时 → 问用户

**PATCH 是发布单元，不是调试次数**：
- 同一个用户问题在尚未 commit / tag / 发布期间，无论经历多少次实机尝试，都只占
  一个目标 PATCH；例如本轮所有 Group 验收修复统一为 `v1.3.1`。
- 只有当前 PATCH 已验收/发布，之后又出现新的独立问题，才递增到下一个 PATCH。
- 需要强制 PowerPoint 读取新前端时，不得靠递增正式版本号；关闭并重开 task pane，
  或使用独立的非版本 cache-buster。

### 1.2 架构铁律

三层架构（v1.2 决定，**禁止回退**）：

```
dialog.js (UI 层)         事件绑定 / 渲染 / toast / debug log
       │
       ▼
radius-core.js (实现层)   8 个 driver 版函数（纯算法 + 业务判断）
       │                   零 Office.js 调用
       ▼
ppt-driver.js (交互层)    16 个方法（Office.js 薄封装）
       │                   零业务逻辑
       ▼
Office.js + PowerPoint
```

**硬约束**：
- ❌ driver 不知道任何业务概念（不认 `LOCK_TAG_KEY` / `LAYOUT_PARENT_TAG_KEY`，不知 strict 是什么）
- ❌ radius-core 不 import Office.js
- ❌ dialog.js 直接调 Office.js（必须走 driver + radius-core）
- ❌ 在 radius-core 函数里 catch 异常后**只 log reason 不 log error message**（v1.2.1 教训：吞了 `GeneralException` 找不到根因）

### 1.3 测试铁律

| 改动目标 | 必须先做的事 |
| --- | --- |
| `writeRadius` / `writeLockState` / `applyLayout` / `syncLayoutChildrenR` 等 radius-core 函数 | `npm test`（112 个单测）全过 |
| `onToggleLock` / `onApply` 等 dialog.js 关键 wiring | 一次性 PPT 验证 7 场景，**之后转单测覆盖** |
| 新 feature | 写完单测（mock driver + radius-core）+ 代码 review，**不 PPT 实测** |
| `ppt-driver.js` 改动 | 烟囱测试 14/14 + 单测 |

**单测位置**：
- `test/test-radius-core.js`（103 个）—— 纯算法
- `test/test-mock-harness.js`（70 个）—— mock PowerPoint.run 上下文
- `test/test-driver-integration.js`（109 个）—— mock driver + radius-core 集成

### 1.4 防误触铁律

strict tag = "1" 的形状，**任何 R 角写入路径都不能跳过**。两道防线：
1. 内存层（PowerPoint.run 之前检查 `selectedShapes[i].strictLocked`）
2. PPT 层（PowerPoint.run 内 writeRadius 实时读 tag 检查）

**绝对禁止**：
- ❌ 给 `writeRadius` / `writeRadiusToShape` 加 `skipStrict` 选项
- ❌ 在 `applyLayout` / `applyPipette` / `syncLayoutChildrenR` 加 bypass 逻辑
- ❌ 让"防误触是用户主动选择开启"被任何代码"贴心地"自动覆盖

### 1.5 通用禁止（高频踩坑）

- ❌ `ctxShape.load('tags'); await ctx.sync();`（在 `writeRadius` 函数体内，已 collection-level load 过 tags）
- ❌ 批量 per-shape `sh.adjustments.load('items/value')` 排队后单 sync（v1.2.6 实测炸）
- ❌ `setAdjFraction` 后用旧 proxy 读 value（Mac LTSC proxy 是 snapshot 风格）
- ❌ per-shape `sh.load('..., adjustments')` —— `.count` 永远 = 0
- ❌ 外层 `mavis-trash ~/Documents/<project>/dist`（iCloud 卡住，30 天后才真删）
- ❌ commit message 用中文标点（bash 解析会炸）
- ❌ catch 块只 log `r.reason` 不 log `e.message`（v1.2.1 教训：根因被吞）

---

## 2. Mac LTSC Office.js 坑 (GOTCHAS)

> 这些是 Mac Office LTSC 16.111 实测行为。Microsoft 365 / Windows 上不一定一样。
> **所有 API 行为以 Mac LTSC 为准**。

### 2.1 `Adjustments.value` 单位是 0~1

OOXML 里 `0~50000`（0~50% 短边），但 **Mac LTSC** Office.js 返回 **0~1 小数**：

```js
// 读: currentCm = adj.value * minSideCm
const adj = sh.adjustments.get(0).value;  // 0~1 fraction
const minSideCm = Math.min(sh.width, sh.height) / PT_PER_CM;
const currentCm = adj * minSideCm;

// 写: ADJ_SCALE = 1（不是 50000）
// 不要 Math.round（round(0.067) = 0）
const newAdj = (targetCm / minSideCm) * 1;
sh.adjustments.set(0, newAdj);
```

### 2.2 `get(0)` 是 ClientResult 代理

- ✅ `sh.adjustments.get(0).value` 直接读
- ❌ `sh.adjustments.get(0).load('value')` —— 报 `load is not a function`
- ⚠️ `sh.tags.getItem('key').load('value')` 是 work 的（tag 不是 ClientResult 代理，行为不同）

### 2.3 load 模式：必须 collection-level

- ✅ `sel.load('items/id, items/adjustments')` 或 `slide.load('shapes/items/..., items/adjustments')`
- ❌ `sh.load('left, top, width, height, adjustments')` —— `.count` 永远 = 0
- 读 value 还需要**显式** `sh.adjustments.load('items/value')` + sync（task pane 必加；dialog 可能不必要）

### 2.4 set+read 必须 fresh get(0) AFTER sync

Mac LTSC proxy 是 **snapshot 风格**，set 不会更新旧 proxy：

```js
// ✅ 对：set → sync → fresh get(0) → 读 value
sh.adjustments.set(0, 0.5);
await ctx.sync();
sh.adjustments.load('items/value');
await ctx.sync();
const v = sh.adjustments.get(0).value;  // 0.5（每次 fresh get）

// ❌ 错：set → sync → 读旧 proxy
const adjProxy = sh.adjustments.get(0);  // 早 get 的
sh.adjustments.set(0, 0.5);
await ctx.sync();
const v = adjProxy.value;  // 还是 0
```

**最稳模式**：set + sync 后**跨 PowerPoint.run 兜底**（同 run 偶尔不稳定，v1.3.1 实测）

### 2.5 持久化用 `shape.tags`（Mac LTSC 唯一 work）

- ❌ `customProperties` / `customXmlParts` 在 task pane 和 dialog 都不可用
- ✅ `shape.tags`（PowerPointApi 1.10+，OOXML `<p:tagLst>` 段，跟 .pptx 走）

```js
// 写
sh.tags.add("key", "value");
// 读
const tag = sh.tags.getItem("key");
tag.load("value");
await ctx.sync();
const v = tag.value;
// 删
sh.tags.delete("key");
```

限制：tag 是每个形状自己的，跨形状要遍历；存的是 string key-value。

### 2.6 `Adjustments.count` 是 primitive

直接用 `sh.adjustments.count > 0`，不需要 load（task pane 也 work）。

### 2.7 driver.box vs driver.size 契约

| 方法 | 返回 | caller 必须 load |
| --- | --- | --- |
| `driver.size(s)` | `{width, height}` | `s.width, s.height` |
| `driver.box(s)` | `{left, top, width, height}` | `s.left, s.top, s.width, s.height` |

业务函数按需选方法（写 R 角用 `size` 只要短边；layout apply 用 `box` 算子位置）。driver 注释里写清楚 load 契约。

### 2.8 没有 shape-level change 事件

Office.js PowerPoint **不提供** `ShapeResized` / `ShapeMoved` / `ShapePropertyChanged`。必须 `setInterval` 轮询：

```js
// 10ms 一次，4 次连续无变化（≈40ms 稳定）= 视为用户松手
// 拖拽中尺寸在变 → 跳过 apply，避免和拖动手感冲突
```

`DocumentSelectionChanged` 事件 work，可监听选区变化。

### 2.9 选区 API

| API | Mac LTSC task pane |
| --- | --- |
| `ctx.presentation.getSelectedShapes()` | ✅（PowerPointApi 1.6+）|
| `sh.width` / `sh.height`（单位 pt）| ✅ |
| `sh.id` / `sh.name` | ✅ |
| `DocumentSelectionChanged` | ✅（Common API）|
| `ShapeResized` / `ShapeMoved` | ❌ 不存在（见 2.8）|

### 2.10 GroupShape 处理（v1.3.1）

**坑**：`getSelectedShapes()` 选组合时**不会**自动展平 group，返回的就是 group proxy（1 个，不是 N 个子）。`group.adjustments.count = 0`（group 本身不是 roundRect），被 `isRoundRect` 过滤掉 → **整个 group 在 UI 上"看不见"**。

**修法**：driver 层加 `flattenSelected` 递归展平 → 业务层拿到的永远是叶子 shape，无感。

```js
// ppt-driver.js v1.3.1 结构方法
driver.isGroup(s)               // 判定 group（只信 s.type，不兜底 s.group）
driver.groupShapes(group)       // 拿子 shape 数组（同步，不调 sync）
driver.flattenSelected(sel)     // 递归展平 → 叶子 shape 数组（普通数组，不是 Office.js collection）
await driver.loadShapeTree(collection, fields) // 分层加载真实 group 子集合 → 叶子数组
driver.hasTopLevelGroup(sel)     // 顶层选区是否含 group（区分整组缩放 vs 叶子编辑）
driver.shapeLevel(s)             // 0=顶层，1+=group 内；caller 先 load level
driver.parentGroupOf(s)          // 最近一次 tree 遍历记录的直接父 group（不碰 Shape.parentGroup）
driver.topGroupOf(s)             // 沿安全索引取最外层 group
```

**关键事实**：
- `PowerPoint.ShapeType.group` 实际是字符串 `'Group'`（不是数字 enum）
- `PowerPointApi 1.8+` 才有 `ShapeGroup`（`sh.group.shapes.items` 拿子）
- group 可嵌套（group 里再 group），要 depth-first 递归
- 防死循环：用 `Set` 记已访问 shape id

**⚠️ 两个禁止**：
- 不要用 `s.group && s.group.shapes` 兜底判定 group。非 group 节点也可能出现空 group proxy，必须只信 `s.type === 'Group'`。
- 不要把 `items/group/shapes/items/...` 嵌套路径无条件 load 到普通 shape。Mac LTSC 单选 `GeometricShape` 时会在 `ctx.sync()` 抛 `GeneralException`。
- 调试日志也不能在普通 shape 上“试读” `s.group`。即使同步前的属性访问异常被 catch，无效代理仍可能污染队列，让下一次无关的 adjustment `ctx.sync()` 抛 `GeneralException`。必须先 `driver.isGroup(s)`，再访问 `driver.groupShapes(s)`。

**caller 改造 pattern**：
```js
const sel = ctx.presentation.getSelectedShapes();
const selLeaves = await driver.loadShapeTree(
  sel,
  'id, name, width, height, adjustments, tags'
);
```

`loadShapeTree` 的加载顺序：
1. 顶层 collection 只 load 普通字段 + `id` + `type`
2. sync 后只找 `type === 'Group'` 的节点
3. 对真实 group 的 `group.shapes` collection 做 collection-level load
4. 每一层一个 sync，直到没有嵌套 group，再 `flattenSelected`

**业务层**：
- `writeRadius` / `pickupFromSelection` / `applyPickedToSelection` / `loadLayoutTags` 接"shape 数组"语义，flatten 后是叶子数组，一样 work
- `applyLayout` / `syncLayoutChildrenR` / `saveLayoutTags` 在整 slide 按 id 查找时，必须用 `loadShapeTree(slide.shapes, fields)` 返回的叶子数组建 Map；只遍历 `slide.shapes.items` 会漏掉 group 内父子
- layout 父/子按 tag 走（`LAYOUT_PARENT_TAG_KEY` / `LAYOUT_CHILD_TAG_KEY`），group 不影响角色识别
- 防误触铁律不动：writeRadius 内部仍查 strict tag，命中跳过——展平后子能正常被拦截/通过

**⚠️ group 整体缩放时禁止重写子 box（v1.3.1 实机调试结论）**：
- PowerPoint 拖 group 手柄时会原生缩放所有后代；monitor 会同时观察到 layout 父的 R 和 size 改变。
- Mac LTSC 在 group 刚完成整体变形后，再逐个给 `group.shapes` 子写 `left/top/width/height`，日志目标坐标即使正确，宿主也可能把部分子甩出 group。
- 顶层选区含 group 时，monitor 联动必须走 **R-only**：调用 `syncLayoutChildrenR`，绝不调用 `applyLayout` / `setBox`。
- 用户在布局面板主动改 rows/cols/padding/gutter 时仍走完整 `applyLayout`；用户进入 group 后选中单个 layout 父时也保留叶子级几何联动。

**⚠️ 多个 group 子的 box 必须逐子 sync（v1.3.1 实机调试尝试，最终由安全事务取代）**：
- 同一个 `ctx.sync()` 里连续给多个 `group.shapes` 子写 `left/top/width/height`，
  Mac LTSC 16.111 可能只落实最后一个；日志无异常，前三个仍保留 group 原生拉伸后的旧尺寸。
- `applyLayout` 必须 load `Shape.level`；只要任一目标子 `level > 0`，每次
  `driver.setBox(child, box)` 后立刻 `await driver.sync()`，不能批量到 final sync。
- 顶层普通 shape 仍可批量 setBox + 一次 final sync。
- 识别方法：`driver.shapeLevel(s)`；0=顶层，1+=group 内（嵌套 group 会 >1）。

**⚠️ v1.3.1 最终结论：已缩放 group 的后代不可直接写**：
- group 原生缩放**拖拽期间**必须零布局写入；R-only 也不安全。只更新 last-known
  状态，不直接改 group.shapes 的 box / adjustment / tag。
- 最终实测：原生缩放也会把实际 padding / gutter 一起按比例缩放，导致 UI 的厘米值
  与画面不一致。尺寸连续 300ms 不变（用户松手）后，必须走安全完整事务：
  `ungroup → 读取新父 box → applyLayout 完整重算 → regroup`。
- 禁止在 group 内“就地修正”子形状；即使目标坐标正确，Mac LTSC 仍可能再次应用
  group transform，把部分子甩出 group。
- `refreshSelection()` 必须先 `stopLockMonitor()`，清掉旧选区 pending timer；否则旧的
  200ms layout apply 会跨选区执行，绕过 group 零写入保护。
- ungroup/regroup 事务期间必须抑制 `DocumentSelectionChanged` 的并发刷新，完成后由
  原 caller 统一 refresh；不能让第二个 `PowerPoint.run` 插入事务中间。
- 用户主动改变布局参数或 R 联动模式时，父 + 子若是同一顶层 group 的直接成员，
  必须走事务：保存成员/名称/tags → ungroup → fresh 顶层 proxy 写布局 → addGroup 恢复。
- **R-only 收窄**：切换 R 联动模式只允许写 adjustment + layout tags，
  `applyLayout({ writeGeometry: false })` 必须跳过 parent box / computeLayout / 全部 setBox；
  只有 rows/cols/padding/gutter 等几何参数变化才允许完整布局。
- regroup 后用 `Slide.setSelectedShapes([newGroupId])` 恢复 group 选区，避免布局面板消失。
- regroup 的 `DocumentSelectionChanged` 可能延迟到 `PowerPoint.run` 返回后才派发；
  事务结束后的短保护窗口必须忽略它，由 caller 主动 refresh 一次，避免重复 monitor。
- 异常路径必须尽力 regroup，避免把用户的 group 留在解组状态。
- 不同 group / 嵌套 group 的 layout 暂时拒绝写入，提示用户先解开为一个顶层 group。
- `Shape.parentGroup` 在顶层 shape 会抛 `GeneralException`，禁止直接探测；只能使用
  `flattenSelected` 遍历已确认 group 时建立的安全 parent 索引。

**Mock 协议**（test 里用）：
```js
// 真实 PPT：s.type === 'Group' + s.group.shapes.items
// Mock：s._isGroup = true + s._groupShapes = [c1, c2, ...]
```

**限制 / 不支持**：
- 跨 slide 组合（slide A 的 group 跨页引用 slide B 的 shape）—— PowerPoint 不支持，跳过
- "把 R 角应用到整个 group 但 group 内不是 roundRect"——没意义，isRoundRect 过滤掉
- 写 R 角到 group 节点本身——group 不是 roundRect，跳过；只写叶子

### 2.11 TagCollection 批量读取（v1.3.1）

- `shapeCollection.load('items/tags')` 只加载 `tags` 导航属性，**不会**填充
  `shape.tags.items` 的 `key/value`。
- 批量业务必须走 `await driver.loadTagsBulk(shapes)`：
  1. 给每个 `shape.tags` 排队 `load('key, value')`
  2. 只做一次 `ctx.sync()`
  3. 再统一读取 `tags.items`
- PowerPoint 会把 tag key 统一存成大写；radius-core 从批量 dict 取 tag 时必须
  大小写不敏感。
- ❌ 不要在形状循环内逐个 `readTag + sync`，会重现后几个 adjustment 写入丢失。

---

## 3. 部署 / 路径

### 3.1 manifest 路径会被 PowerPoint 重启清空

`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` 在 PowerPoint `Cmd+Q` 退出时可能被回收。`.app` 启动器每次都重新注册。

**用户必须 Cmd+Q 完全退出 PowerPoint**，只关窗口不够（macOS 不真退 Office）。

### 3.2 启动器主动找 node

macOS launchd 的精简 PATH 不一定有 `/opt/homebrew/bin/node`，启动器主动找：

```bash
for p in /opt/homebrew/bin/node /usr/local/bin/node /opt/local/bin/node; do
  if [ -x "$p" ]; then NODE="$p"; break; fi
done
```

### 3.3 localhost 用 HTTP

Office Add-in 对 localhost 允许 HTTP，**不**需要 HTTPS / 证书。manifest 所有 URL 都是 `http://localhost:3000`。

### 3.4 ⚠️ iCloud Documents 下的 dist 重建

如果项目在 `~/Documents/`（iCloud 同步）：
- ❌ 别外层 `mavis-trash dist`（iCloud 把 dist 移到自己的 Trash，30 天后才真删，期间系统反复弹 "无法完成此操作，因为需要下载'dist'"）
- ✅ 直接 `bash tools/build-app.sh` 覆盖式重建（脚本内部 `find $DIST -mindepth 1 -maxdepth 1 -exec rm -rf {} +` 已清掉 `.app` / `.dmg` / `dmg-staging` / `AppIcon.iconset`）

如果已被卡住：点"好"消掉 → `Cmd+Q` PowerPoint → `touch /Users/ma/Documents/minimax/radius_in_ppt/dist` 强制 iCloud 重新拉本地 → 重启 .app

---

## 4. Git 推送

用 `gh` CLI（已登录到 github.com Jerrrry666）：

```bash
git push origin minimax     # 推 commits
git push origin v1.2        # 推新 tag
git push origin :refs/tags/v1.2  # 删旧 tag
git push origin v1.2 --force     # 移动 tag（force-push）
```

**注意**：
- commit message **不要用中文标点**（bash 解析会炸）—— 用 ASCII
- force-push 会改写历史，跟别人协作前先确认

---

## 5. 调试技巧

### 5.1 验证 lock 真的跟文件走

```
1. 选个圆角矩形，点「锁定 R 角」
2. Cmd+S 保存 .pptx
3. Cmd+Q 完全退 PPT
4. 重新打开同一个 .pptx
5. 选中刚才那个圆角矩形
6. 状态卡「已锁定」应该显示 1
```

### 5.2 验证代码改动生效

```
1. 改完 src/ 代码
2. 任务窗格关闭再打开（拿新 .js / .css）
   - 改了 manifest.xml 需要在 PPT 里移除加载项后重新添加
3. 改了 server 跑的内容 → Cmd+Q PowerPoint + 重开
```

### 5.3 看 server 日志

```bash
tail -f /tmp/serve.log
```

### 5.4 driver 烟囱测试

任务窗格 → 点「🧪 Driver 烟囱测试」按钮 → 自动跑 16 个 driver 方法 → 14/14 全过即 verified。

### 5.5 调试 `reason='exception'`

如果某个 feature 失败且只 log `r.reason='exception'`，**第一时间**检查对应函数 catch 块有没有把 `e.message` 主动 `console.log` 出来 —— v1.2.1 教训：吞了 `GeneralException` 找根因找了一晚上。

```js
// 模板
catch (e) {
  const msg = e && e.message ? e.message : String(e);
  console.log('[featureName] EXCEPTION:', msg, '| stack:', e?.stack, '| ctx:', ctx);
  return { ok: false, reason: 'exception', error: msg };
}
```

---

## 6. 详细参考（指向其他文档）

| 想知道什么 | 看哪里 |
| --- | --- |
| 项目状态 / 已完成 / 待办 / Bug / 规划 | [LOG.md](./LOG.md) |
| v1.0 详细变更 + commit 历史 | [changelogs/v1.0.md](./changelogs/v1.0.md) |
| v1.1 详细变更（pipette / 锁定分两态 / bug 修复过程）| [changelogs/v1.1.md](./changelogs/v1.1.md) |
| v1.2 详细变更（布局模式 + 三层架构 + Mac LTSC 坑实战）| [changelogs/v1.2.md](./changelogs/v1.2.md) |
| 完整目录结构 | LOG.md §项目结构 |
| 长期路线图（v1.1+ Stage 1-4）| [plans/feature-roadmap.md](./plans/feature-roadmap.md) |
| 项目介绍 + 用户视角 | [README.md](./README.md) |
| 单元测试运行方法 | [test/README.md](./test/README.md) |
