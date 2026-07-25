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

### 1.1 ⭐ 版本号规则

格式 `v{MAJOR}.{MINOR}.{PATCH}`：

| 操作 | AI 能否自主 | 例子 |
| --- | --- | --- |
| Bump PATCH（第 3 位） | ✅ 可以 | v1.2.0 → v1.2.1 → v1.2.7 → v1.2.15 → v1.2.30 |
| Bump MINOR（第 2 位） | ❌ 需用户显式指令 | v1.2.x → v1.3.x |
| Bump MAJOR（第 1 位） | ❌ 需用户显式指令 | v1.x.x → v2.x.x |
| Push git tag | ✅ 可以（不动 manifest 版本号）| git push origin v1.2 |

**触发判断时机**：
- 改完代码要 bump → 看这次改动的"语义"是什么
  - 修 bug / 边界防御 / 内部重构 → PATCH
  - 加新 feature / 大改架构 / 破坏性变更 → MINOR（问用户）
  - 跨大版本不兼容 / 整体重写 → MAJOR（问用户）
- 拿不准时 → 问用户

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
