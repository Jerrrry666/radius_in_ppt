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

```js
// ❌ 错：task pane 里 .value 报 "结果对象的值尚未加载"
shapes.load('items/adjustments');
await ctx.sync();
const v = sh.adjustments.get(0).value;  // ❌ 报错

// ✅ 对：显式 load 子项
shapes.load('items/adjustments');
await ctx.sync();
sh.adjustments.load('items/value');  // ← 显式 load
await ctx.sync();
const v = sh.adjustments.get(0).value;  // ✅
```

（dialog 上下文里这步可能不必要；task pane 必须显式 load。）

### 4.5 `Adjustments.count` 是 primitive，能直接用

`sh.adjustments.count` 在 Mac LTSC task pane 里是 **number**（不是 ClientObject），不需要 load：

```js
const isRoundRect = sh.adjustments.count > 0;
```

### 4.6 选区 API

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
