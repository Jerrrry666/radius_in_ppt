# R 角调整 — Project Memory

> AI agent / 协作者第一次接手这个项目时必读。Mac LTSC Office.js 的行为和
> Windows / Microsoft 365 上有大量差异，这里把踩过的坑都整理好了。

## 1. 项目定位

macOS PowerPoint 加载项，让用户用 **厘米** 设置圆角矩形 R 角（圆角半径），
支持多选和「锁定 R 角绝对值」（改变形状大小时按比例自动调整）。
最终形态是一个 `.app`（约 436 KB），双击启动一个本地静态 server + 把
manifest 注册到 PowerPoint 加载项目录。

参考：iSlide 的「设计」tab 体验。

## 2. 架构概览

**纯 Office.js 路线**（自 `eb5b724` 起，删了所有服务端 .pptx 处理）：

```
┌─────────────┐    HTTP localhost:3000     ┌──────────────────┐
│  PowerPoint │ ◀─────────────────────────▶ │  Office Add-in   │
│  (LTSC Mac) │   Office.js bridge         │  - dialog.html   │
└─────────────┘                            │  - dialog.js     │
        │                                  │  - dialog.css    │
        │ in-memory                        └──────────────────┘
        │ Adjustments.set(0, val)                  ▲
        │ getSelectedShapes()                      │
        ▼                                          │
┌─────────────┐                            ┌──────────────────┐
│  .pptx doc  │   (文件存盘由 PPT 自己处理) │  static server   │
└─────────────┘                            │  tools/serve.js  │
                                           │  ~75 行，只服务  │
                                           │  静态文件        │
                                           └──────────────────┘
```

**为什么不需要服务端处理 .pptx**：
- 之前用 JSZip 在浏览器里读 .pptx，Safari 限制 + Mac LTSC 兼容性差，失败率高
- 切到 Office.js 后，所有读写都在 PowerPoint 进程内完成，**in-memory 即时生效**
- 用户保存 .pptx 时 PPT 自己处理文件 IO，我们完全不碰

## 3. 目录结构

```
.
├── manifest.xml                      # Office Add-in 清单（指向 localhost:3000）
├── src/
│   ├── dialog/
│   │   ├── dialog.html               # 包含 office.js CDN + 调试面板
│   │   ├── dialog.js                 # 核心逻辑（~310 行）
│   │   └── dialog.css
│   └── shared/
│       └── radius.js                 # 常量（PT_PER_CM、ADJ_SCALE）
├── app/
│   └── MacOS/RadiusInPpt             # bash 启动器
├── tools/
│   ├── serve.js                      # ~75 行静态文件 server
│   ├── build-app.sh                  # 打包成 .app
│   └── build-dmg.sh                  # 可选：打包成 .dmg
├── assets/                           # 图标等
├── dist/                             # build 输出（git ignore）
├── test.pptx                         # 测试用文件
└── AGENTS.md                         # ← 本文件
```

## 4. Mac LTSC Office.js 行为差异（重点！）

> 这些是 Mac Office LTSC Standard for Mac 2021（build 16.111）上的实测行为。
> Microsoft 365 / Windows 上的行为可能不一样。**所有 API 行为以 Mac LTSC 为准**。

### 4.1 `Adjustments.value` 单位是 0~1，不是 0~50000

OOXML 里 `<a:gd name="adj" fmla="val X"/>` 的 X ∈ [0, 50000]（对应 0%~50% 短边）。
但 **Mac LTSC dialog 上下文里**，Office.js `shape.adjustments.get(0).value` 返回
的是 **0~1 的小数比例**（OOXML 值 ÷ 50000）。

```js
// ❌ 错误：按 OOXML 假设除以 100000，结果全是 0
const ADJ_SCALE = 100000;
const currentCm = (adj.value / ADJ_SCALE) * minSideCm;  // → 0.00

// ✅ 正确：Mac LTSC 返回的 .value 已经是 0~1 比例
const ADJ_SCALE = 1;
const currentCm = adj.value * minSideCm;  // → 1.28cm
```

**SET 也用 0~1**：`sh.adjustments.set(0, newVal)` 接受 0~1 小数。**不能用
`Math.round`**——`round(0.067) = 0`，所有非整数都会被截成 0 或 1。

```js
// ❌ 错
const newAdj = Math.round((targetCm / minSideCm) * ADJ_SCALE);  // 全 round 成 0

// ✅ 对
const newAdj = (targetCm / minSideCm) * ADJ_SCALE;  // 保留小数
```

来源：调试面板里 `adjustments.get(0).value = 0.17183`，按 OOXML 公式应该是
`17183 / 100000 * 7.46 = 1.28cm`，但用错公式会算成 0。改 ADJ_SCALE=1 后
`0.17183 * 7.46 = 1.28cm` 正确。

### 4.2 `get(0)` 返回 ClientResult 代理，不是 ClientObject

`sh.adjustments.get(0)` 返回的东西 **没有 `.load()` 方法**。直接 `.value` 拿值。
如果想 load nested 属性调 `adjItem.load('value')` 会报
`adjItem.load is not a function`。

```js
// ❌ 错
const adj = sh.adjustments.get(0);
adj.load('value');  // TypeError: adj.load is not a function

// ✅ 对
const adj = sh.adjustments.get(0);
await ctx.sync();
const value = adj.value;  // 直接读
```

### 4.3 `customProperties` 在 dialog 上下文里不可用

`ctx.presentation.customProperties` 在 dialog 里 **返回 undefined**
（不是抛错，是直接没有这个属性）。Task pane 里可能能用，但 dialog 不能。

**降级方案**：锁定信息存 localStorage（key = `radius_in_ppt_locks_v2`）。

```js
let lockBackend = 'none';
if (ctx.presentation.customProperties) {
  // 尝试用 customProperty（key = "lock:{shapeId}"）
  // 实际在 Mac LTSC dialog 走不到这里
} else {
  lockBackend = 'localStorage';
  // 从 localStorage 读 locks
}
```

**代价**：锁不跟着 .pptx 文件走，换台机器/换浏览器/清缓存就丢。
如果用户需要跨设备锁持久化，得把 dialog 迁到 task pane（API 不同）。

### 4.4 `Adjustments.count` 是 primitive，能直接用

`sh.adjustments.count` 在 Mac LTSC dialog 里是 **number**（不是 ClientObject），
不需要 load。可以用它判断是不是圆角矩形：

```js
const adjCount = sh.adjustments.count;  // roundRect=1, 矩形/椭圆=0
const isRoundRect = adjCount > 0;
```

### 4.5 选区 API

| API | Mac LTSC dialog | 备注 |
| --- | --- | --- |
| `ctx.presentation.getSelectedShapes()` | ✅ 工作 | PowerPointApi 1.6+ |
| `sh.width` / `sh.height` | ✅ 工作 | 单位是 pt |
| `sh.id` / `sh.name` | ✅ 工作 | id 是 Office.js 内部 id |
| `Office.context.document.addHandlerAsync(DocumentSelectionChanged, ...)` | ✅ 工作 | Common API，切页也会触发 |

### 4.6 **没有 shape-level change 事件**

Office.js PowerPoint **不提供** `ShapeResized` / `ShapeMoved` / `ShapePropertyChanged`
这类细粒度事件。`Office.EventType` 枚举里能用的只有：

- `DocumentSelectionChanged`（选区变）
- `ActiveViewChanged`（视图变）
- `BindingDataChanged`（Excel/Word 才有）
- `NodeInserted/Deleted/Replaced`（Word CustomXmlPart 才有）

**后果**：检测"形状尺寸变化"只能靠 **setInterval 轮询**。

**当前实现**：500ms 一次轮询，3 次连续无变化（≈1.5s 稳定）= 视为用户松手 → 反算 adj
写回。拖拽中尺寸在变会跳过 apply，避免和拖动手感冲突。

**性能考量**：500ms 间隔 + 只在有 locked 形状时启动，功耗可忽略。

## 5. 部署 / 路径问题

### 5.1 manifest 路径会被 PowerPoint 重启清空

`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` 这个
container 目录 **在 PowerPoint 退出（Cmd+Q）时可能被回收**。

**所以 `.app` 启动器必须每次都**：

```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
cp -f dist/RadiusInPpt.app/Contents/Resources/manifest.xml \
      ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml
```

**用户必须 Cmd+Q 完全退出 PowerPoint**，然后重新打开，新 manifest 才生效。
只关窗口不够（macOS 不会真的退 Office）。

### 5.2 启动器要主动找 node

macOS launchd 的精简 PATH 不一定有 `/opt/homebrew/bin/node`，启动器
（`app/MacOS/RadiusInPpt`）要主动找：

```bash
for p in /opt/homebrew/bin/node /usr/local/bin/node /opt/local/bin/node; do
  if [ -x "$p" ]; then NODE="$p"; break; fi
done
```

### 5.3 localhost 用 HTTP

Office Add-in 允许 `http://localhost` 走 HTTP（**不**需要 HTTPS / 证书）。
manifest 里所有 URL 都是 `http://localhost:3000`。

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
| `eb5b724` | **重构：纯 Office.js，删 server 端 PPTX 处理（-611 行）** |
| `fa1c4ab` | 检测非圆角矩形并 disable apply |
| `4ece5d9` | 用 ClientResult.value 读 adjustments |
| `e2fbfae` | customProperty 不可用时降级到 localStorage |
| `be7e6b1` | 启动器总是 `mkdir -p` wef 目录 |
| `9e9a7b1` | 加调试面板（显示 Office.js raw 值） |
| `ce3e425` | 显式 load adjustments count 和 value |
| `1486764` | 用 `get(0).load('value')` 读（实际是错的，会报 `load is not a function`） |
| `70c1683` | 恢复 `get(0).value`（去掉 .load） |
| `3a92e17` | **ADJ_SCALE 改成 1**（Mac LTSC 返回 0~1 不是 0~50000） |

## 8. 已知限制 / 未来工作

- [ ] **锁定 R 角 cross-machine**：现在 localStorage，换机器就丢
  - 选项 A：把 dialog 迁到 task pane（customProperty 在 task pane 也许能用）
  - 选项 B：lock 存到一个隐藏的 .pptx slide 里（用户看不到但跟着文件走）
- [x] **lock 之后改变形状大小**：✅ 用 setInterval 500ms 轮询 + 3 次稳定检测实现"拖完松手自动重应用"
- [ ] **多选混合**（圆角矩形 + 普通矩形）：现在 UI 标记非圆角 + disable apply，已经可用
- [x] **打包 .dmg**：`tools/build-dmg.sh` 实装完成
- [x] **代码签名 + 公证**：`tools/sign-and-notarize.sh` 实装完成
  - 需要 Apple Developer Program 会员（$99/年）+ Developer ID Application 证书
  - 签名用 hardened runtime（`--options=runtime`），公证走 `xcrun notarytool`
  - 一次性 store credentials：`xcrun notarytool store-credentials "AC_PROFILE" --apple-id ... --password ... --team-id ...`
  - 完整流程详见 `tools/sign-and-notarize.sh` 头部注释
  - 签完直接 `bash tools/build-dmg.sh` 出 .dmg（已集成自动签名检测）

## 9. 调试技巧

### 9.1 调试面板

`src/dialog/dialog.html` 底部有个 `<details>`，展开后是「🔧 调试信息」，会
打印每个选中形状的原始 Office.js 值。改了 Office.js 读法后第一件事是看这里。

### 9.2 重置 .app 状态

```bash
pkill -f "tools/serve.js"
mavis-trash dist
bash tools/build-app.sh
node tools/serve.js > /tmp/serve.log 2>&1 &
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
cp -f dist/RadiusInPpt.app/Contents/Resources/manifest.xml \
      ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml
# 然后用户 Cmd+Q + 重开 PowerPoint
```

### 9.3 看 server 日志

```bash
tail -f /tmp/serve_final.log
```

## 10. PowerPoint 版本

- 目标：**Office LTSC Standard for Mac 2021**（build 16.111 / 26071325）
- API 范围：PowerPointApi 1.1 ~ 1.10
- 已验证可用：`getSelectedShapes`（1.6）、`Adjustments.get/set`（1.10）、
  `customProperties`（1.7，**dialog 上下文不可用**）
- Microsoft 365 用户理论上也能跑，但有些行为可能跟 LTSC 不一样
