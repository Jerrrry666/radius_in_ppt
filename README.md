# R 角调整 — macOS PowerPoint 加载项

> 在 PowerPoint 顶部 ribbon 添加一个 **「R 角调整」** 自定义 Tab，
> 像 iSlide 一样。让你能用 **厘米** 为单位精确设置圆角矩形的 R 角（圆角半径），
> 支持「锁定 R 角绝对值」「防误触」「样式刷」「5 次历史」
> 和 **v1.2 新增的「布局模式」**（rows × cols 网格 + 边距/间距滑块 + R 角联动）。

## 📌 v1.2 更新（2026-07-25）

v1.1 把"按设计意图批量设 R 角"做完了，v1.2 把"嵌套圆角矩形分布"从手算升级为滑块实时拖动。

**本次新增**：
- **布局模式（rows × cols 网格）** — 选 1 父 + N 子，拖滑块实时分布子矩形的位置/尺寸，**完全不用手算**
- **边距 / 间距滑块** — padding（子到父边界的距离）+ gutter（子与子的距离），可拖动可输入
- **R 角联动** — 子 R 角按 `max(0, 父R − padding)` 公式自动算（linkRMode: off / same / subtract 三档）
- **嵌套状态持久化** — 父挂 JSON（rows/cols/padding/gutter/linkR/childIds）+ 子挂 `parentShapeId` 双向 tag，跟 .pptx 走
- **手动设置父子** — 进组合 → 选父 → "建布局"按钮，自动 capture 当前选区为子列表
- **stale childIds 过滤** — 写父 tag 前自动过滤不在当前 slide 的子形状（关掉 PPT → 中间页删子 → 不写坏 JSON）

**架构升级**（v1.2.0-v1.3.5）：
- **三层架构落地** — `dialog.js`（UI）+ `src/lib/radius-core.js`（实现，~880 行）+ `src/lib/ppt-driver.js`（交互，16 方法）
- **driver 层 verified** — 16 方法全 PPT 烟囱测试 14/14 ✅
- **单元测试 112/0** — 103 算法 + 70 mock harness + 109 driver 集成，全过
- **未来 feature 走单测** — 之后 Step 3c/4/5 不再 PPT 实测，信任 `npm test` + 代码 review

完整变更日志：[`changelogs/2026-07-24-v1.2.md`](./changelogs/2026-07-24-v1.2.md)（含 v1.2.0 → v1.3.5 全部 hotfix）　·　进度总览：[`PROGRESS.md`](./PROGRESS.md)　·　路线图：[`plans/feature-roadmap.md`](./plans/feature-roadmap.md)

## 它解决什么问题

PowerPoint 自带的「圆角矩形」形状：
- 圆角大小是 **相对值**（短边的 0% ~ 50%），不是绝对值
- 当你拖动调整形状大小时，R 角会跟着变
- 「设置形状格式」面板里没有直接以厘米输入的入口
- **嵌套多个圆角矩形**时，子矩形的位置/尺寸/R 角全部要手算 + 反复微调

这个加载项让你：
- ✅ 输入 `0.3 厘米` 这种 **绝对值** 来设置 R 角
- ✅ 切换 **锁定** 状态：锁定后，R 角保持厘米值不变，改变形状大小时按比例自动调整
- ✅ 同时作用于 **多个选中的圆角矩形**
- ✅ v1.2：1 父 + N 子 → 拖滑块 → 实时均匀分布 + R 角联动

## 用户视角

打包后是 macOS `.app`（约 440 KB），双击即可：

1. 双击 `R 角调整.app`
2. 弹一个引导框 → 选「退出并重新打开 PowerPoint」
3. 重新打开 PowerPoint → 顶部 ribbon 出现 **「R 角调整」** Tab
4. 点 Tab 里的 **「调整 R 角」** 按钮 → 右侧弹出 **task pane**
5. 选中圆角矩形 → task pane 里输入 `0.3` 厘米（或 `10` %）→ 应用 / 锁定 / 防误触
6. **v1.2**：选 1 父 + N 子 → 进组合 → 「建布局」 → 拖滑块实时分布

> 之后每次使用只需双击 .app 即可（server 后台跑，manifest 已持久化）。
> 注意：改了代码需要 `Cmd + Q` 完全退出 PowerPoint 再重开一次，task pane 才会拉新代码。

## 功能

| 功能 | 说明 |
| --- | --- |
| 顶部自定义 Tab | 「R 角调整」Tab + 「调整 R 角」按钮（与 iSlide 同款位置） |
| 任务窗格（task pane） | 360×560 侧边栏，无模式不挡视野 |
| 厘米 / 百分比输入 | `cm` ↔ `%` 切换；% 按形状短边比例解读；输入 + 「应用 R 角」或回车 |
| **v1.2** 布局模式 | 1 父 + N 子，rows×cols 网格 + padding/gutter 滑块 + R 角联动 |
| **v1.2** R 角联动 | 子 R 角按 `max(0, 父R − padding)` 公式自动算；off/same/subtract 三档 |
| **v1.2** 嵌套状态持久化 | 父挂 JSON + 子挂 parentShapeId 双向 tag，跟 .pptx 走 |
| 使用数值固定 R 角 | 按钮开启/关闭；开启后 PPT 内拖尺寸按比例反算回固定值 |
| 防误触 | 独立 toggle；开启时自动用当前 R 角作 fixed value；拒绝 task pane 改值 + 拖 R 角滑块反算 |
| R 角预设库 | 5 个用户自定义预设，名称/数值可编辑，一键应用 |
| R 角样式刷 | 吸取 1 个 roundRect 的 R 角，连刷到其他目标；可选「刷防误触状态」同步源防误触 |
| 多选支持 | 同时作用于选中的所有圆角矩形；非圆角矩形被跳过并提示 |
| 形状列表 | 任务窗格里实时显示每个选中形状的当前 R 角（PPT 内编辑会同步） |
| 锁定自动重应用 | setInterval 轮询，区分「拖尺寸」和「拖 R 角滑块」两种拖动 |
| 持久化 | 锁定信息存到形状自己的 `shape.tags`（OOXML `<p:tagLst>`），跨设备/换机器都保留 |

## 项目结构

```
radius_in_ppt/
├── manifest.xml                       # Office Add-in 清单（指向 localhost:3000）
├── src/
│   ├── dialog/                        # task pane UI（dialog 是历史命名）
│   │   ├── dialog.html
│   │   ├── dialog.js                  # ~2500 行（含 v1.0-v1.2 全部逻辑，Step 5 重构目标 ~500）
│   │   └── dialog.css
│   └── lib/                           # v1.2 抽出的实现层 + 交互层
│       ├── radius-core.js             # ~880 行，纯算法 + driver 版 feature 函数
│       └── ppt-driver.js              # 109 行，16 个 Office.js 交互方法
├── app/MacOS/RadiusInPpt              # bash 启动器
├── tools/
│   ├── serve.js                       # ~60 行静态文件 server
│   ├── build-app.sh                   # 打包成 .app
│   ├── build-and-deploy.sh            # 一键 build + 部署 + git commit
│   ├── build-dmg.sh                   # 可选：打包成 .dmg
│   └── sign-and-notarize.sh           # 可选：代码签名 + 公证
├── assets/                            # ribbon icon（5 个尺寸，manifest.xml 引用）
├── test/                              # 单元测试（112 个）
│   ├── test-radius-core.js            #   103 个纯算法
│   ├── test-mock-harness.js           #   70 个 mock PowerPoint run 上下文
│   ├── test-driver-integration.js     #  109 个 mock driver + radius-core 集成
│   └── README.md
├── dist/                              # build 输出（git ignore）
│   └── RadiusInPpt.app
├── AGENTS.md                          # 必读：三层架构 + Mac LTSC 踩坑
├── PROGRESS.md                        # 项目进度总览
├── README.md                          # 本文件
├── changelogs/                        # 版本日志
│   ├── 2026-07-23.md                  # v1.0
│   ├── 2026-07-23-v1.1.md             # v1.1
│   └── 2026-07-24-v1.2.md             # v1.2 + v1.2.0 → v1.3.5 全部 hotfix
├── plans/
│   └── feature-roadmap.md             # v1.1+ 路线图
└── package.json                       # npm test 跑 3 个测试文件
```

## 三层架构（v1.2 重构成果）

```
dialog.js (UI 层)          事件绑定 / 渲染 / toast / debug log
       │
       ▼
radius-core.js (实现层)    8 个 driver 版函数：writeRadius / readLockState /
                           writeLockState / reapplyLock / applyLayout /
                           syncLayoutChildrenR / writeRadiusToShapePure /
                           applyLayoutPure
                           零 Office.js 调用 → 100% 单元测试
       │
       ▼
ppt-driver.js (交互层)    16 个方法：load / sync / selectedShapes / activeSlide /
                           slideShapes / shapeId / size / box / isRoundRect /
                           adjFraction / loadAdjValue / setBox / setAdjFraction /
                           addTag / deleteTag / readTag
                           零业务逻辑 → 100% 单元测试
       │
       ▼
Office.js + PowerPoint (Mac LTSC 16.111)
```

- **driver 不知道任何业务概念**（不认 `LOCK_TAG_KEY` / `LAYOUT_PARENT_TAG_KEY`，不知 strict 是什么）
- **radius-core 不 import Office.js**（所有形状读/写/load/sync 走 driver）
- **dialog.js 是搬运工**（`onClick → 开 driver → 调 feature → 渲染结果`）

## 安装到 PowerPoint（开发模式）

manifest 引用 `http://localhost:3000`，所以需要先把本地仓库跑起来。

### 1. 启动本地 HTTP server

```bash
# 在项目根目录
npm start
# 输出示例：
#   [serve] HTTP listening on http://127.0.0.1:3000
```

或者直接双击 `R 角调整.app`（它会自动启动 server + 注册加载项）。

### 2. 注册 manifest 到 PowerPoint

**Mac 上的官方加载方式**（[Microsoft 文档](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)）：

```bash
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
cp manifest.xml "$WEF/manifest.xml"
```

> `R 角调整.app` 启动时自动做这一步。

### 3. 加载到 PowerPoint

1. **完全退出 PowerPoint**（`Cmd + Q`，不要只关窗口）
2. **重新打开** PowerPoint
3. 顶部 **「主页」** tab → **「加载项」** 按钮
4. 弹窗里选 **「R 角调整」**
5. 顶部 ribbon 出现 **「R 角调整」** Tab

## 使用方法

### 基础 R 角调整

1. **选中** 一个或多个圆角矩形（可多选，按住 ⌘ 多选）
2. 顶部 **「R 角调整」** Tab → **「调整 R 角」** 按钮
3. 弹出 Dialog：
   - 顶部状态卡片：当前选区信息 + 当前 R 角（多选时显示最小~最大）
   - 输入框填入厘米值，例如 `0.3` → 点 **「应用 R 角」** 或按回车
   - 点 **「锁定 R 角」**：固定当前选区所有圆角矩形的 R 角绝对值；再点 → 全部解锁
   - 「重新应用锁定（针对当前选区）」：在改完大小后，用这个按钮把存储的 R 角绝对值重新写入

### v1.2 布局模式

1. **画好嵌套关系**：1 个大圆角矩形（父）+ N 个小圆角矩形（子），位置/尺寸随意
2. **全选**（父 + 所有子）→ 顶部 **「R 角调整」** Tab → **「调整 R 角」**
3. 选 1 个作为 **父**（在形状列表里点父那行，会出现"建布局"按钮）
4. 填 **rows × cols**（如 `2 × 3`）→ 选 **子列表** → **「建布局」**
5. 拖 **边距/间距** 滑块 → 子矩形实时均匀分布
6. 切 **R 角联动**（off / same / subtract）：
   - `off`：子 R 角不动
   - `same`：子 R 角 = 父 R 角
   - `subtract`（默认）：子 R 角 = `max(0, 父R − padding)`（按 padding 公式自然联动）
7. 改父 R 角 → 所有子 R 角自动按公式更新

### 锁定语义

v1.1 把"锁定"拆成两个独立开关：

| 开关 | 行为 |
| --- | --- |
| **使用数值固定 R 角**（按钮） | 写 fixed value（cm）→ PPT 内拖尺寸反算回 fixed value；拖 R 角滑块视作主动改（更新 fixed value） |
| **防误触**（toggle） | 拒绝 task pane 改值 + PPT 内拖 R 角滑块反算回当前值；开启时自动 lock（用当前 R 角），关闭时只删防误触标记、保留 fixed value |

共同特点：
- 改变形状大小 → `adjustments[0]` 比例自动按新短边重新计算
- 切换选区时自动重应用（lock monitor 50ms 轮询检测拖动）
- 锁定信息存到形状自己的 `shape.tags`（OOXML `<p:tagLst>`），跨设备/换机器都保留

## 关键技术点

### 单位换算

PowerPoint 内部单位是 EMU（English Metric Units）：
- `1 厘米 = 360000 EMU`
- **OOXML 里** `adjustments[0] ∈ [0, 50000]`（对应 0%~50% 短边）
- **Mac LTSC Office.js** `adjustments[0] ∈ [0, 1]`（OOXML 值 ÷ 50000，task pane / dialog 上下文都返回 0~1）

设 `shortSide = min(width, height)`（EMU）：
- **读**：`radiusCm = adjustments[0] × shortSide / 360000`
- **写**：`adjustments[0] = clamp(radiusCm × 360000 / shortSide, 0, 0.5)`

> ⚠️ **写时不要 `Math.round`**——`round(0.067) = 0`，所有非整数都会被截成 0。`adjustments.set(0, newVal)` 接受 0~1 小数。

### 为什么用 HTTP 不是 HTTPS

Office Add-in 对**生产环境**要求 HTTPS，但对 `localhost` / `127.0.0.1` 是例外，
允许 HTTP 加载。所以我们用纯 HTTP 跑本地 server，**完全不需要证书**。

### Mac LTSC Office.js 踩坑

详细见 [`AGENTS.md`](./AGENTS.md) §4，重点几条：
- `shape.adjustments.get(0).value` 返回 **0~1 小数**，不是 OOXML 0~50000
- 必须 **collection-level load**（`sel.load('items/...')`），per-shape load 不 work
- `set + sync` 后必须 **fresh get(0)** 才能读到新值（旧 proxy 是 snapshot 风格）
- Mac LTSC task pane 里 `customProperties` / `customXmlParts` 不可用，**只能**用 `shape.tags` 持久化

## 测试

```bash
cd /Users/ma/Documents/minimax/radius_in_ppt
npm test                                            # 跑全部 3 个测试文件（112 个）
node test/test-radius-core.js                       # 仅算法（103 个）
node test/test-mock-harness.js                      # 仅 mock harness（70 个）
node test/test-driver-integration.js                # 仅 driver 集成（109 个）
```

**烟囱测试（PPT 内）**：任务窗格 → 点「🧪 Driver 烟囱测试」按钮 → 14/14 全过即 driver verified。

**未来 feature 走单测**：v1.3.5 起，所有新 feature 不再 PPT 实测，信任 `npm test` + 代码 review。

## 兼容性

- **PowerPoint for Mac** 2019 / 2021 / 2024（CustomTab 需要 Office.js 1.4+）
- **PowerPoint for Windows** 2019+（用 `localhost:3000` 即可）
- **PowerPoint for Web** — CustomTab 在 web 上不支持，会回退到任务窗格

## 常见问题

**Q: 双击 .app 提示"无法验证开发者"**
A: Gatekeeper 拦了未签名的 app。右键 .app → 打开 → 在弹窗里再点「打开」（仅一次）。

**Q: 主页 → 加载项 找不到 R 角调整**
A: 检查：
1. 文档是否保存到磁盘（macOS Office Add-in 不会加载未保存文档的加载项）
2. 路径 `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml` 是否存在
3. 完全退出 PowerPoint (Cmd+Q) 重新打开一次
4. 打开 PowerPoint 的「开发人员加载项」tab（如果隐藏了）

**Q: 改了 src/ 代码后没生效**
A: 关闭 Dialog 重新点开即可。
   改了 manifest.xml 需要在 PowerPoint 里移除加载项后重新添加。

**Q: 后台 server 怎么停**
A: Terminal：`lsof -ti tcp:3000 | xargs kill`

**Q: 改完代码后 build + 部署到 wef 怎么搞**
A: `bash tools/build-and-deploy.sh <version> "<commit msg>"` 一键搞定（bump version + build + 部署 + commit + 可选 push）

## License

MIT
