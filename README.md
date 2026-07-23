# R 角调整 — macOS PowerPoint 加载项

> 在 PowerPoint 顶部 ribbon 添加一个 **「R 角调整」** 自定义 Tab，
> 像 iSlide 一样。让你能用 **厘米** 为单位精确设置圆角矩形的 R 角（圆角半径），
> 并支持「锁定 R 角绝对值」和「多选」。

## 📌 v1.1 更新（2026-07-23）

v1.0 是"能精确设 R 角 + 锁定"，v1.1 把这个能力从"一个个设"升级为"按设计意图批量设"。

**本次新增**：
- **R 角预设库** — 5 个用户自定义预设，名称/数值可编辑，一键应用
- **R 角样式刷** — 吸取 1 个 roundRect 的 R 角，连刷到其他目标（idle / sourcing / brushing 三态状态机）
- **「刷防误触状态」** — 样式刷时可选同步源形状的防误触状态
- **「防误触」独立开关** — 之前是「使用数值固定 R 角」的子开关，现在两个独立。开启防误触时自动用当前 R 角作 fixed value，关闭时只删防误触标记、保留 fixed value
- **百分比单位** — 输入框支持 `cm` ↔ `%` 切换，% 模式按短边比例解读
- **实时 R 角显示** — 任务窗格里形状列表每行的 R 角数字随 PPT 内编辑实时更新
- **拖动智能识别** — lock monitor 区分「拖尺寸手柄」（按比例反算）和「拖 R 角黄色滑块」（防误触反算 / 非防误触更新 fixed value）

完整变更日志：[`changelogs/2026-07-23-v1.1.md`](./changelogs/2026-07-23-v1.1.md)　·　路线图：[`plans/feature-roadmap.md`](./plans/feature-roadmap.md)

## 它解决什么问题

PowerPoint 自带的「圆角矩形」形状：
- 圆角大小是 **相对值**（短边的 0% ~ 50%），不是绝对值
- 当你拖动调整形状大小时，R 角会跟着变
- 「设置形状格式」面板里没有直接以厘米输入的入口

这个加载项让你：
- ✅ 输入 `0.3 厘米` 这种 **绝对值** 来设置 R 角
- ✅ 切换 **锁定** 状态：锁定后，R 角保持厘米值不变，改变形状大小时按比例自动调整
- ✅ 同时作用于 **多个选中的圆角矩形**

## 用户视角

打包后是 macOS `.app`（约 440 KB），双击即可：

1. 双击 `R 角调整.app`
2. 弹一个引导框 → 选「退出并重新打开 PowerPoint」
3. 重新打开 PowerPoint → 顶部 ribbon 出现 **「R 角调整」** Tab
4. 点 Tab 里的 **「调整 R 角」** 按钮 → 右侧弹出 **task pane**
5. 选中圆角矩形 → task pane 里输入 `0.3` 厘米（或 `10` %）→ 应用 / 锁定 / 防误触

> 之后每次使用只需双击 .app 即可（server 后台跑，manifest 已持久化）。
> 注意：改了代码需要 `Cmd + Q` 完全退出 PowerPoint 再重开一次，task pane 才会拉新代码。

## 功能

| 功能 | 说明 |
| --- | --- |
| 顶部自定义 Tab | 「R 角调整」Tab + 「调整 R 角」按钮（与 iSlide 同款位置） |
| 任务窗格（task pane） | 360×560 侧边栏，无模式不挡视野（v1.0 起就是 task pane，不是弹出 dialog） |
| 厘米 / 百分比输入 | `cm` ↔ `%` 切换；% 按形状短边比例解读；输入 + 「应用 R 角」或回车 |
| 使用数值固定 R 角 | 按钮开启/关闭；开启后 PPT 内拖尺寸按比例反算回固定值 |
| 防误触 | 独立 toggle；开启时自动用当前 R 角作 fixed value；拒绝 task pane 改值 + 拖 R 角滑块反算 |
| R 角预设库 | 5 个用户自定义预设，名称/数值可编辑，一键应用 |
| R 角样式刷 | 吸取 1 个 roundRect 的 R 角，连刷到其他目标；可选「刷防误触状态」同步源防误触 |
| 多选支持 | 同时作用于选中的所有圆角矩形；非圆角矩形被跳过并提示 |
| 形状列表 | 任务窗格里实时显示每个选中形状的当前 R 角（PPT 内编辑会同步） |
| 锁定自动重应用 | setInterval 轮询，区分「拖尺寸」和「拖 R 角滑块」两种拖动 |
| 持久化 | 锁定信息存在形状自己的 `shape.tags`（OOXML `<p:tagLst>`），跟 .pptx 文件走，跨设备/换机器都保留 |

## 项目结构

```
radius_in_ppt/
├── manifest.xml                       # Office Add-in 清单（指向 localhost:3000）
├── src/
│   └── dialog/                        # task pane UI（dialog 是历史命名）
│       ├── dialog.html
│       ├── dialog.js                  # 核心逻辑（~1300 行）
│       └── dialog.css
├── app/MacOS/RadiusInPpt              # bash 启动器
├── tools/
│   ├── serve.js                       # ~60 行静态文件 server
│   ├── build-app.sh                   # 打包成 .app
│   ├── build-dmg.sh                   # 可选：打包成 .dmg
│   └── sign-and-notarize.sh           # 可选：代码签名 + 公证
├── assets/                            # 图标
├── dist/                              # build 输出（git ignore）
│   └── RadiusInPpt.app
├── test.pptx                          # 测试用文件
├── AGENTS.md                          # 项目专属 AI 协作笔记（Mac LTSC 踩坑）
├── README.md
├── changelogs/                        # 版本日志
│   ├── 2026-07-23.md                  # v1.0
│   └── 2026-07-23-v1.1.md             # v1.1
└── plans/
    └── feature-roadmap.md             # v1.1+ 路线图（10 个 P0/P1/P2，4 个 Stage）
```

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

1. **选中** 一个或多个圆角矩形（可多选，按住 ⌘ 多选）
2. 顶部 **「R 角调整」** Tab → **「调整 R 角」** 按钮
3. 弹出 Dialog：
   - 顶部状态卡片：当前选区信息 + 当前 R 角（多选时显示最小~最大）
   - 输入框填入厘米值，例如 `0.3` → 点 **「应用 R 角」** 或按回车
   - 点 **「锁定 R 角」**：固定当前选区所有圆角矩形的 R 角绝对值；再点 → 全部解锁
   - 「重新应用锁定（针对当前选区）」：在改完大小后，用这个按钮把存储的 R 角绝对值重新写入

### 锁定语义

v1.1 把"锁定"拆成两个独立开关：

| 开关 | 行为 |
| --- | --- |
| **使用数值固定 R 角**（按钮） | 写 fixed value（cm）→ PPT 内拖尺寸反算回 fixed value；拖 R 角滑块视作主动改（更新 fixed value） |
| **防误触**（toggle） | 拒绝 task pane 改值 + PPT 内拖 R 角滑块反算回当前值；开启时自动 lock（用当前 R 角），关闭时只删防误触标记、保留 fixed value |

共同特点：
- 改变形状大小 → `adjustments[0]` 比例自动按新短边重新计算
- 切换选区时自动重应用（lock monitor 50ms 轮询检测拖动）
- 锁定信息存到形状自己的 `shape.tags`（OOXML `<p:tagLst>`），跟 .pptx 文件走，**跨设备/换机器都保留**

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

## 备选方案：menubar/

`menubar/` 目录里有一个 macOS 菜单栏 App 的 Swift 源码（NSStatusItem + AppleScript），
**不依赖 PowerPoint 加载项机制**，直接通过 AppleScript 操作选区。

适用于：PowerPoint AppleScript bridge 完全 broken 的极端情况。

当前默认不编译此 app。如需启用，修改 `tools/build-app.sh`。

## License

MIT
