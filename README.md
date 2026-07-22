# R 角调整 — macOS PowerPoint 加载项

> 在 PowerPoint 顶部 ribbon 添加一个 **「R 角调整」** 自定义 Tab，
> 像 iSlide 一样。让你能用 **厘米** 为单位精确设置圆角矩形的 R 角（圆角半径），
> 并支持「锁定 R 角绝对值」和「多选」。

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

打包后是 macOS `.app`（约 360 KB），双击即可：

1. 双击 `R 角调整.app`
2. 弹一个引导框 → 选「退出并重新打开 PowerPoint」
3. 重新打开 PowerPoint → **「主页」tab → 「加载项」按钮**
4. 选「R 角调整」→ 顶部 ribbon 出现 **「R 角调整」** Tab
5. 选中圆角矩形 → 点 Tab 里的 **「调整 R 角」** 按钮 → 弹 Dialog → 输入 `0.3` → 应用 / 锁定

> 之后每次使用只需双击 .app 即可（server 后台跑，manifest 已持久化）。

## 功能

| 功能 | 说明 |
| --- | --- |
| 顶部自定义 Tab | 「R 角调整」Tab + 「调整 R 角」按钮（与 iSlide 同款位置） |
| 弹出式 Dialog | 360×420 面板，无模式，不挡视野 |
| 厘米输入 | 输入数值 + 「应用 R 角」或回车 |
| 锁定/解锁切换 | 一个按钮，根据当前选区状态显示「🔓 锁定」或「🔒 解锁」 |
| 多选支持 | 同时作用于选中的所有圆角矩形；非圆角矩形被跳过并提示 |
| 实时同步 | Dialog 打开时实时反映选区变化 |
| 锁定自动重应用 | 切换选区时自动把「锁定 R 角绝对值」重新套用到变化的形状上 |

## 项目结构

```
radius_in_ppt/
├── manifest.xml               # Office 加载项清单（CustomTab + Button）
├── package.json
├── README.md
├── app/                       # .app 模板
│   ├── Info.plist
│   ├── PkgInfo
│   └── MacOS/
│       └── RadiusInPpt        # 启动脚本（双击 .app 时执行）
├── assets/                    # 5 个尺寸图标
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
├── src/
│   ├── commands/              # ribbon 按钮回调（FunctionFile）
│   │   ├── commands.html
│   │   └── commands.js
│   ├── dialog/                # 弹出的 R 角调整面板
│   │   ├── dialog.html
│   │   ├── dialog.css
│   │   └── dialog.js
│   └── shared/
│       └── radius.js          # 核心：厘米 ↔ adjustments[0] 换算 + 锁定存储
├── tools/
│   ├── build-app.sh           # 打包成 .app
│   ├── build-dmg.sh           # 打包成 .dmg
│   └── serve.js               # 静态 HTTP server (http://localhost:3000)
├── menubar/                   # 备选方案：macOS 菜单栏 app
│   ├── main.swift
│   ├── menubar-icon.png
│   └── menubar-icon@2x.png
└── dist/                      # 构建产物（git 忽略）
    ├── RadiusInPpt.app
    └── RadiusInPpt-1.0.0.dmg
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

- 锁定 = R 角保持厘米值不变
- 改变形状大小 → adjustments[0] 比例自动按新短边重新计算
- 切换选区时自动重应用（SelectionChanged 事件）
- 锁定信息存在 `localStorage`（`radius_in_ppt_locks_v1`）

## 关键技术点

### 单位换算

PowerPoint 内部单位是 EMU（English Metric Units）：
- `1 厘米 = 360000 EMU`
- `adjustments[0] ∈ [0, 0.5]`（占短边的比例）

设 `shortSide = min(width, height)`（EMU）：
- **读**：`radiusCm = adjustments[0] × shortSide / 360000`
- **写**：`adjustments[0] = clamp(radiusCm × 360000 / shortSide, 0, 0.5)`

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
