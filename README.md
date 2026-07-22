# R 角调整 — PowerPoint 圆角矩形 R 角精确控制加载项

> 像 iSlide 一样在 PowerPoint 顶部添加一个 **「R 角调整」** 自定义 Tab，
> 让你能用 **厘米** 为单位精确设置圆角矩形的 R 角（圆角半径），并支持「锁定 R 角绝对值」。

## 它解决什么问题

PowerPoint 自带的「圆角矩形」形状：
- 圆角大小是 **相对值**（短边的 0% ~ 50%），不是绝对值
- 当你拖动调整形状大小时，R 角会跟着变
- 「设置形状格式」面板里没有直接以厘米输入的入口

这个加载项让你：
- ✅ 输入 `0.3 厘米` 这种 **绝对值** 来设置 R 角
- ✅ 切换 **锁定** 状态：锁定后，R 角保持厘米值不变，改变形状大小时按比例自动调整
- ✅ 同时作用于 **多个选中的圆角矩形**（就像 PPT 自带的长宽调整一样）

## 用户视角

打包后是 macOS `.app`（≈ 360 KB），双击即可：

1. 双击 `R 角调整.app`
2. 弹一个引导框 → 选「打开 PowerPoint」或「在 Finder 中显示 manifest.xml」
3. 首次使用需要在 PowerPoint 里加载一次：`插入 → 我的加载项 → 开发人员加载项 → 从文件添加 → 选 manifest.xml`
4. 之后 PowerPoint 顶部常驻 **「R 角调整」** Tab
5. 选中圆角矩形 → 点 Tab 里的 **「调整 R 角」** → 弹 Dialog → 输入 0.3 厘米 → 应用 / 锁定

> 因为 macOS PowerPoint 没法让第三方自动注册加载项，
> 所以**首次需要你在 PowerPoint 里点「从文件添加」一次**。之后不用再操作。

## 功能

| 功能 | 说明 |
| --- | --- |
| 顶部自定义 Tab | 「R 角调整」Tab + 「调整 R 角」按钮（与 iSlide 同款位置） |
| 弹出式 Dialog | 360×420 面板，无模式，不挡视野 |
| 厘米输入 | 输入数值 + 「应用 R 角」或回车 |
| 锁定/解锁切换 | 一个按钮，根据当前选区状态显示「🔓 锁定」或「🔒 解锁」 |
| 多选支持 | 同时作用于选中的所有圆角矩形；非圆角矩形被跳过并提示 |
| 实时同步 | Dialog 打开时实时反映选区变化 |
| 锁定自动重应用 | 切换选区时，自动把「锁定 R 角绝对值」重新套用到变化的形状上 |

## 项目结构

```
radius_in_ppt/
├── manifest.xml               # Office 加载项清单（CustomTab + Button）
├── package.json
├── README.md
├── app/                       # .app 模板（被打包进 .app/Contents/）
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
│   ├── serve.js               # 静态 HTTP server (http://localhost:3000)
│   ├── build-app.sh           # 打包成 .app
│   └── build-dmg.sh           # 打包成 .dmg（分发用）
└── dist/                      # 构建产物（git 忽略）
    ├── RadiusInPpt.app
    └── RadiusInPpt-1.0.0.dmg
```

## 三种使用方式

### 方式 1：使用打包好的 .app（推荐日常使用）

适合你自己的 Mac：

```bash
# 一次性：打包成 .app
npm run build:app
# → dist/RadiusInPpt.app

# 把 .app 拖入 /Applications/
# 双击运行
```

> 每次双击 .app 都会启动后台 HTTP server（如果还没运行），
> 并弹引导框告诉你接下来在 PowerPoint 里做什么。

**首次需要做的事**（只需一次）：
1. PowerPoint → **插入** → **我的加载项** → **开发人员加载项**
2. 点 **从文件添加** → 选 `dist/RadiusInPpt.app/Contents/Resources/manifest.xml`
3. PowerPoint 顶部出现 **「R 角调整」** Tab

> 之后这步就完成了，每次用双击 .app 即可。

### 方式 2：.dmg 分发

要发给同事用：

```bash
npm run build:dmg
# → dist/RadiusInPpt-1.0.0.dmg
```

同事操作：双击 dmg → 拖入 /Applications → 双击 .app 启动 → 在 PowerPoint 里**从文件添加 manifest**（同方式 1 的首次操作）。

> ⚠️ 由于 macOS Gatekeeper，未签名的 .app 首次双击会提示"无法验证开发者"。
> 解决：右键 .app → 打开 → 弹出"无法验证"对话框里再点"打开"（仅一次）。
> 长期分发建议：开发者账号签名（`codesign` + `notarytool`）—— 见 [Apple 文档](https://developer.apple.com/documentation/security/notarizing_macos_software_before_distribution)。

### 方式 3：纯开发模式

适合改代码时实时调试：

```bash
npm start
# → [serve] HTTP listening on http://127.0.0.1:3000
```

然后在 PowerPoint 里加载 `manifest.xml`（指向 `http://localhost:3000/manifest.xml`）。
修改 `src/` 任何文件后，关闭并重新打开 Dialog 即可看到效果。

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

### 为什么 Office Add-in 路线（不是菜单栏 App）

Mac PowerPoint 不支持 VSTO 那种传统插件；Office Add-in + CustomTab 是**唯一**能在
PPT ribbon 内部加自定义 Tab 的方式（菜单栏 App 只能把按钮放到 macOS 顶部菜单栏，
不能在 PPT ribbon 内）。iSlide、OneKeyTools 等 Mac PPT 插件都是 Office Add-in + .app 包装。

## 兼容性

- **PowerPoint for Mac** 2019 / 2021 / 2024（CustomTab 需要 Office.js 1.4+）
- **PowerPoint for Windows** 2019+（Windows 用户直接 `npm start` 即可，URL 用 `http://localhost:3000`）
- **PowerPoint for Web** — CustomTab 在 web 上不支持，会回退到「插入 → 我的加载项」里的任务窗格

## 常见问题

**Q: 双击 .app 提示"无法验证开发者"**
A: Gatekeeper 拦了未签名的 app。右键 .app → 打开 → 在弹窗里再点「打开」（仅一次）。
   长期方案：开发者签名 + 公证（需要 Apple Developer 账号）。

**Q: PowerPoint 里看不到「R 角调整」Tab**
A: 检查 PowerPoint 版本是否 ≥ 2019 for Mac。
   也确认一下：插入 → 我的加载项 → 开发人员加载项里，R 角调整是否在列表中。
   不在的话重新"从文件添加"一次 manifest。

**Q: 改了 src/ 代码后没生效**
A: 关闭 Dialog 重新点开即可（Dialog 是新加载 HTML 的）。
   改了 manifest.xml 需要在 PowerPoint 里移除加载项后重新添加。

**Q: 后台 server 怎么停**
A: Terminal 里：`lsof -ti tcp:3000 | xargs kill`

## License

MIT
