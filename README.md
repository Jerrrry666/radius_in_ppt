# R 角调整 — PowerPoint 圆角矩形 R 角精确控制加载项

> 在 PowerPoint（macOS / Windows）的顶部功能区添加一个 **「R 角调整」** 自定义 Tab，
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

## 功能

| 功能 | 说明 |
| --- | --- |
| 顶部自定义 Tab | 「R 角调整」tab，内含「调整 R 角」按钮 |
| 弹出式 Dialog | 点击按钮后打开 360×420 的面板（无模式，不挡视野） |
| 厘米输入 | 输入数值 + 点「应用 R 角」或回车 |
| 锁定/解锁切换 | 一个按钮，根据当前选区状态自动显示「锁定」或「解锁」 |
| 多选支持 | 一次处理所有选中的圆角矩形；非圆角矩形会被跳过并在面板提示 |
| 实时同步 | Dialog 打开时实时反映选区变化 |
| 锁定自动重应用 | 切换选区时自动把「锁定 R 角绝对值」重新套用到变化的形状上 |

## 截图位置

- `assets/icon-80.png` — ribbon 按钮图标
- `assets/icon-32.png` — manifest 引用
- `assets/icon-128.png` — 高分辨率

## 项目结构

```
radius_in_ppt/
├── manifest.xml               # Office 加载项清单（定义 CustomTab + Button）
├── package.json               # node 启动脚本
├── README.md
├── .gitignore
├── assets/                    # 图标（manifest 引用的 URL 全部在这里）
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
└── tools/
    └── serve.js               # 静态 HTTPS server (https://localhost:3000)
```

## 安装到 PowerPoint（开发模式）

加载项目前引用 `https://localhost:3000/...`，所以你需要先把本地仓库跑起来。

### 1. 启动本地 HTTPS server

```bash
# 在项目根目录
npm start
# 输出示例：
#   [serve] HTTPS listening on https://localhost:3000
#   [serve] open: https://localhost:3000/src/dialog/dialog.html
```

> `tools/serve.js` 会自动用 `openssl` 在 `certs/` 下生成自签证书并信任。
> 首次启动会要求输入密码（用于在「钥匙串访问」中创建可信任的自签 CA）。
> Windows 上请改用 [`office-addin-dev-certs`](https://learn.microsoft.com/office/dev/add-ins/testing/trust-sideload-add-in-insiders) 或 `npx http-server -S`。

### 2. 加载加载项到 PowerPoint

1. 打开 PowerPoint for Mac
2. 顶部菜单 **插入 → 我的加载项 → 开发人员加载项（开发）**
   - 英文菜单：**Insert → My Add-ins → Developer Add-ins**
3. 选择「**添加我的加载项**」 → 选清单文件 `manifest.xml`
4. PowerPoint 重启 / 重新加载加载项后，顶部会出现 **「R 角调整」** tab
5. 选中一个或多个圆角矩形 → 点击「调整 R 角」按钮 → 在弹出的 Dialog 里输入 R 角值（厘米）

> Windows 上路径略有不同：**文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录**。
> 也可以直接用 `manifest.xml` 文件旁加载。

## 使用方法

1. **选中** 一个或多个圆角矩形（可多选，按住 ⌘ 多选）
2. 顶部 **「R 角调整」** tab → **「调整 R 角」** 按钮
3. 在弹出的 Dialog 中：
   - 顶部状态卡片显示当前选区信息 + 当前 R 角（取所有选区形状的最小/最大）
   - 输入框中填入厘米值，例如 `0.3` → 点 **「应用 R 角」** 或按回车
   - 点 **「锁定 R 角」** 把当前选区的所有圆角矩形 R 角绝对值固定；再点一次会全部解锁
   - 「重新应用锁定（针对当前选区）」：对那些有锁定标记的形状，重新写入存储的 R 角绝对值（在你改完大小后可用）

### 锁定语义

锁定 R 角 = **保持 R 角绝对值（厘米）不变**。
- 锁定状态下，你改变形状大小，R 角绝对值不变；adjustments[0] 比例自动按新短边重新计算
- 切换选区 / 打开 Dialog 时，插件会自动把有锁定标记的形状重新应用 R 角
- 锁定信息存在浏览器 `localStorage` 中（`radius_in_ppt_locks_v1`），删除加载项或清浏览器存储会丢失

## 单位换算细节

PowerPoint 内部单位是 EMU（English Metric Units）：
- `1 厘米 = 360000 EMU`
- `adjustments[0] ∈ [0, 0.5]`（占短边的比例）

设 `shortSide = min(width, height)`（EMU）：
- 读：`radiusCm = adjustments[0] × shortSide / 360000`
- 写：`adjustments[0] = clamp(radiusCm × 360000 / shortSide, 0, 0.5)`

## 兼容性

- **PowerPoint for Mac** 2019 / 2021 / 2024（CustomTab 需要 Office.js 1.4+）
- **PowerPoint for Windows** 2019 / 2021 / 2024
- **PowerPoint for Web** — CustomTab 在 web 上不支持，会回退到「插入 → 我的加载项」里的任务窗格

## 开发提示

- 修改 `src/` 任何文件后，**关闭并重新打开 Dialog** 即可看到效果
- 修改 `manifest.xml` 后，需要在 PowerPoint 里**移除并重新加载**该加载项
- 浏览器 DevTools：macOS 上可以在加载项上右键 → 「检查元素」打开 DevTools

## 路线图

- [ ] 支持非矩形圆角（圆角梯形、扇形圆角等）
- [ ] 锁定持久化到 shape 自定义 XML 而不是 localStorage
- [ ] 同时调整多选时，支持「按比例缩放 R 角」模式
- [ ] 中英文 UI 切换

## License

MIT
