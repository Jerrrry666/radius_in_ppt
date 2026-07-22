# R 角调整 — macOS 菜单栏工具

> Mac 顶部菜单栏点一下 → 选「调整 R 角...」→ 输入 `0.3` → PowerPoint 选中的圆角矩形立刻 R 角变成 0.3 厘米。
> 不在 PowerPoint ribbon 里加任何东西，**不依赖加载项机制**，不会因 PowerPoint 更新失效。

## 它解决什么问题

PowerPoint 自带的「圆角矩形」：
- 圆角大小是 **相对值**（短边的 0% ~ 50%），不是绝对值
- 拖动调整形状大小时，R 角会跟着变
- 「设置形状格式」面板里没有以厘米直接输入的入口

这个工具让你：
- ✅ 输入 `0.3 厘米` 这种 **绝对值** 来设置 R 角
- ✅ **锁定 R 角绝对值**：改完大小后一键恢复
- ✅ 同时作用于 **多个选中的圆角矩形**

## 用户视角

1. 双击 `R 角调整.app`（或拖入 `/Applications/`）
2. 首次运行会弹「控制 PowerPoint」授权，点"允许"
3. macOS 顶部菜单栏（屏幕右上角）出现 **R 角图标**
4. 在 PowerPoint 里选中一个或多个圆角矩形
5. 点菜单栏图标 → 「**调整 R 角...**」 → 输入 `0.3` → 回车
6. 形状的 R 角变成 0.3 厘米

菜单内容：

| 菜单项 | 作用 |
|---|---|
| 调整 R 角... | 弹输入框，输入厘米值应用到所有选中的圆角矩形 |
| 锁定当前选区 R 角 | 把当前 R 角绝对值固化到本地文件（`~/Library/Application Support/RadiusInPpt/locks.json`） |
| 解锁当前选区 | 从锁定表移除选中的形状 |
| 重新应用锁定 | 改变形状大小后，用这个把存储的 R 角绝对值重新写回 |
| 在 Finder 中显示锁定文件 | 打开 `locks.json` 位置 |
| 关于 R 角调整 | 版本说明 + GitHub 链接 |
| 退出 | 关闭菜单栏 app |

> **注意**：按钮在 **macOS 顶部菜单栏**（屏幕最右上角，跟 iStat Menus、Cleanshot 那些图标同一排），
> 不在 PowerPoint ribbon 内部。这是为了在 Mac 上**永远 work**——不需要装加载项、不依赖 PowerPoint 版本。

## 关键技术点

### 单位换算

PowerPoint 内部单位是 EMU（English Metric Units）：
- `1 厘米 = 360000 EMU`
- `adjustment 1`（AppleScript 1-based）∈ `[0, 0.5]`（占短边的比例）

设 `shortSide = min(width, height)`（EMU）：
- **读**：`radiusCm = ratio × shortSide / 360000`
- **写**：`ratio = clamp(radiusCm × 360000 / shortSide, 0, 0.5)`

### 为什么用菜单栏 App + AppleScript

Mac PowerPoint 365 上：
- **Office Add-in** 路线：UI 菜单藏起来了，AddIns 目录不自动扫描
- **VSTO / .ppam** 路线：要编译 VBA 二进制，工具链受限
- **菜单栏 App + AppleScript**：零依赖，永远 work ✅

代价：按钮不在 PPT ribbon 里，而是在 macOS 顶部菜单栏。

## 项目结构

```
radius_in_ppt/
├── menubar/                  # 🆕 菜单栏 app 源码（主推）
│   ├── main.swift            # Swift 主程序：NSStatusItem + 菜单
│   ├── menubar-icon.png      # 22x22 模板图标
│   └── menubar-icon@2x.png   # 1024x1024 retina 模板图标
├── src/                      # 备用：Office Add-in 源码（当前不打包）
│   ├── commands/
│   ├── dialog/
│   └── shared/radius.js
├── manifest.xml              # 备用：Office Add-in manifest
├── assets/                   # 公共图标（用于 AppIcon.icns）
├── tools/
│   ├── build-app.sh          # 编译 Swift 菜单栏 app + 打包成 .app
│   ├── build-dmg.sh          # 把 .app 打成 .dmg
│   └── serve.js              # 备用：Office Add-in 的 HTTP server
├── package.json
└── README.md
```

## 构建产物

```
dist/
├── RadiusInPpt.app       (~700KB，含 Swift 二进制 + 图标)
└── RadiusInPpt-1.0.0.dmg (~700KB，可分发的 dmg 安装包)
```

## 从源码构建

```bash
# 要求：macOS 11+、Xcode CommandLine Tools（自带 swiftc）
npm run build:app    # → dist/RadiusInPpt.app
npm run build:dmg    # → dist/RadiusInPpt-1.0.0.dmg
```

或者直接用：

```bash
bash tools/build-app.sh
bash tools/build-dmg.sh
```

## 使用方法

### 1. 安装

```bash
# 双击 .app
open dist/RadiusInPpt.app

# 或拖入 /Applications/
cp -R dist/RadiusInPpt.app /Applications/
```

### 2. 首次运行

- 首次双击未签名 .app 会被 Gatekeeper 拦：右键 .app → 打开 → 弹窗点「打开」（仅一次）
- 启动后顶部菜单栏出现 R 角图标
- 第一次操作 PowerPoint 时 macOS 弹「控制 PowerPoint」授权 → 点「允许」

### 3. 调整 R 角

1. 在 PowerPoint 中选中一个或多个圆角矩形（按住 ⌘ 多选）
2. 点菜单栏 R 角图标 → 「调整 R 角...」
3. 输入厘米值（如 `0.3`）→ 点「应用」

### 4. 锁定 R 角

- 点「锁定当前选区 R 角」：把当前选区里所有圆角矩形的 R 角绝对值记录到锁定表
- 改变形状大小
- 点「重新应用锁定」：按锁定表的 R 角绝对值重新写入

### 5. 退出

- 点菜单「退出」关闭 app
- 重启：双击 .app

## 锁定信息

存储在：`~/Library/Application Support/RadiusInPpt/locks.json`

格式：
```json
{
  "shape-id-1": { "radiusCm": 0.3, "locked": true },
  "shape-id-2": { "radiusCm": 0.5, "locked": true }
}
```

> ⚠️ 锁定信息存在本地文件里，删除 .app / 卸载时不会自动清理。手动删 `locks.json` 即可重置。

## 兼容性

- **macOS** 11+（Big Sur 及以上）
- **PowerPoint for Mac** 2019+（依赖 PowerPoint AppleScript 字典支持）
- Apple Silicon + Intel 通用（编译时只 target arm64，需要 Intel 重新编译或加 x86_64 slice）

## 常见问题

**Q: 双击 .app 提示"无法验证开发者"**
A: Gatekeeper 拦了未签名的 app。右键 .app → 打开 → 在弹窗里再点「打开」（仅一次）。
   长期方案：开发者签名 + 公证（需要 Apple Developer 账号）。

**Q: 菜单栏图标没出现**
A: 看下 macOS 顶部菜单栏右侧——可能图标位置被挤到外面。
   点「在 Finder 中显示锁定文件」找锁定文件位置确认 app 启动成功。

**Q: "控制 PowerPoint"授权弹窗一直不出现 / 点了允许但还是报错**
A: 打开「系统设置 → 隐私与安全性 → 自动化」里，找到 R 角调整，勾选"PowerPoint"。

**Q: 选中圆角矩形后点「调整 R 角...」没反应**
A: 在 PowerPoint 里先**点一下画布**（让 PowerPoint 拿到焦点），再点菜单栏图标。
   或者选区里没有圆角矩形（其他类型形状会被跳过）。

**Q: R 角绝对值改了但形状看起来没变**
A: 选中的可能不是"圆角矩形"形状（"矩形"和"圆角矩形"是两种不同形状）。
   PowerPoint 的"插入 → 形状 → 圆角矩形"才是目标。

## License

MIT
