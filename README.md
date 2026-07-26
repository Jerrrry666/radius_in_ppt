# R 角调整 — macOS PowerPoint 加载项

> 🌍 **[English version](./README.en.md)**

## 关于这个项目

在 **Windows PowerPoint** 上已经有不少成熟的圆角矩形辅助插件，但在 **macOS PowerPoint** 上一直缺少对应的工具。

所以我自己动手做了这个加载项 —— 功能还比较原始，仍在不断迭代中。

**测试范围**：目前只在 macOS PowerPoint（**版本 16.111.1 (26071913)**）上验证过可用性，理论上 Office.js 加载项是跨平台的，Windows 上应该也能使用 —— **不过未在 Windows 上实测**，欢迎 Windows 用户试用后反馈兼容性。

如果你也在 macOS 上做圆角矩形相关的 PPT，欢迎试用 + 提意见 🙏

## 📌 最新更新

v1.2 把"嵌套圆角矩形分布"做完了，v1.3 专注把布局 / 样式刷的细节做顺手，并完成 dialog.js / radius-core 全 driver 化 + Step 3-4 迁移收尾。

**布局模式精修**：
- **行 / 列互斥联动** — 一条滑块就够，列自动 = 子数 ÷ 行（rows × cols = N 严格成立，不留空位）
- **行的可取值改为离散列表** — N 的所有正因子（datalist tick 提示），例如 N=4 → [1, 2, 4]，不会出现 3×2=6 那种"多空位"的情况
- **边距 / 间距锁链联动**（Photoshop 风格）— 中间一个锁链 icon 跨两行垂直居中，激活后间距 = 边距
- **链接状态 gutter 禁用** — 锁链激活时 gutter 控件整体变灰 + 不可交互（直接消除"间距被改 → 锁链改回 → 形状没改回"那类竞态 bug）

**样式刷 strict 双向覆盖**：
- **勾选「刷防误触状态」= 双向覆盖** — 源 strict=true → 目标 strict=true；源 strict=false → 目标 strict=false
- 顺序关键：source=true 时**先写 R 角再加 strict**（避免 writeRadius 被拦截）；source=false 时**先删 strict 再写 R 角**（让 writeRadius 不被拦截）

**架构收尾**：
- **dialog.js / radius-core 全 driver 化** — 8 个 driver 版函数替代 dialog.js 散落的 ctxShape 操作
- **Step 3-4 迁移完成** — layout tag 读写 + pipette 全部走 driver
- **修 3 个遗留 bug** — 样式刷吸取后无法刷入 / 布局 R 角联动 4 子只写 2 / lockMonitor `GeneralException`
- **双语 UI** — 按系统语言自动选 zh / en（ribbon tab、task pane、启动器弹窗全支持）
- **测试 210/0** — 95 features + 115 radius-core

完整变更日志：[`changelogs/v1.3.md`](./changelogs/v1.3.md)　·　历史：[`changelogs/v1.2.md`](./changelogs/v1.2.md)　·　主日志：[`LOG.md`](./LOG.md)

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
- ✅ v1.2 布局模式：1 父 + N 子 → 拖滑块 → 实时均匀分布 + R 角联动
- ✅ v1.3 样式刷：吸 1 个形状的 R 角，连刷其他；可选「刷防误触状态」双向覆盖

## 功能

| 功能 | 说明 |
| --- | --- |
| 顶部自定义 Tab | 「R 角调整」Tab + 「调整 R 角」按钮（与 iSlide 同款位置） |
| 任务窗格（task pane） | 360×560 侧边栏，无模式不挡视野 |
| 厘米 / 百分比输入 | `cm` ↔ `%` 切换；% 按形状短边比例解读；输入 + 「应用 R 角」或回车 |
| **v1.2** 布局模式 | 1 父 + N 子，rows×cols 网格 + padding/gutter 滑块 + R 角联动 |
| **v1.2** R 角联动 | 子 R 角按 `max(0, 父R − padding)` 公式自动算；off / same / subtract 三档 |
| **v1.2** 嵌套状态持久化 | 父挂 JSON + 子挂 parentShapeId 双向 tag，跟 .pptx 走 |
| **v1.3** 行/列互斥联动 | 一条滑块就够，列 = 子数 ÷ 行（rows × cols = N 严格成立） |
| **v1.3** 行离散因子列表 | 行的可取值 = N 的正因子（[1, 2, 4] / [1, 2, 3, 6] / 质数 [1, N]），不留空位 |
| **v1.3** 边距/间距锁链联动 | Photoshop 风格锁链 icon，激活后间距 = 边距；链接状态 gutter 整体禁用（避免误操作竞态） |
| **v1.3** 样式刷 strict 双向覆盖 | 勾选后源 strict 状态**覆盖**到目标（双向：源开启→目标开启，源未开启→目标也解除） |
| 使用数值固定 R 角 | 按钮开启/关闭；开启后 PPT 内拖尺寸按比例反算回固定值 |
| 防误触 | 独立 toggle；开启时自动用当前 R 角作 fixed value；拒绝 task pane 改值 + 拖 R 角滑块反算 |
| R 角预设库 | 5 个用户自定义预设，名称/数值可编辑，一键应用 |
| R 角样式刷 | 吸取 1 个 roundRect 的 R 角，连刷到其他目标；可选「刷防误触状态」双向覆盖 |
| 多选支持 | 同时作用于选中的所有圆角矩形；非圆角矩形被跳过并提示 |
| 形状列表 | 任务窗格里实时显示每个选中形状的当前 R 角（PPT 内编辑会同步） |
| 锁定自动重应用 | setInterval 轮询，区分「拖尺寸」和「拖 R 角滑块」两种拖动 |
| 持久化 | 锁定信息存到形状自己的 `shape.tags`（OOXML `<p:tagLst>`），跨设备/换机器都保留 |

## 使用方法

打包后是 macOS `.app`（约 405 KB），双击即可：

1. 双击 `R 角调整.app`
2. 弹一个引导框 → 选「退出并重新打开 PowerPoint」
3. 重新打开 PowerPoint → 顶部 ribbon 出现 **「R 角调整」** Tab
4. 点 Tab 里的 **「调整 R 角」** 按钮 → 右侧弹出 **task pane**
5. 选中圆角矩形 → task pane 里输入 `0.3` 厘米（或 `10` %）→ 应用 / 锁定 / 防误触
6. **v1.2 布局模式**：选 1 父 + N 子 → 进组合 → 「建布局」 → 拖滑块实时分布

> 之后每次使用只需双击 .app 即可（server 后台跑，manifest 已持久化）。
> 注意：改了代码需要 `Cmd + Q` 完全退出 PowerPoint 再重开一次，task pane 才会拉新代码。

### Windows 安装

GitHub Release 还提供一个 **Windows 版本**（`RadiusInPpt-win.zip`，约 95 KB）。**未在 Windows 实测**，有问题请提 issue。

1. 从 [GitHub Releases](https://github.com/Jerrrry666/radius_in_ppt/releases/tag/v1.3) 下载 `RadiusInPpt-win.zip`
2. 解压到任意目录
3. 确认已装 [Node.js 18+](https://nodejs.org/)（`.bat` 启动器会自动找）
4. 双击 `RadiusInPpt.bat` → 弹框提示「完全退出 PowerPoint 后重启」
5. 打开 PowerPoint → 完全退出（文件 → 退出）→ 重新打开
6. 顶部 ribbon 出现 **「R 角调整」** Tab

> 之后每次使用只需双击 `RadiusInPpt.bat` 即可。
> 日志位置：`%TEMP%\radius_in_ppt.log`（出问题提 issue 时附上）。

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

## License

[MIT](./LICENSE)
