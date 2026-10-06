# R角调整 — Mac PowerPoint 原生加载项实验

分支：`codex/ribbon-vba-mac`。目标平台：Office LTSC Standard for Mac 2021。

交付的是安装到 PowerPoint 中的 `.ppam` 插件。顶部「R角调整 · Native」直接提供数值输入、cm/%单位、应用、读取和预设；使用时只打开 PowerPoint，无需启用独立 app 或本地服务器。[Microsoft 文档确认 Mac 支持 VBA 加载项和 Ribbon XML](https://learn.microsoft.com/en-us/office/vba/api/overview/office-mac)。

**当前为实验原型：Mac 宿主加载、编译和重启后的自动加载尚未验收，现有功能也尚未全部迁移。**

## 一次安装

1. 解压构建产物 `dist/RadiusInPptNative-mac.zip`。
2. 将 `.ppam` 放到长期保留的位置。可运行同目录的 `Install-RadiusInPptNative.command` 准备稳定安装文件；也可以手动复制文件。
3. 在 PowerPoint 的「工具 → PowerPoint 加载项」中添加 `.ppam`，保持勾选，按 Office 提示允许此插件的宏。
4. 顶部应出现「R角调整 · Native」。Cmd+Q 完全退出后重新打开，确认选项卡仍在。

目标是完成安装后由 PowerPoint 持续加载。安装脚本只复制文件、在 Finder 中显示位置，PowerPoint 注册仍需第 3 步；脚本不参与日常运行。不要从会被重建的 `dist` 目录注册插件。更新前先在 PowerPoint 中卸载旧项，再替换文件并重新添加。

详细步骤及稳定目录见 [安装说明](native/INSTALL.txt)。

## 原型功能

| 功能 | 状态 |
| --- | --- |
| 顶部半径输入及 cm/% | 已有源码 |
| 读取选区及0/0.1/0.3/0.5cm预设 | 已有源码 |
| 多选圆角、限幅、写入保护 | 已有源码 |
| 整组写入及恢复名称/tag/选区 | 已有安全事务源码 |
| 实时固定R、样式刷、布局、历史 | 尚未迁移 |
| Mac LTSC 宿主加载、编译和启动持久性 | 尚未验收 |

原生保护只阻止本加载项写入，还未自动纠正拖动黄色手柄或尺寸。带布局tag、进入组合单选叶子的写入暂时拒绝。完整实现及限制见 [实验说明](plans/ribbon-vba-mac.md)。

## 构建和验证

```sh
bash tools/build-app.sh
npm test
python3 -m venv .venv-native
.venv-native/bin/pip install -r native/requirements-test.txt
.venv-native/bin/python test/test-native-package.py
```

本分支保留的 `build-app.sh` 名称是兼容原构建入口，默认输出 `.ppam` 和安装 ZIP。ZIP 内只有插件、一次文件准备脚本和安装说明；插件使用不依赖 Python 或 Node。源码采用三层结构：Ribbon UI → RadiusNativeCore → PptNativeDriver。

Office.js 原实现保留为迁移对照，历史说明见 [task pane 文档](README.taskpane.md)。需要对照 app 时显式运行 `bash tools/build-app.sh --legacy-taskpane`；原 `.app`/wef 部署及 DMG 脚本不适用于原生插件的安装。

[English](README.en.md) · [变更日志](changelogs/v1.3.md)
