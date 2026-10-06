# Mac Office.js 功能区实验

分支：`codex/ribbon-officejs-mac`。共同基线：v1.3.2；实验没有递增正式产品版本。

顶部「R角调整 · JS」包含设置R角小窗口、0/0.1/0.3/0.5cm预设、固定R角开关、防误触开关、恢复固定值、吸取/停止和刷入。布局及自定义预设保留在「布局与更多」侧栏。关闭侧栏可使用「隐藏侧栏」按钮。

## Mac实现

- XML使用`PrimaryCommandSurface`和`ExecuteFunction`；修正基线中把`CustomTab`当作ExtensionPoint类型的写法。
- `SharedRuntime 1.1`、`lifetime="long"`、FunctionFile和任务窗格指向同一个HTML，去掉共享运行时不支持的TaskpaneId。
- 所有按钮通过同一个dialog.js控制器、宿主队列、selectionEpoch和monitor运行；隐藏侧栏不主动停止monitor，没有第二套形状监测器。
- Office UI调用封装在office-ui-driver.js；radius-core和ppt-driver的形状业务路径保持原样。保护tag的实时预检、Group安全事务及0cm规则继续生效。
- 数值窗口只传JSON，不能调用形状API；输入期间换选区会拒绝提交，取消/非法消息/异常也完成Office命令。
- 功能区固定R角按钮按各形状当前R锁定，避免隐藏输入框的旧值影响操作。百分比沿用现有侧栏按参考短边换算的语义。
- 使用独立Add-in ID及实验cache参数；原版正式版本号保持1.3.2。两个Office.js server都使用3000端口，实测时只运行一份。

## 本地验证和待验收

`npm test`：296/0（共同基线287项＋9项功能区回归），包含真实UI adapter、防误触、组合、0值、输入选区过期、命令串行、异常完成、运行时配置及资源引用检查。`bash tools/build-app.sh`生成`dist/RadiusInPpt.app`。

微软office-addin-manifest验证器检查后已修正TaskPane不支持的RequestedHeight/Width，以及资源Override的bt命名空间；现无XML Schema错误。完整Marketplace验证仍报3项HTTP URL要求，属于正式商店提交的HTTPS要求；本实验遵循项目Mac localhost HTTP约定，不宣称Marketplace验证全部通过。

Mac LTSC宿主验收尚未运行，不把单测通过视为功能区已被PowerPoint加载。重点确认：顶部控件出现；不打开侧栏应用预设；输入cm/%；锁定后隐藏/关闭侧栏并拖尺寸；保护形状混合多选；嵌套组合；输入途中换页/换形状。若宿主不支持SharedRuntime，VersionOverrides回退到原任务窗格体验。

参考：[功能区命令](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-commands)、[共享运行时](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)。
