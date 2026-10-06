# Mac原生功能区实验

分支：`codex/ribbon-vba-mac`，共同基线v1.3.2；不递增正式产品版本。

交付目标以2026-10-07用户要求为准：安装到PowerPoint中，随后只打开PowerPoint即可使用顶部控件。原生分支默认交付PPAM，不再默认构建独立app。Office.js分支保留为功能对照。

## 已实现的原型

`RadiusInPptNative.ppam`是独立PowerPoint VBA加载项。顶部「R角调整 · Native」直接显示半径输入框、cm/%单位选择、应用、读取选区、0/0.1/0.3/0.5cm预设、加载项写入保护及算法自检。无需Node或本地HTTP服务。

原生代码维持三层：RadiusNativeRibbon负责控件回调和错误提示；RadiusNativeCore负责单位、限幅、保护及组合事务；PptNativeDriver薄封装PowerPoint原生对象。使用Collection和Object，避免Scripting.Dictionary、COM工厂、ActiveX、Win32定时器或Windows路径。

- 原生VBA采用1-based `Adjustments.Item(1)`；数值按短边fraction处理，需Mac宿主确认与Office.js的单位一致。
- cm对所有目标应用同一厘米值；%按每个形状自己的短边换算。0值及短边一半限幅均保留。
- 写入前读完整选区的实时保护tag；任一圆角被保护则整批拒绝，真正写adjustment前再次查tag。只有用户点击「解除防误触」才删除保护tag。
- 普通多选及整组选区递归读取；组合写入采用解组、fresh成员、写R/tag、重组，保存名称和全部tag；异常尽力恢复组合和选区，不对变形后的group后代就地写入。
- 单独进入组合选中叶子时，原型拒绝写入，要求选中完整顶层组合。带布局tag的选区也拒绝，避免破坏尚未移植的布局关系。
- 开启保护保留已有固定值；若没有固定值，保存当前R。原生保护只阻止本加载项写入，尚未自动纠正用户直接拖动黄色手柄或尺寸。

## 构建与格式验证

```sh
bash tools/build-app.sh
python3 -m venv .venv-native
.venv-native/bin/pip install -r native/requirements-test.txt
.venv-native/bin/python test/test-native-package.py
```

构建器只依赖Python标准库，生成真正的OPC `.ppam`包、Ribbon XML和MS-CFB/MS-OVBA VBA项目，源代码随包嵌入。PROJECTSYSKIND=Macintosh；库引用按GUID解析，不包含Windows路径。无编译缓存，_VBA_PROJECT版本0xFFFF，请PowerPoint在加载时编译。

10项格式及文件准备测试用独立olefile/oletools读取包、解析项目和抽取3个模块，并与原源码逐字比较；另检查关系目标、回调、Mac平台标识、跨chunk读写、缺失回调失败及可重复构建。新增分发内容/可执行权限、空格和中文路径、重复安装/更新旧包恢复、损坏包拒绝测试。文件准备测试只写临时目录，不注册插件、不运行PowerPoint。

默认构建输出PPAM和`dist/RadiusInPptNative-mac.zip`；ZIP内仅有插件、一次文件准备脚本及说明。脚本将PPAM放到PowerPoint容器下的稳定目录，不修改Office偏好或宏设置；复制文件不等于完成PowerPoint注册。使用时无需运行脚本、Node或Python。旧app构建移为`tools/build-taskpane-app.sh`，仅显式`--legacy-taskpane`时执行。

格式和文件准备测试通过不代表VBA编译、Mac重启加载或宿主功能通过。

## Mac试用与明确限制

1. 解压安装ZIP，把PPAM放到长期保留的目录；可运行同目录一次文件准备脚本。使用测试文稿，在PowerPoint的「工具」菜单进入PowerPoint加载项管理，添加该PPAM并保持勾选，按Office提示允许此插件的宏。不要从会被重建的dist目录注册；详见[native/INSTALL.txt](../native/INSTALL.txt)。
2. Cmd+Q完全退出再打开，确认选项卡仍在；Mac LTSC启动持久性待验收。更新前先卸载旧项，再替换文件并重新添加。
3. 若出现VBA库或编译错误，打开VBE查看References，重新定位本机PowerPoint和Office库并执行Compile。原始`.bas`文件保存在native/，便于诊断或在宿主中新建加载项后导入。
4. 先点顶部「算法自检」，再确认输入框、cm/%、0值、多选、含箭头/十字的选区、防误触和组合恢复。纯算法自检包含4项；本轮没有运行宿主自检。

**未完成Mac宿主加载/编译及实机验收**。当前是可审查、可构建的实验原型，不能称为已兼容Mac LTSC的正式加载项。实时固定R监测、样式刷、复杂布局、自定义预设库及历史尚未移植；输入值只在本次加载项会话保存。现有Office.js代码保持作为完整功能对照。

参考：[Mac Ribbon/VBA支持](https://learn.microsoft.com/en-us/office/vba/api/overview/office-mac)、[Ribbon XML及ppam](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/overview-of-the-office-fluent-ribbon)、[MS-OVBA](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/)、[MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/)。
