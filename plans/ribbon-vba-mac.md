# Mac原生功能区方案与验收

当前分支：`main`。原生实现从`codex/ribbon-vba-mac`合入本地`main`，v1.4.0已获用户批准并完成本机验收；尚未push或创建发布tag。

交付目标以2026-10-07用户要求为准：安装到PowerPoint中，随后只打开PowerPoint即可使用顶部控件。`main`默认交付PPAM，不再默认构建独立app。Office.js代码及实验分支保留为功能对照。

## 已实现的原型

`RadiusInPptNative.ppam`是独立PowerPoint VBA加载项。顶部「R角调整 · Native」直接显示半径输入框、cm/%单位选择、应用、读取选区、0/0.1/0.3/0.5cm预设、加载项写入保护及算法自检。无需Node或本地HTTP服务。

原生代码维持三层：RadiusNativeRibbon负责控件回调和错误提示；RadiusNativeCore负责单位、限幅、保护及组合事务；PptNativeDriver薄封装PowerPoint原生对象。使用Collection和Object，避免Scripting.Dictionary、COM工厂、ActiveX、Win32定时器或Windows路径。

- 原生VBA采用1-based `Adjustments.Item(1)`；数值按短边fraction处理，本机cm/%读写及保存后的OOXML已验证。
- cm对所有目标应用同一厘米值；%按每个形状自己的短边换算。0值及短边一半限幅均保留。
- 写入前读完整选区的实时保护tag；任一圆角被保护则整批拒绝，真正写adjustment前再次查tag。只有用户点击「解除防误触」才删除保护tag。
- 普通多选及整组选区递归读取；组合写入采用解组、fresh直属成员、写R/tag、重组，保存名称和全部tag。Mac `GroupItems`可展平嵌套叶子，`ParentGroup`也可能只返回外层；事务层不能按这些预检代理的ID匹配解组返回值，必须用`Ungroup`实际返回值逐层建节点。异常尽力恢复组合和选区，不对变形后的group后代就地写入。
- 单独进入组合选中叶子时要求选中完整顶层组合。带布局tag的形状支持显式R角/保护操作，布局JSON、父子ID、名称和位置尺寸保留；自动布局联动尚未迁移。
- 开启保护保留已有固定值；若没有固定值，保存当前R。原生保护只阻止本加载项写入，尚未自动纠正用户直接拖动黄色手柄或尺寸。

## v1.4控件

上下箭头每次±0.1当前单位并立即应用；0下限和%50上限禁用对应箭头。写入成功后提交显示值，失败保留原值。全部动作采用包内PNG图标，原始矢量在native/icons/。防误触toggle及状态文字区分未开启、已开启、部分开启，并显示保护数量；选区事件类通过driver连接，组合事务期间抑制中间刷新。

事件类必须同时具备PROJECT Class声明、dir类类型/私有记录和VB_Base等完整属性。仅添加类型声明的包可被oletools解析，但Mac宿主所有回调失效；已实测补齐后加载和事件正常。

## 构建与格式验证

```sh
npm run build
python3 -m venv .venv-native
.venv-native/bin/pip install -r native/requirements-test.txt
.venv-native/bin/python test/test-native-package.py
```

构建器只依赖Python标准库，生成真正的OPC `.ppam`包、Ribbon XML和MS-CFB/MS-OVBA VBA项目，源代码随包嵌入。PROJECTSYSKIND=Macintosh；库引用按GUID解析，不包含Windows路径。无编译缓存，_VBA_PROJECT版本0xFFFF，请PowerPoint在加载时编译。

17项格式及文件准备测试用独立olefile/oletools读取包、解析项目和抽取3个标准模块及1个事件类，并与原源码逐字比较；另检查关系目标、回调、Mac平台标识、跨chunk读写、缺失回调失败及可重复构建。新增分发内容/可执行权限、空格和中文路径、重复安装/更新旧包恢复、损坏包拒绝测试。长模块采用copy token压缩，新增重叠复制、offset位宽边界和随机不可压缩chunk回归。文件准备测试只写临时目录，不注册插件、不运行PowerPoint。

默认构建输出PPAM和`dist/RadiusInPptNative-mac.zip`；ZIP内仅有插件、一次文件准备脚本及说明。脚本将PPAM放到`~/Library/Application Support/RadiusInPptNative`稳定目录，不修改Office偏好或宏设置；复制文件不等于完成PowerPoint注册。使用时无需运行脚本、Node或Python。旧app构建移为`tools/build-taskpane-app.sh`，仅显式`--legacy-taskpane`时执行。

格式和文件准备测试通过不代表VBA编译、Mac重启加载或宿主功能通过。

2026-10-07实测旧包在输入数值时出现`RadiusNativeRibbon`隐藏模块编译错误。
独立诊断PPTM显示两个长模块加载失败、业务模块为空，VBE编译具体定位到
`RadiusNativeCore.ParseNumber`成员缺失。修正源码压缩并重建后，实际Ribbon数值回调及
全部业务入口运行通过。原包格式测试无法捕获此宿主加载问题。

## 2026-10-07本机宿主验收

实际宿主为PowerPoint16.113.3（26092714），与最初16.111目标build不同。已在稳定安装目录替换并注册PPAM，多次Cmd+Q完整退出后重新打开，原生选项卡自动出现，无需app/server。

- 用户原来报错的布局子选区成功应用0.50cm并读回；移除过宽的`Use the Office.js edition for tagged layouts.`拦截，strict两道检查保留。
- 7页普通PPTX通过真实Ribbon执行；保存后独立读取OOXML确认：布局父/子混合多选0.50cm、末项strict整批零写入、保护/解除后0.30cm、20%=0.40cm、0值、短边一半限幅1.00cm、缩放嵌套组0.30cm及组保护、含strict子组解组前拒绝、箭头不变。
- 组名称/层级、全部原始tag值（含大小写）、叶子ID、位置尺寸保留；只有显式半径操作对应的固定值tag更新，防误触tag只由保护按钮改动。
- `test/native-host-fixture.py verify`自动断言7页最终宿主结果；用例步骤见[test/README.md](../test/README.md)，保存状态见[验收记录](../test/native-host-validation-20261007.json)。算法自检10项通过，17项独立格式/安装测试通过，287项Office.js回归通过。
- 调试用source-only PPTM加载及临时测试PPAM卸载各触发一次宿主退出；这些工具已移除，验收改用不含宏的普通PPTX。旧插件备份、副本、诊断包及本分支旧app构建产物已清理。

## Mac试用与明确限制

1. 解压安装ZIP，把PPAM放到长期保留的目录；可运行同目录一次文件准备脚本。使用测试文稿，在PowerPoint的「工具」菜单进入PowerPoint加载项管理，添加该PPAM并保持勾选，按Office提示允许此插件的宏。不要从会被重建的dist目录注册；详见[native/INSTALL.txt](../native/INSTALL.txt)。
2. Cmd+Q完全退出再打开，确认选项卡仍在；本机16.113.3已通过，首次在其他机器安装仍需确认。更新同一路径时先保存并Cmd+Q完全退出，再替换、重开验证；换路径时移除旧项后添加新路径。
3. 若出现VBA库或编译错误，打开VBE查看References，重新定位本机PowerPoint和Office库并执行Compile。原始`.bas`/`.cls`文件保存在native/，便于诊断或在宿主中新建加载项后导入。
4. 顶部「算法自检」仅验证10项算法；宿主回归另用普通PPTX验证输入、单位、多选、防误触和组合恢复，不能把算法通过当作宿主通过。

本机16.113.3加载/编译、重启加载及上述宿主回归已通过，不能据此覆盖16.111目标版本或所有Mac版本。实时固定R监测、样式刷、复杂布局、自定义预设库及历史尚未移植；输入值只在本次加载项会话保存。现有Office.js代码保持作为完整功能对照。

参考：[Mac Ribbon/VBA支持](https://learn.microsoft.com/en-us/office/vba/api/overview/office-mac)、[Ribbon XML及ppam](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/overview-of-the-office-fluent-ribbon)、[MS-OVBA](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/)、[MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/)。
