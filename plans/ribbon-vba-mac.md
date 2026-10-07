# Mac原生功能区方案与验收

当前分支：`main`。原生实现从`codex/ribbon-vba-mac`合入本地`main`，v1.4.0已获用户批准并完成本机验收；尚未push或创建发布tag。

交付目标以2026-10-07用户要求为准：安装到PowerPoint中，随后只打开PowerPoint即可使用顶部控件。`main`默认交付PPAM，不再默认构建独立app。Office.js代码及实验分支保留为功能对照。

## 已实现的原型

`RadiusInPptNative.ppam`是独立PowerPoint VBA加载项。顶部「R角调整 · Native」直接显示半径输入框、cm/%单位选择、应用、读取选区、0/0.1/0.3/0.5cm预设、父子关系、加载项写入保护及算法自检。无需Node或本地HTTP服务。

原生代码维持三层：RadiusNativeRibbon负责控件回调和错误提示；业务层由RadiusNativeCore负责单位、限幅、保护及组合事务，RadiusNativeRelations负责关系归属、待绑定状态和预览，RadiusNativeLayout负责网格配置、联动计划及变化检测；PptNativeDriver薄封装PowerPoint原生对象。使用Collection和Object，避免Scripting.Dictionary、COM工厂、ActiveX、Win32定时器或Windows路径。

- 原生VBA采用1-based `Adjustments.Item(1)`；数值按短边fraction处理，本机cm/%读写及保存后的OOXML已验证。
- cm对所有目标应用同一厘米值；%按每个形状自己的短边换算。0值及短边一半限幅均保留。
- 写入前读完整选区的实时保护tag；任一圆角被保护则整批拒绝，真正写adjustment前再次查tag。只有用户点击「解除防误触」才删除保护tag。
- 普通多选及整组选区递归读取；组合写入采用解组、fresh直属成员、写R/tag、重组，保存名称和全部tag。Mac `GroupItems`可展平嵌套叶子，`ParentGroup`也可能只返回外层；事务层不能按这些预检代理的ID匹配解组返回值，必须用`Ungroup`实际返回值逐层建节点。异常尽力恢复组合和选区，不对变形后的group后代就地写入。
- 单独进入组合选中叶子时，R角/保护写入要求选中完整顶层组合。带旧布局tag的形状支持显式R角/保护操作，布局JSON、父子ID、名称和位置尺寸保留；原生自动布局使用显式重建的原生关系。
- 开启保护保留已有固定值；若没有固定值，保存当前R。原生保护只阻止本加载项写入，尚未自动纠正用户直接拖动黄色手柄或尺寸。

## v1.4控件

上下箭头每次±0.1当前单位并立即应用；0下限和%50上限禁用对应箭头。写入成功后提交显示值，失败保留原值。全部动作采用包内PNG图标，原始矢量在native/icons/。防误触toggle及状态文字区分未开启、已开启、部分开启，并显示保护数量；选区事件类通过driver连接，组合事务期间抑制中间刷新。

半径框与布局边距/间距框使用相同sizeString和方向按钮间距。半径、单位使用独立标签减少Mac自动标签留白；两个微调按钮仍在框右侧并排显示。排列统一只修改Ribbon XML，微调回调和VBA载荷保持一致。

用户明确要求的目标是输入框内侧右边上下叠放的步进箭头，现有框外按钮尚未满足这个样式要求。2026-10-07核对微软[2006 Ribbon完整Schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/5f3e35d6-70d6-47ee-9e11-f5499559f93a)及[2009 CT_EditBox定义](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/80932ffa-a0cb-44f5-810a-70d83a6c1f27)：公开自定义控件没有spinner，editBox不支持子控件或箭头内嵌属性。box只能编排独立控件，不能使按钮进入editBox边框。图示中的持续时间是PowerPoint内置步进控件，不等于加载项可创建同类自定义控件。若必须实现该样式，需要另选自定义参数面板路线，宿主连接与交付方案另行验证；不能继续把框外排列调整记为目标样式完成。

复用检查：[微软Office2019 PowerPoint控件表](https://github.com/OfficeDev/office-fluent-ui-command-identifiers/blob/main/Office%202019/powerpointcontrols.xlsx)将AnimationDuration和TransitionDuration列为`control`，而非自定义`editBox`。Schema中的CT_ControlClone不提供getText/onChange，并禁止onAction；command改写只支持onAction及启用状态，不提供数值onChange。因此公开接口不能把内置持续时间框重绑定到R角或四个布局参数。

事件类必须同时具备PROJECT Class声明、dir类类型/私有记录和VB_Base等完整属性。仅添加类型声明的包可被oletools解析，但Mac宿主所有回调失效；已实测补齐后加载和事件正常。

## 父子关系与编号

关系限定为同页圆角矩形的一父多子。先显式指定父，再改选子并绑定；父保存在当前文稿、SlideId和叶子ID的会话状态中，跨页/文稿事件取消待绑定。组内成员通过Selection.HasChildShapeRange/ChildShapeRange获取，避免把外层组误认成被选中的叶子。

每个成员保存`radiusRelation_v1=G01`及`radiusRelationRole_v1=P/C1/C2…`。编号在页内分配，追加子从现存最大子编号之后开始；移除成员不会重排其他子编号。单个对象不能同时属于多组，已有关系需显式解除。原生标签和旧布局标签冲突、重复父/子编号及读取异常均停止关系操作；复制带标签的成员可能需要手工修复冲突。

RadiusNativeRelations生成按叶子ID定位的元数据计划，由RadiusNativeCore在任何写入/解组前解析全部目标并保存旧值。元数据事务沿用Mac安全解组路径：对Ungroup实际返回的直属成员建立fresh节点，逐层改标签后Regroup，恢复组名称、全部tag及最终组选区。此路径不写半径、fixed或strict标记；保护对象允许建立/解除归属。成功路径已独立核对保存结果，异常回滚仅完成源码审查，尚未注入宿主故障。

兼容读取旧`layoutParent_v1` JSON中的childIds和`layoutChild_v1`，菜单以L加父ID表示旧关系；显式解除会更新对应元数据，部分解除保留JSON其他字段和外层原始文本。不完整旧关系的部分解除拒绝，避免丢失无法解析的成员ID；允许整组解除或清理孤立子。现阶段宿主实测覆盖末子解除和孤立子清理，多子部分解除需补验。

「查看关系」动态菜单按编号列出父/子名称，支持成员定位及批量选中。菜单动作检查文稿、页面、关系修订号并重新读取成员，避免使用已过期对象。状态区显示选区角色、关系编号、子数量及未绑定数量。

编号预览使用当前页的临时文稿副本，复制页面尺寸后添加蓝色父标记、棕色子编号和副本说明。原稿不插入任何标记；副本中加载项编辑及自动同步禁用，再次点击关闭并返回原稿。预览是静态快照，关闭丢弃副本修改。

## 原生布局与R角联动

关系的父对象保存`radiusRelationLayout_v1=1|rows|columns|paddingCm|gapCm|mode|auto`，`mode`为same/subtract/off，`auto`为0或1；数值保留6位小数。`radiusRelationBaseline_v1`保存父的left/top/width/height（pt）、R（cm）及子叶子ID:编号，用于检测父变化和增减成员。最后一个子或整组关系解除时清除两项配置。单个成员只属于一组，避免联动循环；旧JSON布局需显式解除并重建。

网格按现存子编号排序，从左到右、从上到下分配；编号空缺不占空格。每格宽为`(父宽-2×边距-(列数-1)×间距)/列数`，高同理。行列限制1–25，最多625个子；更改一维自动计算另一维。不足空间时拒绝，不自动缩小用户指定的边距/间距。配置编辑先留在会话，应用成功后才持久化。行列箭头每次±1，另一维自动配容量；减小箭头在`ceil(子数量/25)`禁用，增加箭头在`min(子数量,25)`禁用。边距/间距箭头每次±0.1cm，下限0，沿用输入校验并保留6位小数。Ribbon仅转发，参数微调算法在RadiusNativeLayout；Mac自定义上下箭头并排呈现。控件按行/列/子R及边距/间距/状态两列排列，标签独立于输入框，避免宿主将标签空间拉宽。

same使用父的实际厘米R；subtract使用`max(0,父R-边距cm)`；off不写子adjustment，几何缩放时保留原fraction。写R前按新的短边限幅，显式写入更新已有fixed tag。通过加载项修改父R时，先把选区R目标与全部相关子目标合成一个计划，再完整预检，任何strict子都在父R写入前阻止整批；UI同样禁用相应父R按钮。

第一次应用布局开启自动联动；关闭仅保存配置/基线，手动应用仍可用。直接父几何或黄色手柄变化在WindowSelectionChange、WindowActivate、SlideSelectionChanged后检测，PresentationBeforeSave另同步全部页；不使用轮询，不逐帧同步拖动。只有R变化时只写子R；父位置尺寸或成员变化时重排子并联动R。事件在事务中只失效刷新，editing阻止并发读取/写入。同步失败显示完整错误，保存前弹窗报告但不取消用户保存。

布局计划和原生父R联动使用Core.ApplyShapePlan，写前保存全部目标旧几何、fraction和被改tag；每个写入再读strict。沿实际Ungroup返回逐层建fresh节点，先改子尺寸再按新短边写R，Regroup恢复组名称和所有tag。成功恢复原选区，空选区保持空；异常尽力恢复几何/R/tag/选区并附恢复错误。旋转/翻转的相关成员或组合在解组前拒绝。父本身受保护但只修改未保护子布局/R时允许，父R和strict不改变。

## 构建与格式验证

```sh
npm run build
python3 -m venv .venv-native
.venv-native/bin/pip install -r native/requirements-test.txt
.venv-native/bin/python test/test-native-package.py
```

构建器只依赖Python标准库，生成真正的OPC `.ppam`包、Ribbon XML和MS-CFB/MS-OVBA VBA项目，源代码随包嵌入。PROJECTSYSKIND=Macintosh；库引用按GUID解析，不包含Windows路径。无编译缓存，_VBA_PROJECT版本0xFFFF，请PowerPoint在加载时编译。

19项格式及文件准备测试用独立olefile/oletools读取包、解析项目和抽取5个标准模块及1个事件类，并与原源码逐字比较；另检查关系目标、回调、Mac平台标识、跨chunk读写、缺失回调失败及可重复构建。覆盖关系/布局控件、三种R模式、动态菜单、本地图标、分发内容/可执行权限、空格和中文路径、重复安装/更新旧包恢复、损坏包拒绝。长模块采用copy token压缩，含重叠复制、offset位宽边界和随机不可压缩chunk回归。文件准备测试只写临时目录，不注册插件、不运行PowerPoint。

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

随后追加父子关系，当前18项格式/安装测试通过。7页普通关系PPTX验证建立、追加、定位、单子/整组解除、缩放嵌套组、strict子、旧布局末子和孤立子、重复父拒绝、跨页取消、非圆角以及R角预设；保存后独立核对所有关系及其他标签、半径、几何、叶子ID和层级。最终包重启后另用新单页样本一次绑定组合内3个子，核对保存结果，并验证跨文稿取消、混合选区编号/未绑定数量和预览开/关。见[关系验收记录](../test/native-host-validation-relations-20261007.json)。追加变更未commit/push，仍属未发布的v1.4.0。旋转/翻转组合、旧布局多子部分解除及异常回滚尚未做完整宿主验收。

## Mac试用与明确限制

1. 解压安装ZIP，把PPAM放到长期保留的目录；可运行同目录一次文件准备脚本。使用测试文稿，在PowerPoint的「工具」菜单进入PowerPoint加载项管理，添加该PPAM并保持勾选，按Office提示允许此插件的宏。不要从会被重建的dist目录注册；详见[native/INSTALL.txt](../native/INSTALL.txt)。
2. Cmd+Q完全退出再打开，确认选项卡仍在；本机16.113.3已通过，首次在其他机器安装仍需确认。更新同一路径时先保存并Cmd+Q完全退出，再替换、重开验证；换路径时移除旧项后添加新路径。
3. 若出现VBA库或编译错误，打开VBE查看References，重新定位本机PowerPoint和Office库并执行Compile。原始`.bas`/`.cls`文件保存在native/，便于诊断或在宿主中新建加载项后导入。
4. 顶部「算法自检」仅验证26项算法（半径10项、布局/联动8项、参数微调8项）；宿主回归另用普通PPTX验证输入、单位、多选、防误触、布局联动和组合恢复，不能把算法通过当作宿主通过。

本机16.113.3加载/编译、重启加载及上述宿主回归已通过，不能据此覆盖16.111目标版本或所有Mac版本。实时固定R监测、样式刷、复杂布局、自定义预设库及历史尚未移植；输入值只在本次加载项会话保存。现有Office.js代码保持作为完整功能对照。

参考：[Mac Ribbon/VBA支持](https://learn.microsoft.com/en-us/office/vba/api/overview/office-mac)、[Ribbon XML及ppam](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/overview-of-the-office-fluent-ribbon)、[MS-OVBA](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/)、[MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/)。
