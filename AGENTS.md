# R角调整 — 项目协作规则

> 2026-10-07更新。先确认分支和工作目录，再修改代码。
> 当前版本为用户已明确批准的 **v1.4.0**；原生实现已合入本地`main`，尚未push或创建发布tag。

## 0. 当前产品与分支

本项目当前交付macOS PowerPoint **原生PPAM加载项**。安装一次后由PowerPoint加载，在顶部「R角调整 · Native」使用半径、单位、预设及防误触控件。

- 当前开发分支：`main`。
- 当前工作目录：`/Users/ma/Documents/minimax/radius_in_ppt`。
- 原生实现从`codex/ribbon-vba-mac`合入，现为`main`默认产品；`codex/ribbon-officejs-mac`保留作另一条Ribbon实验路线。原生功能在`native/`修改，`src/`保留Office.js迁移对照。
- 交付：`dist/RadiusInPptNative.ppam`和`dist/RadiusInPptNative-mac.zip`。**不要构建.app或DMG，不启动Node/server，不注册wef/manifest。**
- 原目标：Office LTSC Standard for Mac 2021，16.111/26071325。最近布局/半径控件补验宿主：PowerPoint16.113.4/26100421；此前基础/关系记录为16.113.3/26092714。各宿主验收记录不能混为一谈。
- 运行依赖为PowerPoint VBA与Ribbon XML；Office.js的API要求不适用于PPAM。

## 1. 授权与版本

| 操作 | 默认规则 |
| --- | --- |
| 修改代码/文档、运行测试、构建PPAM/ZIP | 可自主执行 |
| 本地更新已安装的本项目插件、使用临时测试文稿验证 | 在用户要求更新插件的范围内执行；保留旧包和文稿 |
| `git add`、commit、push | 必须有用户明确指令；完成本地工作后不要自动执行，也不必反复询问 |
| 合并分支 | 用户明确要求后可提交该分支待合并改动并完成本地merge；不包含push/tag授权 |
| 新建/推送/移动/删除tag、force-push | 必须有用户明确指令 |
| 旧`build-and-deploy.sh`、DMG、task pane app构建 | 当前原生流程不使用 |

版本格式为`vMAJOR.MINOR.PATCH`。PATCH修复可自主递增，MINOR/MAJOR须用户明确授权。本轮已授权v1.4.0，不需要重复确认。同一个尚未发布的改动只占一个版本，不按调试次数递增。更新`package.json`、原生功能区版本文字、changelog及当前文档；不要为PPAM功能递增历史Office.js manifest。

默认完成顺序：实现和review → 原生格式/安装测试 → 涉及VBA或Ribbon时做宿主验证 → 更新记录 → 构建PPAM及ZIP → 说明结果和未验证范围。**构建成功、格式测试通过、宿主编译通过、保存结果正确是不同证据。**

## 2. 原生架构

```text
native/customUI.xml + native/icons/
                ↓
RadiusNativeRibbon.bas       回调、显示值、状态、启用状态、错误提示
                ↓
RadiusNativeCore.bas         数值、单位、限幅、防误触、组合/元数据事务
RadiusNativeRelations.bas    关系归属、待绑定、成员定位、预览编排
                ↓
PptNativeDriver.bas          PowerPoint对象读写的薄封装
                ↓
PowerPoint VBA对象模型
```

事件类只转发选区/窗口刷新和文稿关闭通知，关闭时交由关系业务模块释放会话引用，宿主事件连接走driver。组事务生成的中间选区事件不得触发并发读取或写入，事务结束后统一刷新。

- driver不认识半径/防误触/布局tag的业务含义。
- Ribbon不得直接写形状/adjustment/tag，不复制业务规则。
- core通过driver访问宿主；纯数值函数不能依赖当前选区。
- 不引入Scripting.Dictionary、COM工厂、ActiveX/MSForms、Win32定时器或Windows路径。
- VBA错误必须记录完整`Err.Description`并向调用者传递；恢复组合失败须附带恢复错误，不能只报模糊原因。

## 3. 半径与防误触

- 1cm = 28.3464566929134pt。
- 原生VBA用**1-based** `Adjustments.Item(1)`，值是短边fraction。OOXML采用100000缩放。不要套用Office.js的`get(0)`、proxy/load/sync契约。
- 圆角矩形按真实`AutoShapeType=5`识别；不能把有adjustment的其他形状当圆角矩形。
- cm按每个目标的相同厘米值写入；%按各自短边换算；上限为短边一半，0值有效。
- 微调每次±0.1当前单位，立即走完整写入路径；下限0，%微调上限50。浮点运算保持精度，失败不能显示成已成功应用的新值。
- `radiusLockStrict_v1="1"`禁止任何半径写入。**完整选区预检在任何写入/解组前进行；每个叶子真正写adjustment前再次读取实时tag。** 任一圆角受保护则整批拒绝，不能先写前几个再跳过末项。
- UI写入按钮遇到任一受保护目标时禁用；UI状态不替代core的实时检查。tag读取异常必须报错并拒绝写入，不能当作未保护。
- 只有用户明确操作保护开关/解除按钮才删除strict tag。禁止skipStrict/bypass、自动解除、预设/微调绕过保护。
- 开启保护保留已有`radiusLock_v1`；没有固定值时保存当前R。解除只删除strict，保留固定值。显式修改未保护形状时更新已有固定值tag。
- 状态显示区分未开启、全部开启、部分开启、无圆角选区、组内叶子和读取失败，并显示保护数量/圆角总数。

## 4. Mac组合与tag规则

2026-10-07实机确认：Mac VBA `GroupItems`在嵌套组上可能展平叶子，`ParentGroup`也可能只给外层组。它们可以用于只读预检，**不能据其旧ID推断解组后的直属成员**。

写完整顶层组时：

1. 全选区预检，保存组名称、所有tag及操作参数。
2. `Ungroup`后使用其实际返回的直属成员建立fresh节点。
3. 逐层写入圆角叶子；非圆角成员保留。
4. `Regroup`并恢复名称/全部tag/最终组选区。
5. 异常尽力重组并恢复选区，报告原始错误及恢复错误。

不要对已变形group的旧后代proxy就地写入，不要假设预检数组与`Ungroup`成员一一对应。进入组单选叶子时，R角/保护写入要求改选完整顶层组；父子关系使用ChildShapeRange识别实际叶子，通过安全元数据事务修改其所在组。嵌套深度超过64层拒绝。

tag key大小写不敏感，value必须保留原值。布局tag、父子ID和其他自定义tag不因显式R角操作丢失；重组可能产生新组ID，叶子ID、组层级和几何须验证。带布局tag的形状允许显式半径/保护操作，不能一概拦截。

父子关系仅管理同页圆角矩形的一父多子归属，使用`radiusRelation_v1`及`radiusRelationRole_v1`；不自动布局或联动R角。关系操作不得改fixed/strict/半径，strict对象仍允许绑定/解除。既有关系不得静默覆盖，旧布局显式解除后重建；复制标签造成重复父/子编号时拒绝关系操作。编号预览只写临时副本，原稿不得插入标记；副本关闭时丢弃修改。跨页/文稿取消待绑定。

## 5. 构建与测试

```sh
python3 tools/build-native.py --distribution
# 或 npm run build:native / npm run build
python3 -m venv .venv-native
.venv-native/bin/pip install -r native/requirements-test.txt
.venv-native/bin/python test/test-native-package.py
npm test
```

- `npm run build`应指向原生构建。兼容的`tools/build-app.sh`默认也输出PPAM；名称不是.app交付承诺。
- Python构建器运行只需标准库。`olefile/oletools`为独立消费端测试依赖。重生成本地图标需要Pillow，普通构建和插件运行不需要。
- 原生测试检查OPC类型/关系目标、Mac项目类型、源模块逐字抽取、回调、类模块元数据、嵌入PNG、跨chunk压缩、可重复构建和安装/更新/损坏包拒绝。
- `.bas`为标准模块；`.cls`必须在PROJECT和dir两处都声明为类，带MODULEPRIVATE及完整类属性（尤其VB_Base），不能当标准模块塞入`WithEvents`。缺少类标识的包虽然可抽取，Mac会导致所有回调失效；2026-10-07已实测修复。
- 当前源码按cp1252写入MS-OVBA；VBA源保持可编码文本，中文UI放UTF-8 Ribbon XML。不要静默替换不可编码字符。
- 长VBA模块必须使用正确copy-token压缩。旧raw/literal写法虽然oletools能抽取，Mac会把长模块加载为空；压缩算法改动必须补宿主加载验证。
- `npm test`是保留的Office.js迁移对照回归，不能证明原生VBA业务通过。测试数量随代码变化，以实际运行输出为准。
- VBA、事件类、Ribbon XML/回调、组合或tag路径变更必须验证真实PPT加载/编译与关键动作，再保存临时PPTX独立读OOXML验证。不修改用户真实文稿。
- 普通7页半径基线：`test/native-host-fixture.py`；关系7页及单页批量绑定：`test/native-relations-fixture.py`；布局联动9页：`test/native-layout-fixture.py`。验收分别记录，不修改用户文稿。当前原生格式/安装测试19项，算法自检26项（含8项布局参数微调边界）。自检只测算法，不等于宿主通过。
- 历史诊断source-only PPTM和临时测试PPAM曾触发宿主退出；优先普通PPTX与已安装的真实控件，不重复使用这些诊断产物。

## 6. 安装、更新与实际路径

分发ZIP仅包含`.ppam`、一次文件准备脚本`Install-RadiusInPptNative.command`、`INSTALL.txt`。脚本复制到稳定目录，不注册插件、不启动服务、不改变宏安全设置。

- 默认及本机实际安装目录：`~/Library/Application Support/RadiusInPptNative`。
- 旧包曾使用`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/RadiusInPptNative`，终端可能受macOS容器权限限制。更新时移除旧注册项后添加新目录的PPAM；也可用`--destination`指定其他长期保留的可写目录。
- 不从会重建的`dist/`注册插件，不依赖旧app启动器。
- PowerPoint「工具 → PowerPoint加载项」添加PPAM并保持勾选。只更新已安装插件时不降低全局宏安全设置。
- 更新相同安装路径时，先保存文稿并Cmd+Q完全退出，再替换稳定文件，重新打开验证自动加载。更换安装路径时，须移除旧注册项后添加新路径。已打开的宿主缓存旧VBA/Ribbon，不得直接覆盖后声称生效；用户有未保存文稿时先保存或取消退出，不强制kill。
- 文件准备脚本保留`.previous.ppam`以供恢复。不得覆盖正在加载的包后声称新版已生效。
- 安装/重启后的自动加载需实测，注册成功不代表编译或业务成功。
- 处于iCloud Documents时不要外层整目录trash `dist`；只重建所需文件，不动用户文稿和未授权的历史产物。

## 7. 当前范围与历史对照

原生已提供数值/cm/%、读取、常用预设、多选、半径限幅、加载项写入保护、完整顶层组合事务。v1.4新增状态反馈、全部动作按钮图标、±0.1即时微调，以及指定父/批量绑定子、成员菜单、解除、临时副本编号预览、行列网格和same/subtract/off R角联动。关系/布局及控件优化纳入本地main，沿用未发布的v1.4.0，验收结论以changelog记录为准。

原生布局设置和父变化基线随文稿保存。本加载项修改父R立即联动；直接移动/缩放/黄色手柄修改在改变选区和保存前同步，不逐帧轮询。厘米边距/间距保持不变；off保留子fraction。任何目标子strict在父R/布局写入或解组前拒绝整批，实际写入再读strict。关系元数据仍可绑定/解除strict对象。旋转/翻转成员或组合不支持原生布局及联动，解组前拒绝。

原生防误触当前只阻止本加载项写入，尚未自动纠正用户直接拖黄色手柄或缩放。实时固定R、样式刷、复杂布局联动、自定义预设库和历史尚未迁移，不能把旧task pane的完整功能列为PPAM已实现。

Office.js代码`src/`及相关manifest、server、app工具保留为历史对照。其集合load、0-based adjustment、wef、HTTP和轮询规则仅适用于旧路线。详细教训保留在历史changelog，不与本原生规则混用。

| 内容 | 文档 |
| --- | --- |
| 当前状态/验收 | [LOG.md](LOG.md)、[v1.4变更](changelogs/v1.4.md) |
| 原生方案/安装/限制 | [plans/ribbon-vba-mac.md](plans/ribbon-vba-mac.md)、[native/INSTALL.txt](native/INSTALL.txt) |
| 当前用户说明 | [README.md](README.md)、[README.en.md](README.en.md) |
| 测试与宿主协议 | [test/README.md](test/README.md) |
| 旧路线和历史坑 | [README.taskpane.md](README.taskpane.md)、[changelogs/](changelogs/) |
