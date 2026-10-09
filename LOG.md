# LOG

> 2026-10-07：原生PPAM及v1.4.0已按用户指令合入本地`main`，尚未push或创建发布tag。当前协作规则见[AGENTS.md](AGENTS.md)。

2026-10-09按用户要求仅在`main`优化代码。自定义参数面板的未提交改动已保存到Git stash，未合入本轮主线。版本沿用未发布的v1.4.0。安装和OOXML读取器新增失败检查；`npm test`改为原生消费端检查，历史Office.js检查使用`npm run test:legacy`。本轮VBA变更的宿主验证与历史验收分开记录，不能沿用旧包的通过结论。

## 当前状态

| 指标 | 值 |
| --- | --- |
| 交付 | 原生PPAM和一次安装ZIP，由PowerPoint加载 |
| 2026-10-09 main优化 | 事务流程、失败恢复、Ribbon显示快照、安装失败处理已整理；新版已更新，半径、批量绑定、布局控件和独立R角保存核对通过 |
| 功能 | cm/%、读取、预设、多选、限幅、防误触、完整顶层组合事务 |
| v1.4新增 | 所有动作图标、自动保护状态、半径即时微调、父子关系、编号预览、网格布局/R角联动及四个布局输入微调 |
| 本机宿主 | 本轮16.113.4/26100421；历史基础/关系记录16.113.3/26092714；16.111目标待另行验收 |
| 格式与安装测试 | 25项通过；不执行VBA，含复制失败、OOXML重名拒绝及一键脚本失败输出 |
| 父子关系宿主验证 | 本轮组合内批量绑定、待绑定取消及临时副本预览通过；完整7页为历史验收 |
| 布局联动宿主验证 | 本轮same/subtract/off和加载项父R即时联动通过；完整9页移动/缩放联动协议仍待完成 |
| 布局控件补验 | 本轮暂存无写入、普通/嵌套2×2网格、单子行列边界及strict子禁用通过；其他箭头边界为历史补验 |
| 半径控件补验 | 本轮七页半径、20%/20.1%及0保存核对通过；50%箭头边界和控件间距为历史补验 |
| 框内箭头样式诉求 | 尚未实现；公开Ribbon XML未提供自定义内嵌步进框，需用户选择是否改用自定义参数面板 |
| VBA自检 | 最终包26项算法、14条业务病例及2262次断言通过；4次实跑0.41～0.53秒，原稿选区/会话恢复通过 |
| 日常项目检查 | 双击tools/Run-Tests.command或npm run test:quick，25项原生＋287项历史逻辑通过 |
| Office.js迁移对照 | 保留npm run test:legacy，不能代替原生宿主验证 |
| 未迁移 | 实时固定R、复杂布局、样式刷、自定义预设、历史 |

默认构建使用`npm run build`或`python3 tools/build-native.py --distribution`。安装到`~/Library/Application Support/RadiusInPptNative`，不构建.app、不运行server、不注册wef。原生架构为Ribbon/事件通知→RadiusNativeCore、RadiusNativeRelations和RadiusNativeLayout→PptNativeDriver。父子关系管理归属并显示编号，网格布局保留厘米边距/间距，R支持same/subtract/off。原生父R操作即时联动，直接父变化在选区改变和保存前同步。验收记录见[changelogs/v1.4.md](changelogs/v1.4.md)。关系/布局及控件优化纳入本地main，沿用未发布的v1.4.0。

2026-10-09代码整理阶段的23项原生消费端和287项历史回归通过，源码可编码检查及交叉review完成。正常退出PowerPoint后替换稳定文件，重开自动加载main控件。PowerPoint16.113.4/26100421的26项自检与七页半径保存核对通过；组合内一次绑定三个子对象、取消待绑定和编号预览返回原稿通过。布局暂存后保存不写文稿；普通/缩放嵌套组2×2网格、父R即时联动和same/subtract/off分别核对。无关重复父关系旁，独立圆角0.30cm保存通过，损坏联动父仍禁用半径写入。

布局控件协议核对九页保存状态，仅第1、6页应用布局；额外父R修改再恢复使用独立协议，允许按规则更新已有fixed值的六位小数文本，不放宽原布局协议。本轮未重跑完整关系7页及父移动/缩放9页联动协议；真实宿主故障注入恢复、VBE全项目Compile命令及16.111仍未验证。详见[本轮宿主记录](test/native-host-validation-main-20261009.json)。

整理阶段验收包的PPAM SHA-256为`008a16a5551e0cc0394616fb6efe5ee17e6aa0755685ee7137a56d5cf6222657`，ZIP为`63877b36ad2da1d4a515cf9de58cfca9c43bedb562f1f6aae53e7109a509790c`，仅含PPAM、文件准备脚本及说明。安装更新不生成备份；旧原生包与临时PPTX清理后保留源码和验收记录。代码未commit/push/tag。

按用户“点一下快速测试逻辑”的要求，将原算法自检扩展为真实业务快测，补齐半径、保护、组合、元数据、关系和布局的成功/拒绝用例。每例用自己的临时页，调用生产入口并断言最终状态；结束关闭样本、恢复原窗口选区及待绑定/布局草稿。测试期望独立于被测计算，事务guard和事件guard分开，不绕过strict。双击项目检查、完整覆盖清单和本轮优化经验见[test/README.md](test/README.md)。算法、内存业务、控件事件、保存OOXML及故障恢复分别记录。

一键入口最终实跑25项原生消费端＋287项历史逻辑通过。初始锁屏限制解除后，正常退出宿主并更新稳定PPAM，再重开自动加载。最终包在PowerPoint16.113.4/26100421上4次快速自检均为26项算法、14条业务病例、2262次断言、0失败，耗时0.41～0.53秒；未选形状、待绑定父与未应用草稿、重复运行及实际嵌套组内父选区分别核对。编号预览副本中自检禁用，关闭预览返回原稿。独立保存OOXML与九页基线完全一致：32个形状的ID/层级/几何/R/全部tag、29段文字均保留，无测试标记。

补验发现组内单选被外层ShapeRange扩成整组选区，已改为实际ChildShapeRange读数，半径/保护/解除保护在联动计划前拒绝局部选区。真实成员菜单选父后只计1个圆角，读回1.215cm与保存基线一致，写入控件禁用；自检后恢复同一叶子。临时新建组不能可靠生成真实child的7次试验保留为失败证据，最终第7病例改为完整组中的联动%、末项strict及损坏配置零写入；未将child的Core三动作拒绝计为自动通过。移除试验DoEvents及选择helper，专项协议明确剩余范围。

最终已安装PPAM与dist SHA-256均为`6c4a0d2a20221399f85a393fcb19e70bb0fe05c68c0ff5a9ae9b7ad97ef54373`；分发ZIP及详细证据见[本轮快测记录](test/native-host-validation-quickcheck-20261009.json)。旧包和临时文稿清理，仅保留源码、JSON和log；未add/commit/push/tag。真实child的直接Core拒绝、文本/多页/多窗口恢复、完整控件事件与保存重开、真实写失败恢复及16.111仍需专项验收。

## 以下为历史Office.js路线记录

历史功能、测试策略和目录结构保留供迁移参考，不代表当前原生插件已实现。

## 已实现

| 版本 | 范围 |
| --- | --- |
| **v1.0** | R 角单形状 / 多选 / 锁定（shape.tags 持久化）/ 防误触 / 预设库 / 样式刷 / 5 次历史 |
| **v1.1** | 批量化的核心闭环 + 锁定分两态（独立「使用数值固定 R 角」+「防误触」开关）|
| **v1.2** | 布局模式（rows × cols 网格 + padding/gutter 滑块 + R 角联动）+ 三层架构（dialog.js / radius-core / ppt-driver）+ 交互层 verified |
| **v1.3.0** | 测试框架分层（driver 单独验证 + fixtures + harness 模拟功能反馈）+ dialog.js / radius-core 全 driver 化 + Step 3-5 完整收尾 + 修 #1 #2 #3 #4 bug |
| **v1.3.1** | Group 兼容 bugfix：组合内形状读取/布局、缩放后固定边距/间距重排、R-only 联动与链条按钮修复 |
| **v1.3.2（本地）** | 防误触完整预检、保护tag错误处理、0值、真实几何识别、宿主串行与过期选区、组合安全事务及打包/服务修复 |

**核心架构**（v1.2 落地）：
```
dialog.js (UI 层)            事件绑定 / 渲染 / toast / debug log
        │
        ▼
radius-core.js (实现层)      算法、保护规则、组合安全事务及布局/R角业务
        │
        ▼
ppt-driver.js (交互层)       形状树、几何识别、属性/tag读写、组合API及宿主队列
        │
        ▼
Office.js + PowerPoint (Mac LTSC 16.111)
```

约束：
- **driver 不知道任何业务概念**（不认 `LOCK_TAG_KEY` / `LAYOUT_PARENT_TAG_KEY`，不知 strict 是什么）
- **radius-core 不 import Office.js**（所有形状读/写/load/sync 走 driver）
- **dialog.js 是搬运工**（`onClick → 开 driver → 调 feature → 渲染结果`）

---

## 待办

按依赖关系，从近到远：

### Step 6 — 路线图余下（v1.4+ 候选）

详见 [plans/feature-roadmap.md](./plans/feature-roadmap.md)：
- 3.1 嵌套等距缩进 R 角（外层 + 内层 + 边距 d → 内层 R 自动 = 外层 R − d）
- 3.4 history 跨 session 持久化（关 PPT 不丢）
- 3.5 黄金比例 R 角建议（10/20/30% 短边一键）
- 3.6 视觉比例统一（按各自短边 X% 批量）
- 3.7 直角 ↔ 圆角一键转换
- 3.8/3.9/3.10 快捷键 / 滑块预览 / 暗色模式

**前置依赖**：3.1 跟现有 layout 模式有重叠风险（都是父子联动），先做 3.4 再上 3.1 避免重写。

### 可选 — dialog.js UI 层进一步重构（路线图外）

dialog.js 现在 2379 行（v1.3.0），离 v1.2 路线图「500 行」目标差 1879 行。
- layout setup UI（手动指定父子）+ presets UI + pipette UI + renderLayoutPanel 都很长
- 结构化重构成可选项：抽 view module / 抽 render module

---

## 已知 Bug / 限制

| # | 优先级 | 描述 | 状态 |
|---|--------|------|------|
| 1 | P2 | pipette 吸取后无法刷入任何形状 | ✅ v1.3 修 |
| 2 | P2 | 布局 R 角联动失败（拖父 R 角子不变）| ✅ v1.3 修 |
| 2b | P2 | **子 bug**：4 个子只写 2 个（Mac LTSC per-call sync 累积）| ✅ v1.3 修（同一类坑）|
| 3 | P3 | lockMonitor 偶发 `GeneralException` | ✅ v1.3 修（同一类坑，collection-level load + readTagsBulk）|
| 4 | P3 | 调试 log 还开着（`[applyLayout/driver]` 等）| ✅ v1.3 修（保留 2 个 catch 兜底）|
| 5 | P4 | driver 烟囱测试 step 6 setAdjFraction 总走「跨 run 兜底」路径 | **已知**（Mac LTSC 限制，不修）|

---

## 关键设计决策

1. **shape.tags 持久化**（Mac LTSC 唯一 work 的方案）
   - `customProperties` / `customXmlParts` 在 Mac LTSC task pane 都不可用
   - shape.tags 直接挂 OOXML `<p:tagLst>` 段，跟 .pptx 文件走
   - 保存 .pptx → 关 PPT → 重开 → tag 还在

2. **Driver 层不 throw**
   - `driver.adjFraction` 内部 try/catch 返回 0（defensive）
   - driver API 契约：永不 throw，调用方不需要保护

3. **set+read 必须 fresh get(0) AFTER sync**
   - Mac LTSC proxy 是 snapshot 风格
   - set 之后旧 proxy 不会 reload value
   - 跨 PowerPoint.run 兜底

4. **防误触 = 最高优先级**
   - strict tag = "1" 的形状，任何 R 角写入路径都不能跳过
   - 两道防线：内存层 + PPT 层（防 race）

5. **未来 feature 信任单元测试**
   - driver 16 方法烟囱测试 14/14 + 112 个单测 + 7 场景 PPT 验证
   - → Step 3c/4/5 不再 PPT 实测

---

## 长期规划

详见 [plans/feature-roadmap.md](./plans/feature-roadmap.md)（v1.1+ 路线图，10 个 P0/P1/P2 功能，4 个 Stage）

当前 Stage 1（v1.0 基础）✅
当前 Stage 2（v1.1 批量化）✅
当前 Stage 3（v1.2 嵌套布局）✅
Stage 4（v1.3+ history 跨 session 持久化 + 嵌套等距缩进）⏳

---

## 项目结构

```
radius_in_ppt/
├── manifest.xml                       # Office Add-in 清单（指向 localhost:3000）
├── src/
│   ├── dialog/                        # task pane UI
│   │   ├── dialog.html
│   │   ├── dialog.js                  # ~2400 行（v1.2 路线图目标 ~500，可选重构）
│   │   └── dialog.css
│   └── lib/                           # v1.2 抽出的实现层 + 交互层
│       ├── radius-core.js             # ~1170 行（v1.3 全 driver 化完成）
│       └── ppt-driver.js              # ~150 行（含 readTagsBulk 一次性拿全部 tag）
├── app/MacOS/RadiusInPpt              # bash 启动器
├── tools/
│   ├── serve.js                       # ~60 行静态 server
│   ├── build-app.sh                   # 打包 .app
│   ├── build-and-deploy.sh            # 一键 build + 部署 + git commit
│   ├── build-dmg.sh                   # 打包 .dmg
│   └── sign-and-notarize.sh           # 公证
├── assets/                            # ribbon icon（5 个尺寸，manifest.xml 引用）
├── test/                              # 单元测试（95 个，v1.3 分层后）
│   ├── fixtures.js                   # 标准 5+ R 角矩形
│   ├── test-harness.js               # createHarness + assertShape + assertCalled (debug)
│   ├── test-radius-core.js           # 纯算法（46）
│   ├── test-features.js              # 功能测试（49）
│   └── README.md
├── dist/                              # build 输出（git ignore）
├── AGENTS.md                          # 三层架构 + Mac LTSC 踩坑
├── LOG.md                             # 本文件 — 主日志
├── README.md
├── changelogs/                        # 子 log（per-version 详细变更）
│   ├── v1.0.md
│   ├── v1.1.md
│   └── v1.2.md
├── plans/
│   └── feature-roadmap.md             # v1.1+ 路线图
└── package.json                       # npm test
```

---

## 测试

```bash
cd /Users/ma/Documents/minimax/radius_in_ppt
npm test                                            # 跑全部 2 个测试文件（95 个）
node test/test-radius-core.js                       # 仅算法（46 个）
node test/test-features.js                          # 仅功能（49 个）
```

**测试分层**（v1.3 重整后）：

| 层 | 文件 | 测什么 | 怎么跑 |
| --- | --- | --- | --- |
| driver 层 | `ppt-driver.js` 16 方法 | Mac LTSC Office.js 兼容性 | **真实 PPT 烟囱测试**（不在 npm test） |
| 纯算法 | `test-radius-core.js` | `computeLayout` / `valueToCm` / 业务规则 | npm test |
| 功能 | `test-features.js` | 业务函数（`writeRadius` / `applyLayout` / `syncLayoutChildrenR` / `readLockState` / `writeLockState` / `reapplyLock`） | npm test |

**fixtures + harness**：

- `test/fixtures.js` — 标准 5+ R 角矩形（basic/medium/large/tiny/wide + locked/strict/locked+strict + clamp 边界 + 0 尺寸 + 非圆角 + layout 父子）
- `test/test-harness.js` — `createHarness` + `assertShape`（**主断言**） + `assertCalled`（debug 用，不作为主断言）

写法新功能测试（功能层）：

```js
const f = makeStandardFixture();
const h = createHarness({ shapes: f.allShapes });
const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);
// 验证最终状态（功能测试只关心"调完后 shape 长啥样"）
h.assertShape(f.shapes.r1_basic, { adjFraction: 0.5 / 3, tags: {} });
```

**driver 烟囱测试（PPT 内）**：任务窗格 → 点「🧪 Driver 烟囱测试」按钮 → 14/14 全过即 driver verified。

---

## 部署

```bash
bash tools/build-and-deploy.sh <version> "<commit msg>"   # 一键：bump + build + 部署 + commit
git push origin minimax                                  # 推 commits
git push origin v1.2                                     # 推 tag（移动用 git tag -f + force push）
```

注意：
- `build-and-deploy.sh` 会自动 bump manifest `<Version>` + cache buster `?v=`，确保用户拿到新代码
- 改了代码后 PPT 需要 `Cmd+Q` 完全退出再重开
- 项目在 `~/Documents/`（iCloud 同步），别外层 `mavis-trash` 整个目录，会卡住
