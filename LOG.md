# LOG

> 项目主日志 — high level 视角
> per-version 详细变更见 [changelogs/](./changelogs/)

## 状态

| 指标 | 值 |
| --- | --- |
| 当前里程碑 | v1.3+ 测试框架重构（driver + radius-core 全 driver 化）|
| 单元测试 | **100 / 0**（算法 46 + driver 集成 54，去重后无重复）|
| Driver 烟囱测试 | **14 / 14** |
| End-to-end PPT 验证 | **6 / 7**（#6 布局 R 角联动 / #7 pipette 刷入 是 feature bug，待修）|
| 未来 feature 测试策略 | 走 `npm test` + 代码 review，**不再 PPT 实测** |

---

## 已实现

| 版本 | 范围 |
| --- | --- |
| **v1.0** | R 角单形状 / 多选 / 锁定（shape.tags 持久化）/ 防误触 / 预设库 / 样式刷 / 5 次历史 |
| **v1.1** | 批量化的核心闭环 + 锁定分两态（独立「使用数值固定 R 角」+「防误触」开关）|
| **v1.2** | 布局模式（rows × cols 网格 + padding/gutter 滑块 + R 角联动）+ 三层架构（dialog.js / radius-core / ppt-driver）+ 交互层 verified |
| **v1.3** | 测试框架重构（fixtures + harness）+ dialog.js / radius-core 全 driver 化 + Step 3-4 迁移 |

**核心架构**（v1.2 落地）：
```
dialog.js (UI 层)            事件绑定 / 渲染 / toast / debug log
        │
        ▼
radius-core.js (实现层)      8 个 driver 版函数（writeRadius / readLockState /
                             writeLockState / reapplyLock / applyLayout /
                             syncLayoutChildrenR / writeRadiusToShapePure /
                             applyLayoutPure）
        │
        ▼
ppt-driver.js (交互层)       16 个方法（load / sync / selectedShapes / activeSlide /
                             slideShapes / shapeId / size / box / isRoundRect /
                             adjFraction / loadAdjValue / setBox / setAdjFraction /
                             addTag / deleteTag / readTag）
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

### Step 3c — layout tag 读写迁移 + stale state 检测
- `radius-core.loadLayoutTags(driver, slide)` —— 读当前 slide 所有 layout tag
- `radius-core.saveLayoutTags(driver, parentShape, childIds, params)` —— 写父 tag
- `refreshSelection` 加 stale state 检测（中间页删除的子 → 自动 unlink）
- **修 #6 bug**：布局 R 角联动失败

### Step 4 — pipette + history 迁移
- 抽 `pickupFromSelection` / `applyPickedToSelection` / `pushHistory` 为 driver 版
- **修 #7 bug**：pipette 吸取后无法刷入任何形状
- 删 dialog.js 旧 pipette + history 逻辑

### Step 5 — dialog.js UI 层重构
- `lockMonitor` 改成 `radius-core.monitorTick(driver, ...)` 纯函数
- 清掉所有调试 log（`[applyLayout/driver]` 等）
- dialog.js 缩到 ~500 行（删 2000 行混合逻辑）
- 修 #3 #4：lockMonitor 偶发 `GeneralException` + 调试 log

---

## 已知 Bug / 限制

| # | 优先级 | 描述 | 归属 |
|---|--------|------|------|
| 1 | P2 | pipette 吸取后无法刷入任何形状（rect / roundRect 都失败）| feature bug，Step 4 修 |
| 2 | P2 | 布局 R 角联动失败：改父 R 角，子 R 角没变化 | feature bug，Step 3c 修 |
| 3 | P3 | lockMonitor 偶发 `GeneralException`（try/catch 兜住，无影响）| Step 5 重构时清 |
| 4 | P3 | 调试 log 还开着（`[applyLayout/driver]` 等）| Step 5 清 |
| 5 | P4 | driver 烟囱测试 step 6 setAdjFraction 总走「跨 run 兜底」路径 | Mac LTSC 限制，已知 |

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
│   │   ├── dialog.js                  # ~2500 行（Step 5 重构目标 ~500）
│   │   └── dialog.css
│   └── lib/                           # v1.2 抽出的实现层 + 交互层
│       ├── radius-core.js             # ~880 行
│       └── ppt-driver.js              # 109 行
├── app/MacOS/RadiusInPpt              # bash 启动器
├── tools/
│   ├── serve.js                       # ~60 行静态 server
│   ├── build-app.sh                   # 打包 .app
│   ├── build-and-deploy.sh            # 一键 build + 部署 + git commit
│   ├── build-dmg.sh                   # 打包 .dmg
│   └── sign-and-notarize.sh           # 公证
├── assets/                            # ribbon icon（5 个尺寸，manifest.xml 引用）
├── test/                              # 单元测试（100 个，v1.3 去重后）
│   ├── fixtures.js                   # 标准 5+ R 角矩形
│   ├── test-harness.js               # driver call tracker + assertion
│   ├── test-radius-core.js           # 纯算法（46）
│   ├── test-driver-integration.js    # driver 集成（54）
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
npm test                                            # 跑全部 2 个测试文件（100 个）
node test/test-radius-core.js                       # 仅算法（46 个）
node test/test-driver-integration.js                # 仅 driver 集成（54 个）
```

**测试框架**（v1.3 重整后）：

- `test/fixtures.js` — 标准 5+ R 角矩形 fixture（basic/medium/large/tiny/wide + locked/strict/locked+strict + clamp 边界 + 0 尺寸 + 非圆角 + layout 父子）
- `test/test-harness.js` — driver call tracker + snapshot + assertion helpers（`assertCalled` / `assertNotCalled` / `assertCallCount` / `assertShape`）
- 2 个测试文件：纯算法（test-radius-core.js，46）+ driver 集成（test-driver-integration.js，54）
- **v1.3 重整**：删了 test-mock-harness.js（70 个，跟 driver 集成 100% 重复）+ 删了 radius-core.writeRadiusToShapePure / applyLayoutPure（业务只走 driver 路径）

写法新功能测试：

```js
const f = makeStandardFixture();
const h = createHarness({ shapes: f.allShapes });
const r = await RC.writeRadius(h.driver, f.shapes.r1_basic, 0.5);
h.assertCalled('setAdjFraction');
h.assertNotCalled('addTag');
h.assertShape(f.shapes.r1_basic, { adjFraction: 0.5 / 3 });
```

**烟囱测试（PPT 内）**：任务窗格 → 点「🧪 Driver 烟囱测试」按钮 → 14/14 全过即 driver verified。

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
