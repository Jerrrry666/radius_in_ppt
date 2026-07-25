# radius_in_ppt — 项目进度总览

> 最近更新：2026-07-25 v1.3.5

## TL;DR

macOS PowerPoint task pane 加载项，让用户用 **厘米** 或 **百分比** 设置圆角矩形的 R 角。
支持多选 / 锁定 R 角绝对值 / 防误触 / 样式刷 / 5 次历史 / **布局模式**。

**当前里程碑**：v1.2 布局模式已完成 + 交互层（driver + radius-core + dialog.js）正式 verified。
**当前版本**：v1.3.5（git tag `v1.2` 标记完整 v1.2 里程碑）

---

## 版本历程

| Tag | 日期 | 内容 |
|-----|------|------|
| v1.0 | 2026-07-23 | R 角单形状调整 + 锁定（shape.tags 持久化）+ 防误触 + 预设库 + 样式刷 + 历史 |
| v1.1 | 2026-07-23 | pipette 样式刷增强 + history 简化（纯内存）|
| v1.2 | 2026-07-25 | **布局模式 + 交互层 verified**（含 v1.2.0 → v1.3.5 全部 hotfix）|

---

## 核心功能矩阵

| 功能 | 实现 | 持久化 | 单测覆盖 | PPT 验证 |
|------|------|--------|---------|---------|
| 单形状 R 角调整（cm / %）| ✅ | n/a | ✅ | ✅ |
| 批量 R 角调整 | ✅ | n/a | ✅ | ✅ |
| 防误触（strict） | ✅ | shape.tag | ✅ | ✅ |
| 锁定 R 角绝对值 + 自动按比例 | ✅ | shape.tag | ✅ | ✅ |
| 锁定 toggle disable | ✅ | shape.tag | ✅ | ✅ v1.3.5 fix |
| 预设库 | ✅ | n/a | ❌ | ✅ |
| 样式刷（pipette） | ✅ | n/a | ⚠️ 部分 | ❌ feature bug |
| 历史（5 次） | ✅ | 内存 | ❌ | ✅ |
| 锁定状态下拖改大小反算 | ✅ | n/a | ✅ | ✅ |
| **布局模式（rows×cols + padding/gutter）** | ✅ v1.2 | shape.tag × 2 | ✅ | ✅ |
| **布局 R 角联动** | ⚠️ partial | n/a | ✅ | ❌ feature bug |

---

## 三层架构（v1.2 重构成果）

```
┌──────────────────────────────────────────────────────────────┐
│  dialog.js (UI 层) — ~2500 行                                 │
│  事件绑定 / 渲染 / toast / debug log                          │
│  极薄，每个 handler 5-50 行                                   │
└──────────────────────────────────────────────────────────────┘
                              │ 调用
                              ▼
┌──────────────────────────────────────────────────────────────┐
│  src/lib/radius-core.js (实现层) — ~880 行                     │
│  8 个 driver 版函数：writeRadius / readLockState /            │
│  writeLockState / reapplyLock / applyLayout /                │
│  syncLayoutChildrenR / writeRadiusToShapePure /              │
│  applyLayoutPure                                             │
│  零 Office.js 调用 → 100% 单元测试                            │
└──────────────────────────────────────────────────────────────┘
                              │ 调用
                              ▼
┌──────────────────────────────────────────────────────────────┐
│  src/lib/ppt-driver.js (交互层) — 109 行                       │
│  16 个方法：load / sync / selectedShapes / activeSlide /     │
│  slideShapes / shapeId / size / box / isRoundRect /          │
│  adjFraction / loadAdjValue / setBox / setAdjFraction /      │
│  addTag / deleteTag / readTag                                │
│  零业务逻辑 → 100% 单元测试                                   │
└──────────────────────────────────────────────────────────────┘
                              │ 调
                              ▼
┌──────────────────────────────────────────────────────────────┐
│  Office.js + PowerPoint (Mac LTSC 16.111)                    │
└──────────────────────────────────────────────────────────────┘
```

### 关键约束
- **driver 不知道任何业务概念** —— 不认识 `LOCK_TAG_KEY` / `LAYOUT_PARENT_TAG_KEY`，不知道 strict 是什么
- **radius-core 不 import Office.js** —— 所有形状读/写/load/sync 走 driver
- **dialog.js 是搬运工** —— `onClick → 开 driver → 调 feature → 渲染结果`

---

## 测试状态

```
112 passed, 0 failed
├─ test-radius-core.js         103 个（纯算法）
├─ test-mock-harness.js         70 个（mock PowerPoint run 上下文）
└─ test-driver-integration.js  109 个（mock driver + radius-core 集成）
```

### 跑测试

```bash
cd /Users/ma/Documents/minimax/radius_in_ppt
npm test                                            # 跑全部 3 个测试文件
node test/test-radius-core.js                       # 仅算法
node test/test-mock-harness.js                      # 仅 mock harness
node test/test-driver-integration.js                # 仅 driver 集成
```

### 烟囱测试（PPT 内）

任务窗格 → 点「🧪 Driver 烟囱测试」按钮 → 14/14 全过即 driver verified。

---

## 已知 Bug / 待修

| # | 优先级 | 描述 | 归属 |
|---|--------|------|------|
| 1 | P2 | pipette 吸取后无法刷入任何形状（rect / roundRect 都失败）| feature bug，Step 4 修 |
| 2 | P2 | 布局 R 角联动失败：改父 R 角，子 R 角没变化 | feature bug，Step 3c 修 |
| 3 | P3 | lockMonitor 偶发 `GeneralException`（try/catch 兜住，无影响） | Step 5 重构时清 |
| 4 | P3 | 调试 log 还开着（`[applyLayout/driver]` 等）| Step 5 清 |
| 5 | P4 | driver 烟囱测试 step 6 setAdjFraction 总走「跨 run 兜底」路径 | Mac LTSC 限制，已知 |

---

## 下一步（Step 3c → 4 → 5）

### Step 3c — layout tag 读写迁移 + stale state 检测
- `radius-core.loadLayoutTags(driver, slide)` —— 读当前 slide 所有 layout tag
- `radius-core.saveLayoutTags(driver, parentShape, childIds, params)` —— 写父 tag
- `refreshSelection` 加 stale state 检测（中间页删除的子形状 → 自动 unlink）

### Step 4 — pipette + history 迁移
- 抽 `radius-core.pickupFromSelection` / `applyPickedToSelection` / `pushHistory` 为 driver 版
- 修 pipette 刷入 bug
- 删 dialog.js 旧 pipette + history 逻辑

### Step 5 — dialog.js UI 层重构
- `lockMonitor` 改成 `radius-core.monitorTick(driver, ...)` 纯函数
- 清掉所有调试 log
- dialog.js 缩到 ~500 行（删 2000 行混合逻辑）

**所有上述步骤不再 PPT 实测**，走 `npm test` + 代码 review。

---

## 项目结构

```
radius_in_ppt/
├── src/
│   ├── dialog/
│   │   ├── dialog.html
│   │   ├── dialog.css
│   │   └── dialog.js                  # ~2500 行（迁移后目标 ~500）
│   └── lib/
│       ├── radius-core.js             # ~880 行实现层
│       └── ppt-driver.js              # 109 行交互层
├── test/
│   ├── test-radius-core.js            # 103 算法
│   ├── test-mock-harness.js           # 70 mock harness
│   ├── test-driver-integration.js     # 109 driver 集成
│   └── README.md
├── tools/
│   ├── build-app.sh                   # 打包 .app
│   ├── build-and-deploy.sh            # 一键 build + 部署 + git commit
│   ├── serve.js                       # ~60 行静态 server
│   ├── build-dmg.sh                   # 打包 .dmg
│   └── sign-and-notarize.sh           # 公证（待 Apple Developer 账号）
├── app/MacOS/RadiusInPpt              # bash 启动器
├── assets/                            # 图标
├── dist/                              # build 输出
├── AGENTS.md                          # 必读：架构 + Mac LTSC 坑
├── PROGRESS.md                        # 本文件：进度总览
├── changelogs/
│   ├── 2026-07-23.md                  # v1.0
│   ├── 2026-07-23-v1.1.md             # v1.1
│   └── 2026-07-24-v1.2.md             # v1.2 + v1.2.0 → v1.3.5 全 hotfix
├── manifest.xml                       # version 1.3.5
├── package.json                       # npm test
└── test.pptx                          # 测试用 .pptx
```

---

## 关键设计决策

### 1. shape.tags 持久化（Mac LTSC 唯一 work 的方案）
- `customProperties` / `customXmlParts` 在 Mac LTSC task pane 都不可用
- shape.tags 直接挂 OOXML `<p:tagLst>` 段，跟 .pptx 文件走
- 保存 .pptx → 关 PPT → 重开 → tag 还在
- 跨设备 / 发文件 / 换机器都保留

### 2. Driver 层不 throw
- `driver.adjFraction` 内部 try/catch 返回 0（defensive）
- driver API 契约：永不 throw，调用方不需要保护
- 这是 v1.2.5 实测决定的（轮询时一个 shape 失败不能带垮整个 monitor）

### 3. set+read 必须 fresh get(0) AFTER sync
- Mac LTSC proxy 是 snapshot 风格，set 之后旧 proxy 不会 reload value
- v1.2.7 / v1.2.8 / v1.2.9 三次修法都不稳
- v1.3.1 真修法：同 run set+read 不可靠 → 跨 PowerPoint.run 兜底

### 4. 防误触 = 最高优先级
- strict tag = "1" 的形状，任何 R 角写入路径都不能跳过
- layout apply / pipette / 联动 hook 都不行
- 两道防线：内存层 + PPT 层（防 race）

### 5. 未来 feature 信任单元测试
- driver 16 方法 v1.3.4 烟囱测试 14/14 ✅
- 112 个单测 v1.3.5 全过 ✅
- 7 场景 PPT 验证 v1.3.5 通过（除 feature bug 外）✅
- → 未来 step 3c/4/5 **不再 PPT 实测**
