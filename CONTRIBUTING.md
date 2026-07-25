# 贡献指南

> 🌍 **[English version](./CONTRIBUTING.en.md)**

感谢考虑为 **R 角调整** 添砖加瓦！🎉

这是一个 macOS PowerPoint 加载项，让用户用 **厘米** 或 **百分比** 精确设置圆角矩形的 R 角。

## 提 issue

- **Bug 报告** — 用 [GitHub Issues](https://github.com/Jerrrry666/radius_in_ppt/issues/new)，标题简短描述现象
- **功能建议** — 同上，但前面加个 `[Feature Request]`
- **问答 / 讨论** — 用 [GitHub Discussions](https://github.com/Jerrrry666/radius_in_ppt/discussions)（如果有开的话）

提 issue 前先搜一下有没有重复的。

## 提 PR

### 1. Fork + 开分支

```bash
# fork 之后
git clone https://github.com/<你的用户名>/radius_in_ppt.git
cd radius_in_ppt
git checkout -b feat/your-feature-name
```

### 2. 开发约定

- **三层架构** — `dialog.js`（UI）→ `radius-core.js`（算法）→ `ppt-driver.js`（Office.js 交互）。改动落哪一层先想清楚。
- **driver 不知道任何业务概念**（不认 `LOCK_TAG_KEY` / `LAYOUT_PARENT_TAG_KEY`）。
- **radius-core 不 import Office.js**。所有 Office 调用走 driver。
- **AGENTS.md** 是 AI agent 必读，里面有 Mac LTSC Office.js 踩坑清单。人类读者也建议扫一眼 §1-§2。

### 3. 跑测试

```bash
npm test
```

应该看到 210/0 全过（或更新后更多）。新功能必须带单测。

### 4. Commit message

- 用英文，**避免中文标点**（commit message 在 bash pipeline 里会被解析）
- 格式：`<scope>: <what changed>`
  - `feat: 新增 xxx`
  - `fix: 修复 xxx`
  - `refactor: 重构 xxx`
  - `docs: 文档更新`
  - `test: 加测试`
- 一个 commit 做一件事

### 5. 提 PR

- 标题清晰描述改动
- 正文里：
  - **What** — 改了什么
  - **Why** — 为什么改（解决了什么 issue / 场景）
  - **How to test** — 怎么验证
  - **Screenshots** — 如果有 UI 改动
- 关联相关 issue（`Fixes #123`）

## 第一次 setup（开发者）

```bash
git clone https://github.com/Jerrrry666/radius_in_ppt.git
cd radius_in_ppt
npm start   # 启动本地 HTTP server (localhost:3000)
```

然后把 manifest 注册到 PowerPoint：

```bash
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF" && cp manifest.xml "$WEF/manifest.xml"
```

完全退出 PowerPoint（`Cmd+Q`），重新打开，顶部 ribbon 会出现「R 角调整」Tab。

## 跑 build

```bash
bash tools/build-app.sh   # 生成 dist/RadiusInPpt.app
```

## 沟通风格

- 直接、友好、就事论事
- 提 issue / PR 不需要客套
- 代码 review 以事实为准，不搞权威崇拜

## License

提交 PR 即表示你同意按 [MIT License](./LICENSE) 授权你的贡献。
