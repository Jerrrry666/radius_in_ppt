# Security Policy

> 🌍 **[English version](./SECURITY.en.md)**

## 报告漏洞

如果你在 R 角调整里发现了安全问题，**请不要在 GitHub Issues 公开提**。

请发邮件到 **security 联系方式**：见项目主页 / GitHub profile。

邮件里请包含：
- 问题描述
- 复现步骤
- 潜在影响（你判断的）
- 你的环境（PowerPoint 版本 / macOS 版本）

我会在 **48 小时内**回复确认，并在合理时间内修复。

## 支持的版本

| 版本 | 支持 |
| --- | --- |
| v1.3.0+ | ✅ 活跃支持 |
| v1.2.x | ⚠️ 仅严重安全修复 |
| v1.0 / v1.1 | ❌ 不再支持 |

## 范围说明

这是一个**纯本地**的 PowerPoint 加载项：
- 没有服务端
- 没有数据收集
- 不联网（除了 PowerPoint Add-in framework 本身）
- 不访问网络资源

所以常见的安全风险（数据泄露、远程攻击）**理论上不适用**。但以下仍然可能：
- Manifest 注入 / 篡改（如果你手动编辑了 manifest.xml 并用非标准方式加载）
- 持久化数据被恶意修改（`shape.tags` 里的 JSON）
- Office Add-in 框架本身的漏洞

## 安全实践

用户侧：
- 只从 [GitHub Releases](https://github.com/Jerrrry666/radius_in_ppt/releases) 下载 `.app`
- 验证下载来源（不要从第三方站点拿）
- Gatekeeper 首次打开提示是正常的，**右键 → 打开** 即可

开发者侧：
- 不要 commit `*.pem` / `*.key` / `certs/`
- 不要在 source code 里 hardcode token / API key
- PR review 时关注 input validation（虽然本地加载项风险低，但好习惯）

## 致谢

负责任地披露的漏洞会在修复后致谢（如果你同意）。
