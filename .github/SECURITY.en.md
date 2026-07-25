# Security Policy

> 🌍 **[中文版](./SECURITY.md)**

## Reporting a Vulnerability

If you discover a security issue in RadiusInPpt, **please do not open a public GitHub Issue**.

Send an email to the **security contact** listed on the project homepage / GitHub profile.

Please include:
- Description of the issue
- Reproduction steps
- Potential impact (your assessment)
- Your environment (PowerPoint version / macOS version)

You will receive an acknowledgement within **48 hours**, and a fix within a reasonable timeframe.

## Supported Versions

| Version | Support |
| --- | --- |
| v1.3.0+ | ✅ Active support |
| v1.2.x | ⚠️ Critical security fixes only |
| v1.0 / v1.1 | ❌ End of life |

## Scope

This is a **fully local** PowerPoint add-in:
- No server backend
- No data collection
- No network access (beyond the Office Add-in framework itself)
- No external resources

So common security risks (data exfiltration, remote attacks) are **theoretically not applicable**. Possible concerns remain for:
- Manifest injection / tampering (if you manually edit `manifest.xml` and load it via non-standard means)
- Malicious modification of persisted data (JSON in `shape.tags`)
- Vulnerabilities in the Office Add-in framework itself

## Security Practices

User side:
- Only download the `.app` from [GitHub Releases](https://github.com/Jerrrry666/radius_in_ppt/releases)
- Verify the download source (do not grab it from third-party sites)
- The Gatekeeper "cannot be verified" prompt on first open is normal — **right-click → Open** to bypass it once

Developer side:
- Do not commit `*.pem` / `*.key` / `certs/`
- Do not hardcode tokens / API keys in source
- Review PRs for input validation (low risk for a local add-in, but good habit)

## Acknowledgements

Vulnerabilities disclosed responsibly will be credited after the fix is released (with your consent).
