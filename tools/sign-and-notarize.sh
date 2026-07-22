#!/bin/bash
#
# sign-and-notarize.sh — 给 .app 签名 + 公证，让 Gatekeeper 不再拦
#
# 前置条件：
#   1. Apple Developer Program 会员（$99/年）：https://developer.apple.com/programs/enroll/
#   2. 在 Xcode 里创建 "Developer ID Application" 证书
#      （Xcode → Settings → Accounts → 选 Apple ID → Manage Certificates → + → Developer ID Application）
#   3. 一次性存储 notarytool 凭据：
#      xcrun notarytool store-credentials "AC_PROFILE" \
#        --apple-id "you@example.com" \
#        --password "abcd-efgh-ijkl-mnop"   ← appleid.apple.com 生成的 app-specific password
#        --team-id "ABCDE12345"
#
# 用法：
#   export AC_SIGNING_IDENTITY="Developer ID Application: Your Name (ABCDE12345)"
#   export AC_NOTARY_PROFILE="AC_PROFILE"
#   bash tools/sign-and-notarize.sh
#
# 不传环境变量会走交互输入（适合临时用一次）
#
set -eo pipefail

cd "$(dirname "$0")/.."
ROOT="$(pwd)"
APP="$ROOT/dist/RadiusInPpt.app"

if [ ! -d "$APP" ]; then
  echo "[sign] .app 不存在，先跑 build-app.sh"
  bash "$ROOT/tools/build-app.sh"
fi

# ---------------- 1. 拿 signing identity ----------------
if [ -z "${AC_SIGNING_IDENTITY:-}" ]; then
  echo "[sign] 当前可用的 signing identities:"
  security find-identity -v -p codesigning 2>&1 | sed 's/^/    /' || true
  if [ "$(security find-identity -p codesigning 2>/dev/null | wc -l)" = "0" ]; then
    echo ""
    echo "❌ 系统里没有任何代码签名证书。请先："
    echo "   1. 注册 Apple Developer Program（\$99/年）"
    echo "   2. 在 Xcode → Settings → Accounts → Manage Certificates 创建 'Developer ID Application' 证书"
    echo "   3. 重跑这个脚本"
    exit 1
  fi
  echo ""
  read -p "[sign] 输入 signing identity（如 'Developer ID Application: Your Name (TEAMID)'）: " AC_SIGNING_IDENTITY
fi
echo "[sign] identity = $AC_SIGNING_IDENTITY"

# ---------------- 2. 签名 ----------------
echo "[sign] codesign --deep --force --options=runtime --timestamp ..."
codesign --deep --force --options=runtime --timestamp \
  --sign "$AC_SIGNING_IDENTITY" \
  "$APP"

# 验证签名
echo "[sign] verify codesign:"
codesign --verify --deep --strict --verbose=2 "$APP" 2>&1 | sed 's/^/    /'

# ---------------- 3. 公证 ----------------
if [ -z "${AC_NOTARY_PROFILE:-}" ]; then
  echo ""
  echo "[notary] 没设 AC_NOTARY_PROFILE 环境变量，交互输入 notarytool profile 名"
  read -p "[notary] notarytool profile 名（store-credentials 时设的第一个参数）: " AC_NOTARY_PROFILE
fi
echo "[notary] profile = $AC_NOTARY_PROFILE"
echo "[notary] submitting to Apple notary service（可能要 1-5 分钟）..."

# 先打成 zip（notarytool 对 .app 直接 submit 也行，但 zip 更稳）
NOTARY_ZIP="$ROOT/dist/RadiusInPpt-for-notarization.zip"
ditto -c -k --sequesterRsrc --keepParent "$APP" "$NOTARY_ZIP"

xcrun notarytool submit "$NOTARY_ZIP" \
  --keychain-profile "$AC_NOTARY_PROFILE" \
  --wait

# ---------------- 4. Staple ticket ----------------
echo "[notary] stapling ticket to .app..."
xcrun stapler staple "$APP"
xcrun stapler validate "$APP"

# 清理
rm -f "$NOTARY_ZIP"

# ---------------- 5. 最终验证 ----------------
echo ""
echo "[verify] spctl（Gatekeeper 评估）:"
spctl --assess --type execute --verbose=2 "$APP" 2>&1 | sed 's/^/    /'

echo ""
echo "[verify] codesign 信息:"
codesign -dvv "$APP" 2>&1 | sed 's/^/    /'

echo ""
echo "✅ done. 现在的 $APP 可以直接双击打开，Gatekeeper 不会再拦。"
echo "   建议再用 hdiutil 打成 .dmg 分发：bash tools/build-dmg.sh"
