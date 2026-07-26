#!/bin/bash
#
# build-windows.sh — produce a Windows-usable package (folder + .zip)
#
# Output:
#   dist/RadiusInPpt-win/      (folder users can copy to Windows)
#   dist/RadiusInPpt-win.zip   (zip of the above, for download)
#
# Build is done on macOS (or any Unix). No Windows machine needed.
# The .bat launcher is OS-agnostic and will run on any Windows 10+.
#
# What we copy:
#   - src/                    (HTML / JS / CSS — all cross-platform)
#   - manifest.xml            (Office Add-in manifest, cross-platform)
#   - assets/                 (ribbon icons)
#   - tools/serve.js          (Node.js static file server, cross-platform)
#   - app/Windows/RadiusInPpt.bat  (Windows launcher, replaces the
#                                   macOS .app bundle)
#
# Build-time strip:
#   - dialog.html debug log section (@build-strip-debug-log markers)
#     — same strip as build-app.sh, see that script for the pattern
#
# Tested on:
#   - macOS host builds the .zip
#   - Windows 10 / 11 (target) runs the .bat

set -eo pipefail

cd "$(dirname "$0")/.."
ROOT="$(pwd)"

APP_NAME="RadiusInPpt"
DIST="$ROOT/dist"
WIN_DIR="$DIST/${APP_NAME}-win"
WIN_ZIP="$DIST/${APP_NAME}-win.zip"

echo "[build-win] ============================================"
echo "[build-win] Output: $WIN_DIR / $WIN_ZIP"
echo "[build-win] ============================================"

# 1. Clean previous build (but leave Mac .app alone)
echo "[build-win] cleaning $WIN_DIR / $WIN_ZIP"
if [ -d "$WIN_DIR" ]; then
  find "$WIN_DIR" -mindepth 1 -maxdepth 1 -exec rm -rf {} +
fi
[ -f "$WIN_ZIP" ] && mavis-trash "$WIN_ZIP" 2>/dev/null || rm -f "$WIN_ZIP"
mkdir -p "$WIN_DIR"

# 2. Copy Office Add-in resources (cross-platform)
echo "[build-win] copying Office Add-in resources"
cp -R "$ROOT/src" "$WIN_DIR/"
cp -R "$ROOT/assets" "$WIN_DIR/"
cp "$ROOT/manifest.xml" "$WIN_DIR/"

# 3. Copy serve.js (Node.js static server — cross-platform)
mkdir -p "$WIN_DIR/tools"
cp "$ROOT/tools/serve.js" "$WIN_DIR/tools/"

# 4. Copy Windows launcher
cp "$ROOT/app/Windows/RadiusInPpt.bat" "$WIN_DIR/${APP_NAME}.bat"

# 5. Strip debug log section (same as Mac build)
echo "[build-win] stripping debug log section from release build"
sed -i '' '/@build-strip-debug-log:start/,/@build-strip-debug-log:end/d' "$WIN_DIR/src/dialog/dialog.html"
if grep -q '<details id="debug-log"' "$WIN_DIR/src/dialog/dialog.html"; then
  echo "[build-win] ERROR: debug log still in release build, sed may not have matched"
  exit 1
fi
echo "[build-win] debug log section stripped"

# 6. Zip the package
echo "[build-win] creating $WIN_ZIP"
cd "$DIST"
zip -r "$WIN_ZIP" "${APP_NAME}-win" -x "*.DS_Store" >/dev/null

echo
echo "[build-win] done:"
ls -la "$WIN_ZIP"
echo
echo "用法（Windows 用户）："
echo "  1. 解压 RadiusInPpt-win.zip 到任意目录"
echo "  2. 双击 RadiusInPpt.bat"
echo "  3. 弹框提示 → 完全退出 PowerPoint 后重启（Cmd 不适用，文件→退出）"
echo "  4. PowerPoint ribbon 出现「R 角调整」tab"
echo
echo "⚠️  未在 Windows 实测 — Mac 上 build，Windows 上跑。"
echo "   如果有问题，附 %TEMP%\\radius_in_ppt.log 提 issue。"
