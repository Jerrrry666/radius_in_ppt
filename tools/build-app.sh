#!/bin/bash
#
# build-app.sh — 把项目打包成 macOS .app（Office Add-in 路线）
#
# 产物：./dist/RadiusInPpt.app
# 入口：双击 .app → 启动 bash 脚本
#   1. 启动 http://localhost:3000 server（serve.js）
#   2. 复制 manifest.xml 到 ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
#   3. 弹引导框：完全退出 PowerPoint (Cmd+Q) → 重新打开
#      → PowerPoint 顶部出现「R 角调整」tab
#
set -eo pipefail

cd "$(dirname "$0")/.."
ROOT="$(pwd)"

APP_NAME="RadiusInPpt"
DISPLAY_NAME="R 角调整"
BUNDLE_ID="com.jerrrry666.radiusinppt"
VERSION="1.0.0"

DIST="$ROOT/dist"
APP="$DIST/$APP_NAME.app"
CONTENTS="$APP/Contents"
MACOS_DIR="$CONTENTS/MacOS"
RES_DIR="$CONTENTS/Resources"
ICONSET_DIR="$DIST/AppIcon.iconset"

echo "[build] cleaning $DIST"
if [ -d "$DIST" ]; then
  find "$DIST" -mindepth 1 -maxdepth 1 \( \
    -name '*.app' -o -name '*.dmg' -o -name 'dmg-staging' -o -name 'AppIcon.iconset' \
    -o -name '*.bak.*' -o -name 'dist 2' -o -name '*~' \
  \) -exec rm -rf {} +
fi
mkdir -p "$DIST"

echo "[build] creating .app bundle structure"
mkdir -p "$MACOS_DIR" "$RES_DIR"

# 1. Info.plist
echo "[build] writing Info.plist"
cat > "$CONTENTS/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleName</key><string>$APP_NAME</string>
  <key>CFBundleDisplayName</key><string>$DISPLAY_NAME</string>
  <key>CFBundleExecutable</key><string>$APP_NAME</string>
  <key>CFBundleIdentifier</key><string>$BUNDLE_ID</string>
  <key>CFBundleVersion</key><string>$VERSION</string>
  <key>CFBundleShortVersionString</key><string>$VERSION</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleSignature</key><string>????</string>
  <key>LSMinimumSystemVersion</key><string>10.15</string>
  <key>NSHighResolutionCapable</key><true/>
  <key>NSPrincipalClass</key><string>NSApplication</string>
  <key>CFBundleIconFile</key><string>AppIcon</string>
  <key>NSAppleEventsUsageDescription</key>
  <string>R 角调整需要控制 PowerPoint（关闭后重新打开）以加载 R 角调整加载项。</string>
</dict>
</plist>
EOF

# 2. PkgInfo
printf 'APPL????' > "$CONTENTS/PkgInfo"

# 3. 启动脚本（bash）
echo "[build] copying launcher"
cp "$ROOT/app/MacOS/$APP_NAME" "$MACOS_DIR/$APP_NAME"
chmod +x "$MACOS_DIR/$APP_NAME"

# 4. 拷贝 Office Add-in 资源
echo "[build] copying Office Add-in resources"
cp -R "$ROOT/src" "$RES_DIR/"
cp -R "$ROOT/assets" "$RES_DIR/"
cp "$ROOT/manifest.xml" "$RES_DIR/"
mkdir -p "$RES_DIR/tools"
cp "$ROOT/tools/serve.js" "$RES_DIR/tools/"
# 注意：jszip.min.js / package.json / node_modules 已删除（v2.0 起不再需要）

# 5. 生成 .icns
echo "[build] generating AppIcon.icns"
mkdir -p "$ICONSET_DIR"
SRC_PNG="$ROOT/assets/icon-128.png"
gen() {
  local size=$1
  local name=$2
  sips -z "$size" "$size" "$SRC_PNG" --out "$ICONSET_DIR/$name" >/dev/null
}
gen 16   "icon_16x16.png"
gen 32   "icon_16x16@2x.png"
gen 32   "icon_32x32.png"
gen 64   "icon_32x32@2x.png"
gen 128  "icon_128x128.png"
gen 256  "icon_128x128@2x.png"
gen 256  "icon_256x256.png"
gen 512  "icon_256x256@2x.png"
gen 512  "icon_512x512.png"
gen 1024 "icon_512x512@2x.png"
iconutil -c icns "$ICONSET_DIR" -o "$RES_DIR/AppIcon.icns"
rm -rf "$ICONSET_DIR"

echo "[build] done: $APP"
echo
echo "用法："
echo "  1. 双击 $APP（或拖入 /Applications/）"
echo "  2. 完全退出 PowerPoint（Cmd + Q）"
echo "  3. 重新打开 PowerPoint，顶部 ribbon 出现「R 角调整」tab"
echo
echo "⚠️  首次双击未签名 .app 会被 Gatekeeper 拦："
echo "   右键 $APP → 打开 → 弹窗里点「打开」（仅一次）"
