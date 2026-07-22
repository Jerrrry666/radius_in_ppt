#!/bin/bash
#
# build-app.sh — 把项目打包成 macOS .app
#
# 产物：./dist/RadiusInPpt.app
# 入口：双击 .app → 启动 Swift 菜单栏 app（NSStatusItem）
#       顶部菜单栏出现图标，点开有菜单
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

# 1. Info.plist（LSUIElement=true → 不在 Dock 显示）
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
  <key>LSMinimumSystemVersion</key><string>11.0</string>
  <key>LSUIElement</key><true/>
  <key>NSHighResolutionCapable</key><true/>
  <key>NSPrincipalClass</key><string>NSApplication</string>
  <key>CFBundleIconFile</key><string>AppIcon</string>
  <key>NSAppleEventsUsageDescription</key>
  <string>R 角调整需要控制 PowerPoint 来读取和修改你选中的圆角矩形。</string>
</dict>
</plist>
EOF

# 2. PkgInfo
printf 'APPL????' > "$CONTENTS/PkgInfo"

# 3. 编译 Swift 主程序
echo "[build] compiling Swift"
swiftc -O \
  -target arm64-apple-macosx11.0 \
  -o "$MACOS_DIR/$APP_NAME" \
  "$ROOT/menubar/main.swift" 2>&1
chmod +x "$MACOS_DIR/$APP_NAME"

# 4. 菜单栏图标（template image）
echo "[build] copying menubar icon"
cp "$ROOT/menubar/menubar-icon.png" "$RES_DIR/menubar-icon.png"
cp "$ROOT/menubar/menubar-icon@2x.png" "$RES_DIR/menubar-icon@2x.png" 2>/dev/null || true

# 5. 生成 .icns（用于 Finder 图标）
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
echo "  2. macOS 顶部菜单栏会出现 R 角图标"
echo "  3. 点图标 → 选「调整 R 角...」输入厘米值"
echo
echo "⚠️  首次双击未签名 .app 会被 Gatekeeper 拦："
echo "   右键 $APP → 打开 → 弹窗里点「打开」（仅一次）"
