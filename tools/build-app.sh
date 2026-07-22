#!/bin/bash
#
# build-app.sh — 把项目打包成 macOS .app
#
# 产物：./dist/RadiusInPpt.app
# 可直接拖入 /Applications/，或运行 `npm run dist`
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
# 清掉旧的 .app / .dmg / 临时 bak 目录；保留 $DIST
if [ -d "$DIST" ]; then
  find "$DIST" -mindepth 1 -maxdepth 1 \( -name '*.app' -o -name '*.dmg' -o -name 'dmg-staging' -o -name 'AppIcon.iconset' -o -name '*.bak.*' -o -name 'dist 2' -o -name '*~' \) -exec rm -rf {} +
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
  <string>R 角调整需要激活 PowerPoint 来读取和修改你选中的圆角矩形。</string>
</dict>
</plist>
EOF

# 2. PkgInfo
printf 'APPL????' > "$CONTENTS/PkgInfo"

# 3. 启动脚本
echo "[build] copying launcher"
cp "$ROOT/app/MacOS/$APP_NAME" "$MACOS_DIR/$APP_NAME"
chmod +x "$MACOS_DIR/$APP_NAME"

# 4. 拷贝源文件到 Resources
echo "[build] copying resources"
cp -R "$ROOT/src" "$RES_DIR/"
cp -R "$ROOT/assets" "$RES_DIR/"
cp "$ROOT/manifest.xml" "$RES_DIR/"

# 不需要拷 tools/serve.js 之外的（package.json 等用户用不到）
mkdir -p "$RES_DIR/tools"
cp "$ROOT/tools/serve.js" "$RES_DIR/tools/"

# 5. 生成 .icns
echo "[build] generating AppIcon.icns"
mkdir -p "$ICONSET_DIR"
SRC_PNG="$ROOT/assets/icon-128.png"  # 实际 128x128

# 用 sips 生成各尺寸
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
echo "  2. 首次会在 PowerPoint 中：插入 → 我的加载项 → 开发人员加载项 → 从文件添加 → 选 manifest.xml"
echo "  3. 之后每次用直接双击 .app 即可"
echo
echo "如需停止后台 server：lsof -ti tcp:3000 | xargs kill"
