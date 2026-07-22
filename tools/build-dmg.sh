#!/bin/bash
#
# build-dmg.sh — 把 .app 打包成 .dmg 用于分发
#
# 依赖：hdiutil（macOS 自带）、npm run build:app
# 产物：./dist/RadiusInPpt-1.0.0.dmg
#
set -eo pipefail

cd "$(dirname "$0")/.."
ROOT="$(pwd)"

if [ ! -d "$ROOT/dist/RadiusInPpt.app" ]; then
  echo "[dmg] .app 不存在，先跑 npm run build:app"
  bash "$ROOT/tools/build-app.sh"
fi

APP="$ROOT/dist/RadiusInPpt.app"
DMG_DIR="$ROOT/dist/dmg-staging"
DMG_PATH="$ROOT/dist/RadiusInPpt-1.0.0.dmg"

echo "[dmg] staging: $DMG_DIR"
rm -rf "$DMG_DIR" "$DMG_DIR.bak.*" 2>/dev/null || true
mkdir -p "$DMG_DIR"
cp -R "$APP" "$DMG_DIR/"

# 创建 Applications 的软链（拖入安装的经典操作）
ln -sf /Applications "$DMG_DIR/Applications"

echo "[dmg] creating: $DMG_PATH"
rm -f "$DMG_PATH" "$DMG_PATH.bak.*" 2>/dev/null || true

# -fs HFS+ 兼容老 Mac；-ov 不覆盖
hdiutil create \
  -volname "R 角调整" \
  -srcfolder "$DMG_DIR" \
  -ov \
  -format UDZO \
  "$DMG_PATH" 2>&1 | tail -3

# 清理
rm -rf "$DMG_DIR" 2>/dev/null || true

echo "[dmg] done: $DMG_PATH"
echo "分发给其他用户：双击 dmg → 拖入 /Applications 即可"
