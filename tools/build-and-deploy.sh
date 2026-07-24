#!/bin/bash
#
# build-and-deploy.sh — 一键 build + 部署到 PowerPoint wef 路径
#
# 用法：
#   bash tools/build-and-deploy.sh 1.3.1           # bump 到 1.3.1 + build + 部署
#   bash tools/build-and-deploy.sh 1.3.1 "fix XXX" # bump 到 1.3.1 + 自定义 commit message
#
# 流程：
#   1. 备份当前 manifest.xml
#   2. bump Version 和 cache buster ?v=v1.x.y
#   3. bash tools/build-app.sh（重建 .app）
#   4. 拷 manifest.xml 到 ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
#   5. 备份的 manifest.xml 删掉
#
# 用完 user 双击 dist/RadiusInPpt.app 重启 → PowerPoint Cmd+Q → 重开 → 新版本加载

set -eo pipefail

cd "$(dirname "$0")/.."
ROOT="$(pwd)"

NEW_VERSION="${1:-}"
COMMIT_MSG="${2:-}"

if [ -z "$NEW_VERSION" ]; then
  CURRENT_VERSION=$(grep -oE '<Version>[^<]+</Version>' manifest.xml | sed -E 's|</?Version>||g')
  echo "[error] 用法: $0 <new_version> [commit_msg]"
  echo "        当前版本: $CURRENT_VERSION"
  echo "        例如: $0 1.3.1 'fix R 角边界'"
  exit 1
fi

echo "[deploy] ============================================"
echo "[deploy] 新版本: $NEW_VERSION"
echo "[deploy] ============================================"

# 1. 备份当前 manifest.xml（万一 bump 失败能回滚）
cp manifest.xml manifest.xml.bak

# 2. bump Version
sed -i '' "s|<Version>[^<]*</Version>|<Version>${NEW_VERSION}</Version>|" manifest.xml

# 3. bump cache buster ?v=v1.x.y → ?v=NEW_VERSION
sed -i '' "s|?v=v[0-9.]*|?v=${NEW_VERSION}|g" manifest.xml

echo "[deploy] manifest.xml updated:"
grep -E '<Version>|SourceLocation' manifest.xml | head -2

# 4. build .app
echo "[deploy] running build-app.sh..."
bash tools/build-app.sh 2>&1 | tail -5

# 5. 拷 manifest.xml 到 wef
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
WEF_MANIFEST="$WEF_DIR/manifest.xml"
if [ -d "$WEF_DIR" ]; then
  cp -f dist/RadiusInPpt.app/Contents/Resources/manifest.xml "$WEF_MANIFEST"
  echo "[deploy] manifest copied to: $WEF_MANIFEST"
  echo "[deploy] deployed version: $(grep -oE '<Version>[^<]+</Version>' "$WEF_MANIFEST" | sed -E 's|</?Version>||g')"
else
  echo "[deploy] WARNING: $WEF_DIR not found (PowerPoint 可能没装 or 没跑过)"
  echo "[deploy]         .app 会自己拷，但用户需要先双击一次 .app"
fi

# 6. 备份删掉
rm -f manifest.xml.bak

echo "[deploy] ============================================"
echo "[deploy] ✅ v${NEW_VERSION} deployed"
echo "[deploy]    下一步：双击 dist/RadiusInPpt.app"
echo "[deploy]    然后 Cmd+Q PowerPoint → 重开 → 新版本"
echo "[deploy] ============================================"

# 7. 可选：git commit
if [ -n "$COMMIT_MSG" ]; then
  echo "[deploy] git commit..."
  git add -A
  git commit -m "$COMMIT_MSG" 2>&1 | tail -3 || echo "[deploy] git commit failed (no git or no remote?)"
fi
