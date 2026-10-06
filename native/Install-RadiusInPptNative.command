#!/bin/bash
# One-time file preparation. PowerPoint owns registration and future loading.
set -euo pipefail

PACKAGE_DIR="$(cd "$(dirname "$0")" && pwd)"
PAYLOAD="$PACKAGE_DIR/RadiusInPptNative.ppam"
DESTINATION="$HOME/Library/Application Support/RadiusInPptNative"
REVEAL=1

while [ "$#" -gt 0 ]; do
  case "$1" in
    --destination)
      [ "$#" -ge 2 ] || { echo 'Missing destination directory.' >&2; exit 2; }
      DESTINATION="$2"
      shift 2
      ;;
    --no-reveal) REVEAL=0; shift ;;
    *) echo "Unknown option: $1" >&2; exit 2 ;;
  esac
done

if [ ! -f "$PAYLOAD" ]; then
  echo '请先解压完整插件包，将此文件与 RadiusInPptNative.ppam 放在同一文件夹。' >&2
  exit 1
fi
/usr/bin/unzip -tqq "$PAYLOAD"
mkdir -p "$DESTINATION"
TARGET="$DESTINATION/RadiusInPptNative.ppam"

if [ ! -f "$TARGET" ] || ! /usr/bin/cmp -s "$PAYLOAD" "$TARGET"; then
  if [ -f "$TARGET" ]; then
    /bin/cp -p "$TARGET" "$DESTINATION/RadiusInPptNative.previous.ppam"
  fi
  STAGED="$(mktemp "$DESTINATION/.radius-install.XXXXXX")"
  trap 'rm -f "$STAGED"' EXIT
  /bin/cp "$PAYLOAD" "$STAGED"
  /bin/chmod 644 "$STAGED"
  /bin/mv -f "$STAGED" "$TARGET"
  trap - EXIT
fi

printf '\n插件文件已准备好：\n%s\n\n' "$TARGET"
printf '%s\n' \
  '下一步只需在 PowerPoint 中操作一次：' \
  '1. 工具 → PowerPoint 加载项 → 添加，选择上面的 .ppam，并保持勾选。' \
  '2. 按 Office 的提示允许此插件的宏。' \
  '3. 顶部应出现「R角调整 · Native」选项卡。' \
  '4. Cmd+Q 退出再打开，确认选项卡仍在；本机 PowerPoint 16.113.3 已验证自动加载。' \
  '' \
  '此脚本仅准备安装文件；PowerPoint 注册仍需完成第 1 步。' \
  '以后使用只打开 PowerPoint 即可，无需再运行此脚本。' \
  '更新同一路径前先保存并Cmd+Q退出PowerPoint；换路径时移除旧项再添加新路径。'

if [ "$REVEAL" -eq 1 ]; then
  /usr/bin/open -R "$TARGET"
fi
