#!/bin/bash
#
# guard.sh — 守护 R 角调整的 manifest（LaunchAgent 触发，每 30s 跑一次）
#
# 问题：PowerPoint Cmd+Q 退出时，会回收 sandbox 里的 wef/ 目录，
#       导致下次开 PPT 时 manifest 找不到，弹"加载项错误"。
#
# 修法：每 30s 检查一次，PowerPoint 在跑 + wef 缺 manifest → 自动从稳定位置恢复
#
# 源 manifest 在哪：$HOME/Library/Application Support/RadiusInPpt/manifest.xml
# 目标 manifest 在哪：$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml
#
# 这个脚本不抛错、不弹窗（daemon 行为），只 log 到 /tmp/radius_in_ppt_guard.log
#
set -u

LOG="/tmp/radius_in_ppt_guard.log"
TS() { date '+%Y-%m-%d %H:%M:%S'; }
log() { echo "[$(TS)] $*" >> "$LOG"; }

# 稳定位置的源 manifest（启动器装 .app 时拷过去）
SOURCE_MANIFEST="$HOME/Library/Application Support/RadiusInPpt/manifest.xml"

# PowerPoint 的 wef 路径
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
WEF_MANIFEST="$WEF_DIR/manifest.xml"

# 检查 PowerPoint 是否在跑
PP_RUNNING=0
if pgrep -f "Microsoft PowerPoint" >/dev/null 2>&1; then
  PP_RUNNING=1
fi

# PowerPoint 没在跑 → 不需要恢复（避免 PPT 没开就触发无谓的写入）
if [ "$PP_RUNNING" = "0" ]; then
  exit 0
fi

# PowerPoint 在跑 + wef 缺 manifest → 恢复
if [ ! -f "$WEF_MANIFEST" ]; then
  if [ -f "$SOURCE_MANIFEST" ]; then
    mkdir -p "$WEF_DIR"
    if cp -f "$SOURCE_MANIFEST" "$WEF_MANIFEST" 2>/dev/null; then
      log "manifest 已恢复 → $WEF_MANIFEST"
    else
      log "manifest 恢复失败：cp 失败（权限？）"
    fi
  else
    log "源 manifest 缺失：$SOURCE_MANIFEST（需要先双击 .app 一次）"
  fi
fi

# 顺便检查 serve.js 是否在跑，没在跑就 log 一下
# （用户双击 .app 才会启动 serve.js，guard 不能直接起——需要用户主动）
if ! lsof -ti tcp:3000 -sTCP:LISTEN >/dev/null 2>&1; then
  log "serve.js 没在跑（http://localhost:3000 无响应）"
fi

exit 0
