#!/bin/bash
# Double-click in Finder, or use --no-pause for terminal/CI execution.
set -euo pipefail

PAUSE=1
finish_tests() {
  local result=$?
  trap - EXIT
  if [ "$result" -eq 0 ]; then
    printf '\n项目检查全部通过。\n'
  else
    printf '\n项目检查未全部通过（退出码 %s）。请查看上面的错误。\n' "$result" >&2
  fi
  if [ "$PAUSE" -eq 1 ] && [ -t 0 ]; then
    printf '\n按回车结束检查（Enter）：'
    read -r _tests_reply || true
  fi
  exit "$result"
}
trap finish_tests EXIT

if [ "$#" -gt 0 ]; then
  if [ "$#" -eq 1 ] && [ "$1" = '--no-pause' ]; then
    PAUSE=0
  else
    printf 'Use Run-Tests.command or Run-Tests.command --no-pause\n' >&2
    exit 2
  fi
fi

PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$PROJECT_DIR"
# Finder-launched terminals may not include Homebrew's usual locations.
PATH="$PATH:/opt/homebrew/bin:/usr/local/bin"
if [ -x "$PROJECT_DIR/.venv-native/bin/python3" ]; then
  PATH="$PROJECT_DIR/.venv-native/bin:$PATH"
fi
export PATH

printf '\nR角调整 · 项目快速检查\n目录：%s\n\n' "$PROJECT_DIR"
printf '%s\n' \
  '检查原生PPAM格式/安装模拟，以及Office.js历史逻辑回归。' \
  '这些检查不执行PowerPoint VBA；真实加载和原生业务需另做宿主验收。' \
  ''

missing=0
for runtime in python3 node npm; do
  if ! command -v "$runtime" >/dev/null 2>&1; then
    printf '缺少运行工具：%s\n' "$runtime" >&2
    missing=1
  fi
done
if [ "$missing" -ne 0 ]; then
  printf '%s\n' \
    '请先安装Python 3及Node.js/npm，再重试；本入口不会自动安装。' \
    'Python测试依赖准备方式见test/README.md。' >&2
  exit 1
fi

# Keep the package scripts as the single source of the suite list. The native
# launcher probes olefile/oletools and prefers the project's local venv.
npm run test:all
