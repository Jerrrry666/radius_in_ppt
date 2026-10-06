#!/bin/bash
# Native branch: keep the familiar build entry point, deliver an embedded PPAM.
set -euo pipefail
PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
if [ "${1:-}" = '--legacy-taskpane' ]; then
  shift
  exec bash "$PROJECT_DIR/tools/build-taskpane-app.sh" "$@"
fi
exec python3 "$PROJECT_DIR/tools/build-native.py" --distribution "$@"
