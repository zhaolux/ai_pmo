#!/bin/sh
set -eu
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
PYTHON_BIN=${AI_PMO_PYTHON:-/Users/zhaolu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3}
PYTHONPATH="$SCRIPT_DIR/src" exec "$PYTHON_BIN" -m ai_pmo "$@"

