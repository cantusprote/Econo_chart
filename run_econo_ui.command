#!/bin/zsh

set -e

SCRIPT_DIR="${0:A:h}"
PYTHON_BIN="$SCRIPT_DIR/.venv/bin/python"

cd "$SCRIPT_DIR"

if [[ ! -x "$PYTHON_BIN" ]]; then
    print -u2 ".venv가 없습니다. 먼저 setup_venv.command를 실행합니다."
    "$SCRIPT_DIR/setup_venv.command"
fi

exec "$PYTHON_BIN" "$SCRIPT_DIR/econo_ui.py"
