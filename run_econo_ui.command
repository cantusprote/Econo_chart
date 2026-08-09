#!/bin/zsh

set -e

PROJECT_DIR="/Users/sim/Dev/Econo_chart"
PYTHON_BIN="$PROJECT_DIR/.venv/bin/python"

cd "$PROJECT_DIR"

if [[ ! -x "$PYTHON_BIN" ]]; then
    print -u2 ".venv가 없습니다. 먼저 setup_venv.command를 실행합니다."
    "$PROJECT_DIR/setup_venv.command"
fi

exec "$PYTHON_BIN" "$PROJECT_DIR/econo_ui.py" "$@"
