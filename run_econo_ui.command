#!/bin/zsh

set -e

PROJECT_DIR="/Users/sim/Dev/Econo_chart"
PYTHON_BIN="$PROJECT_DIR/.venv/bin/python"

cd "$PROJECT_DIR"

if [[ ! -x "$PYTHON_BIN" ]] || ! "$PYTHON_BIN" -c 'import tkinter' >/dev/null 2>&1; then
    print -u2 ".venv가 없거나 Tk를 사용할 수 없습니다. 환경을 구성합니다."
    "$PROJECT_DIR/setup_venv.command"
fi

exec "$PYTHON_BIN" "$PROJECT_DIR/econo_ui.py" "$@"
