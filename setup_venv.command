#!/bin/zsh

set -e

SCRIPT_DIR="${0:A:h}"
VENV_DIR="$SCRIPT_DIR/.venv"
BUNDLED_PYTHON="/Users/sim/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

cd "$SCRIPT_DIR"

if [[ -x "$BUNDLED_PYTHON" ]]; then
    BASE_PYTHON="$BUNDLED_PYTHON"
elif command -v python3 >/dev/null 2>&1; then
    BASE_PYTHON="$(command -v python3)"
else
    print -u2 "Python 3를 찾지 못했습니다."
    exit 1
fi

if ! "$BASE_PYTHON" -c 'import tkinter' >/dev/null 2>&1; then
    print -u2 "선택된 Python에 tkinter가 없습니다. Tk 지원 Python이 필요합니다: $BASE_PYTHON"
    exit 1
fi

if [[ -x "$VENV_DIR/bin/python" ]] && ! "$VENV_DIR/bin/python" -c 'import tkinter' >/dev/null 2>&1; then
    BROKEN_VENV="$SCRIPT_DIR/.venv.broken.$(date +%Y%m%d-%H%M%S)"
    print -u2 "기존 .venv에 tkinter가 없어 백업 후 다시 만듭니다: $BROKEN_VENV"
    mv "$VENV_DIR" "$BROKEN_VENV"
fi

if [[ ! -x "$VENV_DIR/bin/python" ]]; then
    "$BASE_PYTHON" -m venv "$VENV_DIR"
fi

"$VENV_DIR/bin/python" -m pip install --upgrade pip
"$VENV_DIR/bin/python" -m pip install -r "$SCRIPT_DIR/requirements.txt"
"$VENV_DIR/bin/python" -m pip check

print "환경 구성이 완료되었습니다: $VENV_DIR"
