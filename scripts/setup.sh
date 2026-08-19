#!/bin/bash
set -e

SKILL_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VENV_DIR="$SKILL_DIR/.venv"

echo "=== DOCX Format Normalizer - Environment Setup ==="

# Find a suitable Python
PYTHON=""
for py in python3.12 python3.11 python3.13 python3; do
    if command -v "$py" &>/dev/null; then
        PYTHON="$py"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "ERROR: No suitable Python found. Need Python 3.8+."
    exit 1
fi

echo "[INFO] Using Python: $($PYTHON --version 2>&1)"

# Remove old venv if exists and broken
if [ -d "$VENV_DIR" ]; then
    if ! env -u PYTHONHOME -u PYTHONPATH "$VENV_DIR/bin/python" -c "import docx" 2>/dev/null; then
        echo "[INFO] Removing broken venv..."
        rm -rf "$VENV_DIR"
    else
        echo "[INFO] Existing venv is working."
        echo "SUCCESS: Python path = $VENV_DIR/bin/python"
        echo "NOTE: Use 'env -u PYTHONHOME -u PYTHONPATH $VENV_DIR/bin/python' to run scripts"
        exit 0
    fi
fi

# Create virtual environment
if command -v uv &>/dev/null; then
    echo "[INFO] Creating venv with uv..."
    uv venv --python 3.12 "$VENV_DIR" 2>/dev/null || uv venv "$VENV_DIR"
    PY="$VENV_DIR/bin/python"
    uv pip install --python "$PY" 'python-docx>=0.8.11' 'lxml>=4.9,<6'
else
    echo "[INFO] Creating venv with python -m venv..."
    "$PYTHON" -m venv "$VENV_DIR"
    PY="$VENV_DIR/bin/python"
    "$PY" -m pip install --quiet --upgrade pip
    "$PY" -m pip install --quiet 'python-docx>=0.8.11' 'lxml>=4.9,<6'
fi

# Verify installation (unset conflicting env vars)
env -u PYTHONHOME -u PYTHONPATH "$PY" -c "from docx import Document; from docx.shared import Pt, Mm; print('OK: python-docx ready')" 2>&1 || {
    echo "ERROR: python-docx installation failed."
    echo "Try manually: $PY -m pip install 'python-docx>=0.8.11' 'lxml>=4.9,<6'"
    exit 1
}

echo ""
echo "SUCCESS: Environment ready."
echo "  Python path: $PY"
echo "  Use this path to run docx_formatter.py"
