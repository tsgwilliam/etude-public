#!/usr/bin/env bash
# Run the Etude OneClick app locally (macOS/Linux).
# Usage from repo root:  bash apps/oneclick/run.sh
# Or from this folder:   bash run.sh

set -euo pipefail
cd "$(dirname "$0")"

PYTHON=""
for candidate in python3 python; do
  if command -v "$candidate" >/dev/null 2>&1; then
    PYTHON="$candidate"
    break
  fi
done

if [ -z "$PYTHON" ]; then
  echo "Python 3 not found."
  echo "Install Python 3, then retry. Options:"
  echo "  - https://www.python.org/downloads/"
  echo "  - brew install python"
  exit 1
fi

echo "Using: $($PYTHON --version)"

if [ ! -d .venv ]; then
  echo "Creating virtual environment in apps/oneclick/.venv ..."
  "$PYTHON" -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo ""
echo "Starting Streamlit at http://localhost:8501"
echo "Press Ctrl+C to stop."
echo ""

python -m streamlit run app.py
