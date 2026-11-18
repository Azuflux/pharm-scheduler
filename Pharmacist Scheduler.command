#!/bin/bash
# Pharmacist Scheduler — double-click to start the app in your browser.
# Leave the Terminal window open while you use it; close the window to stop the app.

PORT=8501
URL="http://localhost:$PORT"

# ---- find the project folder (next to this file, or the usual spots)
HERE="$(cd "$(dirname "$0")" && pwd)"
for d in "$HERE" "$HOME/pharm-scheduler" "$HOME/pharm-scheduler/frontend-1/pharm-scheduler" \
         "$HOME/Documents/pharm-scheduler" "$HOME/Desktop/pharm-scheduler"; do
  if [ -f "$d/app.py" ] && [ -d "$d/scheduler" ]; then APP_DIR="$d"; break; fi
done
if [ -z "$APP_DIR" ]; then
  echo "Couldn't find the pharm-scheduler folder (the one with app.py in it)."
  echo "Put it at $HOME/pharm-scheduler, or put this file inside that folder."
  read -n 1 -s -r -p "Press any key to close."; exit 1
fi
cd "$APP_DIR" || exit 1

# ---- already running? just open the browser
if curl -s -o /dev/null "$URL"; then open "$URL"; exit 0; fi

# ---- find Python 3.11 or newer
ok_py() { "$1" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)' 2>/dev/null; }
PY=""
for p in python3.13 python3.12 python3.11 python3 \
         /Library/Frameworks/Python.framework/Versions/3.13/bin/python3 \
         /opt/homebrew/bin/python3 /usr/local/bin/python3; do
  if command -v "$p" >/dev/null 2>&1 && ok_py "$p"; then PY="$(command -v "$p")"; break; fi
done

# ---- make the private Python environment once (or rebuild it if it's broken/too old)
if [ ! -x .venv/bin/python ] || ! ok_py .venv/bin/python; then
  if [ -z "$PY" ]; then
    echo "Python 3.11 or newer isn't installed. Get it from https://www.python.org/downloads/"
    read -n 1 -s -r -p "Press any key to close."; exit 1
  fi
  echo "First-time setup (about a minute)..."
  rm -rf .venv && "$PY" -m venv .venv || { read -n 1 -s -r -p "Setup failed. Press any key."; exit 1; }
fi

# ---- install/update packages only when requirements.txt has changed
if ! cmp -s requirements.txt .venv/.installed-requirements; then
  echo "Installing packages..."
  .venv/bin/python -m pip install -q --upgrade pip &&
  .venv/bin/python -m pip install -q -r requirements.txt &&
  cp requirements.txt .venv/.installed-requirements ||
  { read -n 1 -s -r -p "Package install failed (check your internet). Press any key."; exit 1; }
fi

# ---- start the app, then open the browser when it's ready
echo "Starting the Pharmacist Scheduler..."
.venv/bin/python -m streamlit run app.py --server.headless true --server.port "$PORT" \
  --browser.gatherUsageStats false &
APP_PID=$!
trap 'kill $APP_PID 2>/dev/null' EXIT INT TERM
for _ in $(seq 1 60); do
  if curl -s -o /dev/null "$URL"; then break; fi
  sleep 0.5
done
open "$URL"
echo ""
echo "Running at $URL"
echo "Leave this window open while you use it. Close it (or press Ctrl+C) to stop."
wait $APP_PID
