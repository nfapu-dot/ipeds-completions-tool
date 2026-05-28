#!/usr/bin/env bash
# Launches the IPEDS Completions web app.
#
# Double-click this file in Finder to start the tool.
# A Terminal window will open, the app will start, and your browser
# will open to http://localhost:8501.
#
# To stop the app: close this Terminal window, or press Ctrl+C in it.

set -e

# cd into the directory this script lives in, so relative paths work
# regardless of where Finder launched it from.
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

cat <<'BANNER'
────────────────────────────────────────────────────────────
  APU · IPEDS Completions Analysis Tool
────────────────────────────────────────────────────────────
  Starting the web app…
  Your browser will open to http://localhost:8501
  To stop: close this window or press Ctrl+C
────────────────────────────────────────────────────────────
BANNER

# Use `python3 -m streamlit` so it works regardless of whether the
# `streamlit` binary is on PATH.
exec python3 -m streamlit run src/app.py
