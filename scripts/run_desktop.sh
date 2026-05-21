#!/usr/bin/env bash
# ============================
# Turnitin Assistant v5.8-server-ready
# Run Desktop Mode (pywebview)
# ============================
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

echo "[RUN] Starting Turnitin Assistant in Desktop mode..."
echo "[RUN] Using $(python3 --version)"

source .venv/bin/activate
python app.py