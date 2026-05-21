#!/usr/bin/env bash
# ============================
# Turnitin Assistant v5.8-server-ready
# Run Server Mode (FastAPI + uvicorn)
# ============================
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

echo "[RUN] Starting Turnitin Assistant in Server mode..."
echo "[RUN] Using $(python3 --version)"

source .venv/bin/activate
uvicorn server.main:app --host 0.0.0.0 --port ${PORT:-8000}