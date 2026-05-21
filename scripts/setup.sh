#!/usr/bin/env bash
# ============================
# Turnitin Assistant v5.8-server-ready
# One-time setup script
# ============================
set -e

echo "========================================"
echo " Turnitin Assistant v5.8-server-ready"
echo " Setup Script"
echo "========================================"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

echo "[SETUP] Project root: $PROJECT_DIR"

# 1. Create virtual environment
if [ ! -d ".venv" ]; then
    echo "[SETUP] Creating Python virtual environment..."
    python3 -m venv .venv
else
    echo "[SETUP] Virtual environment already exists, skipping."
fi

# 2. Activate and install dependencies
echo "[SETUP] Installing Python dependencies..."
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

# 3. Install Playwright Chromium
echo "[SETUP] Installing Playwright Chromium browser..."
python3 -m playwright install chromium

# 4. Create required directories
echo "[SETUP] Creating required directories..."
mkdir -p data
mkdir -p storage/ojs-downloads
mkdir -p storage/turnitin-reports
mkdir -p storage/template_screening_reports
mkdir -p storage/screenshots
mkdir -p profiles/ojs-profile
mkdir -p profiles/turnitin-profile

# 5. Copy .env.example to .env if not exists
if [ ! -f ".env" ]; then
    echo "[SETUP] Creating .env from .env.example..."
    cp .env.example .env
    echo "[SETUP] ⚠️  Edit .env to set your API_TOKEN and other settings."
else
    echo "[SETUP] .env already exists, skipping."
fi

echo ""
echo "========================================"
echo " ✅ Setup complete!"
echo "========================================"
echo ""
echo "Next steps:"
echo ""
echo "  # Desktop mode (macOS):"
echo "  bash scripts/run_desktop.sh"
echo ""
echo "  # Server mode (VPS/Codespace):"
echo "  bash scripts/run_server.sh"
echo ""
echo "  # Test server health:"
echo "  curl http://localhost:8000/health"
echo ""
echo "  # Set API_TOKEN in .env for secure access."
echo "========================================"