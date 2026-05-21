# Turnitin Assistant v5.8-server-ready

Workflow automation system developed for Unit Jurnal, Penerbit UTM Press.

Automates OJS manuscript download → Turnitin similarity checking → Template screening → Report upload back to OJS.

## Modes of Operation

### 1. Local Desktop Mode (macOS — original)

```bash
bash scripts/setup.sh
bash scripts/run_desktop.sh
```

Launches the native pywebview desktop window with the full dashboard UI.

### 2. GitHub Codespace / Dev Mode

```bash
bash scripts/setup.sh
bash scripts/run_server.sh
```

Runs the FastAPI server on port 8000. Use forwarded port to test `/health` and `/api/status`.

> **Note:** Browser automation (Playwright) requires a headed browser environment. Codespace has limited GUI persistence. For full automation with login sessions, use a VPS or local machine.

### 3. VPS / Server Mode

```bash
git clone <repo-url>
cd turnitin-assistant-v5.8-server-ready
bash scripts/setup.sh
bash scripts/run_server.sh
```

The FastAPI server exposes REST endpoints. Optionally run behind nginx reverse proxy with systemd for production.

See [DEPLOY_VPS.md](DEPLOY_VPS.md) for full instructions.

### 4. Vercel Frontend (External Testers)

The `web/` folder contains a standalone static frontend that can be deployed to Vercel.

- Deploy the `web/` folder to Vercel.
- Set the API base URL to your VPS server address.
- Testers interact with the web UI; **all Playwright automation runs on your VPS/backend**, not on Vercel.
- Vercel only serves static frontend files — it does not run Python or Playwright.

See [web/README.md](web/README.md) for deployment details.

## Project Structure

```
├── app.py                          # Desktop app entry point (pywebview)
├── config.py                       # Central configuration
├── requirements.txt                # Python dependencies
├── .env.example                    # Environment template
├── .gitignore
├── README.md
├── CODESPACES.md                   # Codespace setup guide
├── DEPLOY_VPS.md                   # VPS deployment guide
├── scripts/
│   ├── setup.sh                    # One-time setup script
│   ├── run_desktop.sh              # Launch desktop mode
│   └── run_server.sh               # Launch server mode
├── server/
│   └── main.py                     # FastAPI backend
├── web/                            # Vercel-deployable frontend
│   ├── README.md
│   ├── package.json
│   ├── index.html
│   └── src/
│       └── main.js
├── engines/                        # Workflow automation engines
│   ├── ojs_download_engine.py
│   ├── turnitin_engine_api.py
│   ├── ojs_report_upload_engine.py
│   └── template_screening_engine.py
├── data/                           # Excel bridge data
├── storage/                        # Downloads, reports, screenshots
├── profiles/                       # Browser profiles (gitignored)
├── ui/                             # Desktop UI files
├── .devcontainer/
│   └── devcontainer.json           # Codespace dev container config
└── changes/                        # Changelogs
```

## Requirements

- Python 3.10+
- Playwright (Chromium)
- macOS for desktop mode; Linux for VPS mode

## Quick Start

```bash
# 1. Clone and enter directory
cd turnitin-assistant-v5.8-server-ready

# 2. Run setup
bash scripts/setup.sh

# 3. Choose mode:
# Desktop:
bash scripts/run_desktop.sh

# Or server:
bash scripts/run_server.sh