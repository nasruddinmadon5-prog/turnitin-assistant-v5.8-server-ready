# Changelog v5.8-server-ready

## Overview

This release adds server/API mode alongside the existing desktop app. The project can now be:
- Deployed to a VPS as a FastAPI server
- Developed in GitHub Codespaces
- Connected to a Vercel frontend for external testers
- Installed via a single bash command

## New Files Created

```
requirements.txt                  # Root-level dependencies (includes FastAPI stack)
.gitignore                        # Proper ignores for server deployment
.env.example                      # Updated with server settings
README.md                         # Comprehensive setup guide
CODESPACES.md                     # Codespace-specific instructions
DEPLOY_VPS.md                     # VPS deployment guide with systemd + nginx

scripts/
├── setup.sh                      # One-time setup: venv, deps, playwright, dirs
├── run_desktop.sh                # Launch pywebview desktop mode
└── run_server.sh                 # Launch FastAPI server

server/
└── main.py                       # FastAPI REST backend with workflow orchestration

web/
├── README.md                     # Vercel deployment instructions
├── package.json
├── index.html                    # Remote dashboard UI
└── src/main.js                   # API client JS

.devcontainer/
└── devcontainer.json             # GitHub Codespace config

changes/
└── CHANGELOG_v5.8_server_ready.md  # This file
```

## Modified Files

### config.py
- Added `python-dotenv` loading with fallback manual parsing
- Added platform-specific `APP_DATA_DIR`:
  - macOS: `~/Library/Application Support/TurnitinAssistant`
  - Linux/VPS: `~/.turnitin-assistant/`
- Added env var path overrides:
  - `MASTER_EXCEL_PATH`
  - `OJS_DOWNLOADS_DIR`
  - `TURNITIN_REPORTS_DIR`
  - `SCREENING_REPORTS_DIR`
- Updated `TEMPLATE_SCREENING_CONFIG` to use `SCREENING_REPORTS_DIR`

### app.py
- `_open_existing_path` is now cross-platform:
  - macOS: `open`
  - Windows: `os.startfile`
  - Linux: `xdg-open`
  - Graceful fallback if command unavailable

## Server Architecture

- **FastAPI** with CORS enabled for Vercel frontend
- Endpoints:
  - `GET /health` — public health check
  - `GET /api/status` — version, paths, workflow states
  - `POST /api/ojs-download/start`
  - `POST /api/turnitin-upload/start`
  - `POST /api/template-screening/start`
  - `POST /api/ojs-report-upload/start`
  - `GET /api/logs` — in-memory log buffer (last 200)
  - `GET /api/files/master-excel` — download Excel
- **Workflow locking**: prevents concurrent runs of the same workflow
- **API token auth**: optional Bearer token via `API_TOKEN` env var
- All workflows run in background daemon threads

## What Was NOT Changed

- All engine logic (selectors, automation) — untouched
- pywebview desktop mode — fully preserved
- Excel bridge (openpyxl) — compatible
- AI verifier config — preserved
- Profile migration logic — preserved
- UI files — untouched

## Removed

- Duplicate nested folder `turnitin-assistant-v5.7/` inside the project root