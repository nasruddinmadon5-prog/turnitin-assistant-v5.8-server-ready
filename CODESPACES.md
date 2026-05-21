# GitHub Codespace Setup Guide

## Quick Start

1. **Open in Codespace**  
   Go to your GitHub repo → Click **Code** → **Open with Codespaces** → **New codespace**.

2. **Wait for automatic setup**  
   The `.devcontainer/devcontainer.json` runs `bash scripts/setup.sh` automatically.  
   This installs Python dependencies, Playwright Chromium, and creates required folders.

3. **Start the server**  
   ```bash
   bash scripts/run_server.sh
   ```

4. **Forward port 8000**  
   Codespace will detect the forwarded port and show a notification.  
   Click **Open in Browser** or find the URL in the **Ports** tab (Cmd+Shift+P → "Ports: Focus on Ports View").

5. **Test the health endpoint**  
   ```bash
   curl http://localhost:8000/health
   ```
   Expected response: `{"status":"ok","version":"5.8-server-ready"}`

## Limitations

- **Codespace has limited GUI/browser persistence.**  
  Playwright automation that requires a headed browser (login sessions, file downloads) may not work reliably in a Codespace environment. The browser profile does not persist between Codespace rebuilds.

- **Use Cases that work well in Codespace:**
  - Testing the FastAPI server endpoints
  - Testing `/api/status`, `/api/logs`
  - Testing the Vercel frontend connection
  - Code development and debugging

- **Use Cases better suited for a VPS or local machine:**
  - Full workflow runs with browser login sessions
  - Long-running automation tasks
  - Template screening with document processing

## Manual Setup (if devcontainer does not run automatically)

```bash
bash scripts/setup.sh
bash scripts/run_server.sh