# VPS Deployment Guide

## Prerequisites

- Ubuntu 22.04+ or Debian 12+
- Python 3.10+
- Git
- Root or sudo access

## Step-by-Step Deployment

### 1. Update system and install dependencies

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y python3 python3-venv python3-pip git curl wget xvfb
```

> **`xvfb`** is required if you need to run Playwright in headed mode (visible browser).  
> For headless mode (most VPS setups), you can skip `xvfb`.

### 2. Clone the repository

```bash
git clone <your-repo-url>
cd turnitin-assistant-v5.8-server-ready
```

### 3. Run setup

```bash
bash scripts/setup.sh
```

This will:
- Create a Python virtual environment (`.venv`)
- Install all Python dependencies
- Install Playwright Chromium
- Create required data directories
- Copy `.env.example` to `.env` (edit this file!)

### 4. Configure environment

```bash
nano .env
```

Set at minimum:
```env
APP_MODE=server
HOST=0.0.0.0
PORT=8000
HEADLESS=true
API_TOKEN=your-secure-random-token
```

### 5. Test the server

```bash
bash scripts/run_server.sh
```

In another terminal:
```bash
curl http://localhost:8000/health
```

### 6. Run as a systemd service (production)

Create a service file:

```bash
sudo nano /etc/systemd/system/turnitin-assistant.service
```

```ini
[Unit]
Description=Turnitin Assistant API Server
After=network.target

[Service]
Type=simple
User=ubuntu
WorkingDirectory=/home/ubuntu/turnitin-assistant-v5.8-server-ready
ExecStart=/home/ubuntu/turnitin-assistant-v5.8-server-ready/.venv/bin/uvicorn server.main:app --host 0.0.0.0 --port 8000
Restart=always
RestartSec=5
EnvironmentFile=/home/ubuntu/turnitin-assistant-v5.8-server-ready/.env

[Install]
WantedBy=multi-user.target
```

Enable and start:

```bash
sudo systemctl daemon-reload
sudo systemctl enable turnitin-assistant
sudo systemctl start turnitin-assistant
sudo systemctl status turnitin-assistant
```

### 7. Optional: nginx reverse proxy with SSL

```bash
sudo apt install -y nginx certbot python3-certbot-nginx
```

Create nginx config:

```bash
sudo nano /etc/nginx/sites-available/turnitin-assistant
```

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Enable and get SSL:

```bash
sudo ln -s /etc/nginx/sites-available/turnitin-assistant /etc/nginx/sites-enabled/
sudo certbot --nginx -d your-domain.com
sudo systemctl restart nginx
```

### 8. Firewall configuration

```bash
sudo ufw allow 22/tcp      # SSH
sudo ufw allow 80/tcp      # HTTP
sudo ufw allow 443/tcp     # HTTPS
sudo ufw enable
```

## Profile Persistence

- Browser profiles are stored at `~/.turnitin-assistant/profiles/`
- These persist across server restarts
- Login sessions remain valid as long as the browser profile is intact
- Profiles on the VPS are **separate** from your local macOS profiles

## How Testers Connect via Vercel

1. Deploy the `web/` folder to Vercel
2. Set the API URL to `https://your-vps-domain.com` (or `http://your-vps-ip:8000`)
3. Enter the API token configured in `.env`
4. All automation runs on the VPS — Vercel only shows the UI

## Monitoring Logs

```bash
journalctl -u turnitin-assistant -f   # systemd logs
curl http://localhost:8000/api/logs   # API logs endpoint