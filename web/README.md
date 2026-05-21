# Turnitin Assistant — Vercel Frontend

This folder contains a standalone static frontend that connects to your VPS backend via API.

## Deploy to Vercel

1. Go to [vercel.com](https://vercel.com) and click **Add New → Project**.
2. Import your Git repository (or drag & drop this `web/` folder).
3. Set the **Root Directory** to `web` (if using a monorepo) or deploy the `web/` folder directly.
4. No build step required — it's plain HTML/CSS/JS.

## Usage

1. Open the deployed Vercel URL.
2. Enter your VPS API base URL (e.g., `https://your-vps-domain.com:8000`).
3. Enter the API token if you set one on the server.
4. Click **Connect & Check Health**.
5. Use the buttons to start workflows.

## Important Notes

- **Vercel does NOT run Python or Playwright.**  
  This frontend is purely a static UI. All browser automation runs on your VPS backend.
- **Your VPS must be online and reachable** for the frontend to work.
- For production, configure your VPS with:
  - SSL/HTTPS (e.g., nginx + Let's Encrypt)
  - A firewall allowing only necessary ports
  - systemd to keep the FastAPI server running