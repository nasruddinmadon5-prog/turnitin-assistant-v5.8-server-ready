# Academic Auto Bridge v5.4.1 Changes

## UI Changes

### Storage Paths → Recent Paper Progress
- **Removed** the "Storage Paths" section from the right sidebar (previously showed OJS Downloads, Turnitin Reports, and Screenshots paths)
- **Added** "Recent Paper Progress" section with 4 paper progress cards aligned with actual data from `data/master update.xlsx`:
  - **SUB-26801** (Mohd Ibrahim et al.): Downloaded - Pending Turnitin
  - **SUB-26800** (Karim et al.): Turnitin Download Failed (Score: 25%)
  - **SUB-26799** (Surbakti et al.): Downloaded - Pending Turnitin
  - **SUB-26791** (Kadum et al.): Turnitin Completed (Score: 34%) - Pending Upload
- Each card shows real submission statuses (Download ✅/TURNITIN ✅/❌/⬜/UPLOAD ⬜) matching Excel sheets (download, turnitin, upload)

### Version Updates
- Updated all version references from v5.3/v5.3.1 to v5.4.1:
  - `app.py` - module docstring, startup banner, window title
  - `config.py` - module docstring, app name, version
  - `ui/index.html` - header title, bridge label, initial log message

## Dark Mode / Light Mode Toggle
- **Added** a theme toggle button (sun/moon icon) in the top navigation bar
- **Implemented** CSS custom properties (`--primary`, `--secondary`, `--surface`, etc.) that dynamically switch between dark and light color palettes
- **Light mode** colors: white surfaces, dark text, blue primary (#005ac2), green secondary (#008a5e)
- **Dark mode** colors: dark surfaces, light text, blue primary (#adc6ff), green secondary (#4edea3)
- Theme preference is saved in `localStorage` and persists across sessions
- Respects system `prefers-color-scheme` on first visit

## No Code Changes
- All functions, JavaScript code, and backend logic remain unchanged
- Only UI layout, styling (CSS variables), and version strings were modified
