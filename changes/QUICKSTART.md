# Quick Start Guide - Academic Auto Bridge v5.1

## Installation

1. Ensure you have Python 3.8+ installed
2. Install required packages:
```bash
cd /Users/nasruddinmadon/turnitin-assistant-v5.1
pip3 install pywebview playwright openpyxl
playwright install chromium
```

## Running the Application

```bash
python3 app.py
```

The dashboard will open automatically.

## Workflow Guide

### OJS Download Workflow

1. **Click "Open OJS Login"**
   - Browser window opens to OJS submissions page
   - Dashboard shows: "Browser opened successfully"

2. **Login manually in the browser**
   - Use your OJS credentials
   - Navigate to the submissions page if needed
   - Workflow automatically detects login (no ENTER needed)

3. **Click "Start OJS Download"**
   - Dashboard terminal shows real-time progress
   - Downloads latest unprocessed submission
   - Logs to `data/master update.xlsx` (download sheet)
   - Files saved to `storage/ojs-downloads/{submission_no}/`

### Turnitin Upload Workflow

1. **Click "Open Turnitin Login"**
   - Browser window opens to Turnitin login page
   - Dashboard shows: "Browser opened successfully"

2. **Login manually in the browser**
   - Use your Turnitin credentials
   - Wait for dashboard confirmation

3. **Click "Start Turnitin Upload"**
   - Dashboard terminal shows real-time progress
   - Reads from Excel "download" sheet automatically
   - Uploads to Turnitin
   - Filters UTM sources
   - Downloads similarity report
   - Logs to `data/master update.xlsx` (turnitin sheet)
   - Reports saved to `storage/turnitin-reports/`

## Key Features

✅ **No Terminal Prompts** - Everything controlled from UI buttons
✅ **Single Dashboard** - No separate windows
✅ **Real-Time Logging** - See progress in dashboard terminal
✅ **Auto Login Detection** - No need to press ENTER
✅ **Excel Bridge** - Automatic data flow from OJS to Turnitin

## File Structure

```
turnitin-assistant-v5.1/
├── app.py                    # Main application
├── config.py                 # Configuration
├── ui/
│   └── index.html           # Dashboard UI
├── engines/
│   ├── __init__.py          # Package init
│   ├── ojs_download_engine.py      # OJS automation
│   ├── turnitin_upload_engine.py   # Turnitin automation (original)
│   └── turnitin_engine_api.py      # Turnitin API wrapper
├── data/
│   └── master update.xlsx   # Excel logging
├── storage/
│   ├── ojs-downloads/       # Downloaded articles
│   ├── turnitin-reports/    # Similarity reports
│   └── screenshots/         # Debug screenshots
└── profiles/
    ├── ojs-profile/         # OJS browser profile
    └── turnitin-profile/    # Turnitin browser profile
```

## Troubleshooting

**Browser doesn't open:**
- Check Playwright installation: `playwright install chromium`
- Check permissions for browser execution

**Dashboard doesn't show logs:**
- Refresh the dashboard window
- Check browser console for JavaScript errors

**Workflow stops unexpectedly:**
- Check dashboard terminal for error messages
- Check `storage/screenshots/` for debug images
- Review Excel file for logged errors

## Support

For issues or questions, refer to:
- `REFACTORING_NOTES.md` - Technical details
- `README_v5.1.md` - Project overview
- Dashboard terminal - Real-time logs
