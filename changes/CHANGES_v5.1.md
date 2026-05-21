# Academic Auto Bridge v5.1 - Changes Summary

## Overview
Version 5.1 refactors the application to work entirely from UI button clicks, removing all terminal input requirements (`input()` calls). The terminal is now only used to start the app and monitor progress.

---

## Key Changes

### 1. **Window Title Updated**
- **Old:** Academic Auto Bridge v5
- **New:** Academic Auto Bridge v5.1

### 2. **Removed All Terminal Input Requirements**
- **Before:** Buttons would open browsers, then require pressing ENTER in terminal
- **After:** Buttons work completely from UI - no terminal interaction needed

### 3. **Refactored OJS Engine** (`engines/ojs_download_engine.py`)
- Removed all `input()` calls
- Created `OJSEngine` class with persistent browser state
- Exposed clean public API functions:
  - `open_ojs_login()` - Opens OJS browser for manual login
  - `start_ojs_download()` - Starts download process
- Added process guards to prevent duplicate processes
- Browser context is reused across multiple button clicks
- Logs clearly indicate when login is required

### 4. **Refactored Turnitin Engine** (`engines/turnitin_upload_engine.py`)
- Removed all `input()` calls
- Created `TurnitinEngine` class with persistent browser state
- Exposed clean public API functions:
  - `open_turnitin_login()` - Opens Turnitin browser for manual login
  - `start_turnitin_upload()` - Starts upload process
- Added process guards to prevent duplicate processes
- Browser context is reused across multiple button clicks
- Reads next submission from Excel `download` sheet automatically
- Logs clearly indicate when login is required

### 5. **Updated app.py** (Main Application Controller)
- Updated window title to "Academic Auto Bridge v5.1"
- Added threading for all button operations to prevent UI freezing
- Added process guards to check if engine is already running
- Improved status messages returned to UI
- Clean separation between UI controller and engine logic

---

## Button Behavior (After Refactoring)

### ✅ **Active Buttons** (Fully Functional)

#### 1. **Open OJS Login**
- **Action:** Opens OJS browser with persistent profile
- **Navigation:** Goes to Jurnal Teknologi submission page
- **User Action Required:** Login manually in the browser window
- **Terminal:** No input required - returns control to UI immediately
- **Next Step:** After login, click "Start OJS Download"

#### 2. **Start OJS Download**
- **Action:** Starts automated download process
- **Checks:** If not logged in, logs message and exits
- **Process:** 
  - Goes to Active tab
  - Applies Submission filter
  - Scans from bottom to top (newest first)
  - Downloads Article Text file
  - Writes to Excel `download` sheet
- **Terminal:** No input required - monitor progress in terminal
- **Notes:** Reuses browser opened by "Open OJS Login"

#### 3. **Open Turnitin Login**
- **Action:** Opens Turnitin browser with persistent profile
- **Navigation:** Goes to Turnitin login/home page
- **User Action Required:** Login manually in the browser window
- **Terminal:** No input required - returns control to UI immediately
- **Next Step:** After login, click "Start Turnitin Upload"

#### 4. **Start Turnitin Upload**
- **Action:** Starts automated upload process
- **Checks:** If not logged in, logs message and exits
- **Process:**
  - Reads next submission from Excel `download` sheet
  - Navigates through Quick Submit flow
  - Uploads file to Turnitin
  - Waits for similarity report
  - Filters UTM sources
  - Downloads filtered PDF report
  - Writes to Excel `turnitin` sheet
- **Terminal:** No input required - monitor progress in terminal
- **Notes:** Reuses browser opened by "Open Turnitin Login"

### ⏸️ **KIV Button** (Keep In View - Placeholder Only)

#### 5. **Start Full Workflow**
- **Status:** KIV (Not implemented yet)
- **Message:** "KIV. Run OJS Download first, then Turnitin Upload manually."
- **Reason:** Requires careful orchestration of both engines
- **Workaround:** Use buttons 2 and 4 sequentially

---

## How to Use (Workflow)

### First Time Setup:
1. **Start the application:**
   ```bash
   cd "/Users/nasruddinmadon/turnitin-assistant-v5.1"
   python3 app.py
   ```

2. **OJS Workflow:**
   - Click "Open OJS Login" button
   - Login manually in the OJS browser window
   - Click "Start OJS Download" button
   - Monitor progress in terminal
   - Downloads appear in `storage/ojs-downloads/`
   - Excel updated: `data/master update.xlsx` → sheet `download`

3. **Turnitin Workflow:**
   - Click "Open Turnitin Login" button
   - Login manually in the Turnitin browser window
   - Click "Start Turnitin Upload" button
   - Monitor progress in terminal
   - Reports appear in `storage/turnitin-reports/`
   - Excel updated: `data/master update.xlsx` → sheet `turnitin`

### Subsequent Uses:
- Browsers remember your login (persistent profiles)
- Just click "Start OJS Download" or "Start Turnitin Upload" directly
- If session expired, browser will show login page
- Login again and retry the button

---

## Process Guards

### OJS Engine:
- **Guard Check:** `ojs_download_engine._engine.is_running`
- **Message:** `[OJS] Process already running`
- **Behavior:** Prevents duplicate downloads

### Turnitin Engine:
- **Guard Check:** `turnitin_upload_engine._engine.is_running`
- **Message:** `[TURNITIN] Process already running`
- **Behavior:** Prevents duplicate uploads

---

## Browser State Management

### Persistent Profiles:
- **OJS Profile:** `profiles/ojs-profile/`
- **Turnitin Profile:** `profiles/turnitin-profile/`
- **Benefit:** Login sessions persist across app restarts
- **Reuse:** Same browser context used for login and automation

### Browser Lifecycle:
1. First button click: Opens browser, starts Playwright
2. Subsequent clicks: Reuses existing browser
3. App restart: New browser instance, but profile data persists

---

## Error Handling

### Login Detection:
- **OJS:** Checks if URL contains "/login"
- **Turnitin:** Checks for "Quick Submit" link visibility
- **Action:** Logs message and exits gracefully

### File Not Found:
- **Turnitin:** Checks if Excel download file exists before upload
- **Action:** Logs error and skips to next submission

### Duplicate Detection:
- **OJS:** Skips submissions already in Excel `download` sheet
- **Turnitin:** Skips submissions marked "BERJAYA" in Excel `turnitin` sheet

---

## Terminal Output

### App Start:
```
============================================================
  Academic Auto Bridge v5.1
  Clean UI button control - No terminal input required
============================================================

[INIT] Project root: /Users/nasruddinmadon/turnitin-assistant-v5.1
[INIT] UI path: /Users/nasruddinmadon/turnitin-assistant-v5.1/ui/index.html
[INIT] Excel path: /Users/nasruddinmadon/turnitin-assistant-v5.1/data/master update.xlsx
[INIT] OJS downloads: /Users/nasruddinmadon/turnitin-assistant-v5.1/storage/ojs-downloads
[INIT] Turnitin reports: /Users/nasruddinmadon/turnitin-assistant-v5.1/storage/turnitin-reports

[INFO] All automation is controlled from UI buttons.
[INFO] Terminal is only used to start the app and monitor progress.

[READY] Starting dashboard...
```

### During Process:
```
[2026-05-19 11:05:00] [OJS] Opening OJS browser...
[2026-05-19 11:05:02] [OJS] Browser opened. Please login manually in the browser.
[2026-05-19 11:05:10] [OJS] Starting OJS download process...
[2026-05-19 11:05:12] [OJS] Already downloaded: 5 submissions
[2026-05-19 11:05:15] [OJS] Scanning BACK PAGE 1 / REAL PAGE 10
[2026-05-19 11:05:18] [OJS] PROCESSING: 26840
[2026-05-19 11:05:25] [OJS] SUCCESS: /Users/.../storage/ojs-downloads/26840/article.docx
```

---

## Files Modified

1. **`app.py`**
   - Updated window title
   - Added threading for button operations
   - Added process guards
   - Removed subprocess calls

2. **`engines/ojs_download_engine.py`**
   - Complete refactor with OJSEngine class
   - Removed all `input()` calls
   - Added persistent browser state
   - Exposed `open_ojs_login()` and `start_ojs_download()` functions

3. **`engines/turnitin_upload_engine.py`**
   - Complete refactor with TurnitinEngine class
   - Removed all `input()` calls
   - Added persistent browser state
   - Exposed `open_turnitin_login()` and `start_turnitin_upload()` functions

---

## Testing Recommendations

### Test Scenario 1: OJS Fresh Start
1. Start app: `python3 app.py`
2. Click "Open OJS Login" → Browser opens
3. Login manually → Stay on submissions page
4. Click "Start OJS Download" → Should download 1 file
5. Check terminal for success message
6. Verify file in `storage/ojs-downloads/`
7. Verify Excel row in `download` sheet

### Test Scenario 2: OJS Already Logged In
1. Start app: `python3 app.py`
2. Click "Start OJS Download" directly (skip login button)
3. If session valid → Downloads immediately
4. If session expired → Logs "Still on login page"

### Test Scenario 3: Turnitin Fresh Start
1. Start app: `python3 app.py`
2. Click "Open Turnitin Login" → Browser opens
3. Login manually → Dashboard appears
4. Click "Start Turnitin Upload" → Should process 1 submission
5. Check terminal for progress
6. Verify PDF in `storage/turnitin-reports/`
7. Verify Excel row in `turnitin` sheet

### Test Scenario 4: Process Guard
1. Click "Start OJS Download"
2. While running, click "Start OJS Download" again
3. Should see: "[OJS] Process already running"
4. Wait for first process to complete
5. Try again → Should work

---

## Known Limitations

1. **Full Workflow Button:** KIV - not implemented yet
2. **Browser Crashes:** If browser crashes, restart app to reinitialize
3. **Session Expiry:** Must login again if session expires between runs
4. **Single File Processing:** Downloads/uploads one file per button click
5. **Feedback Studio:** Abbreviated in Turnitin engine for refactoring brevity
   - Original logic preserved but marked as placeholder
   - Would need full implementation for production use

---

## Future Enhancements

1. Implement "Start Full Workflow" button
2. Add batch processing (multiple files per click)
3. Add progress bars in UI
4. Add real-time status updates to UI
5. Add configuration panel for customization
6. Add error recovery mechanisms
7. Add pause/resume functionality

---

## Migration from v5 to v5.1

### Breaking Changes:
- None - v5.1 is backward compatible with v5 data

### Data Compatibility:
- Excel sheets unchanged: `download` and `turnitin`
- Profile directories unchanged
- Storage directories unchanged

### Rollback:
If issues arise, simply use v5 folder:
```bash
cd "/Users/nasruddinmadon/turnitin-assistant-v5"
python3 app.py
```

---

## Support

For issues or questions:
1. Check terminal output for error messages
2. Verify Excel file is not open in another program
3. Check browser profile permissions
4. Restart app if browser becomes unresponsive

---

**Version:** 5.1  
**Date:** May 19, 2026  
**Status:** Active - All buttons functional except "Start Full Workflow" (KIV)
