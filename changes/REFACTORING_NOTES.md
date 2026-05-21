# Academic Auto Bridge v5.1 - Refactoring Notes

## Date: May 19, 2026

## Summary of Changes

This refactoring converted the Academic Auto Bridge from a terminal-dependent, subprocess-based system to a fully integrated UI-controlled dashboard application.

## Key Changes

### 1. OJS Download Engine (`engines/ojs_download_engine.py`)

**Removed:**
- All `input()` prompts that blocked workflow (3 instances)
- Terminal dependency for user interaction

**Added:**
- `set_ui_log_callback()` - Set callback for UI logging
- `ui_log()` - Unified logging function (UI or console)
- `wait_manual_login_no_input()` - Automatic login detection (no ENTER prompt)
- `open_ojs_login_browser()` - Opens browser and returns control immediately
- `run_ojs_download_workflow()` - Main callable workflow function

**Changes:**
- Login waits automatically by polling URL instead of waiting for ENTER
- All logs go through callback system to dashboard UI
- Browser stays open for reuse between workflows

### 2. Turnitin Upload Engine (New API Layer)

**Created:** `engines/turnitin_engine_api.py`

This new wrapper provides clean API access to the existing Turnitin engine without modifying its complex logic.

**Functions:**
- `set_ui_log_callback()` - Configure UI logging
- `open_turnitin_login_browser()` - Opens browser for manual login
- `run_turnitin_upload_workflow()` - Executes complete upload workflow
- `close_turnitin_browser()` - Cleanup function

**Key Features:**
- No separate webview window (uses Playwright directly)
- Wraps existing `turnitin_upload_engine.py` without breaking it
- Browser context reused across workflows
- All logs sent to dashboard UI

### 3. Main Application (`app.py`)

**Removed:**
- `subprocess.run()` calls to engines
- Terminal-based workflow execution

**Added:**
- Direct imports of engine modules
- `AcademicBridgeAPI.send_log_to_ui()` - Real-time log streaming
- Thread-based execution for non-blocking UI
- Global browser state management

**Changes:**
- All engine functions called directly (no subprocess)
- Logs appear in dashboard terminal panel in real-time
- Browser windows managed by main app
- Window reference passed to API for logging

### 4. Engine Package Structure

**Created:** `engines/__init__.py` - Makes engines a proper Python package

**Files:**
- `ojs_download_engine.py` - Refactored OJS automation
- `turnitin_upload_engine.py` - Original (preserved, working)
- `turnitin_engine_api.py` - New API wrapper for Turnitin

## Architecture Changes

### Before:
```
app.py
  ├─> subprocess.run(ojs_download_engine.py)
  │     └─> Terminal: input() prompts [BLOCKING]
  │
  └─> subprocess.run(turnitin_upload_engine.py)
        └─> Creates separate webview window [DUPLICATE UI]
```

### After:
```
app.py (Main Dashboard)
  ├─> Direct import: ojs_download_engine
  │     ├─> Browser: Playwright (persistent)
  │     ├─> Logging: Callback to dashboard UI
  │     └─> Login: Auto-detect (no prompts)
  │
  └─> Direct import: turnitin_engine_api
        ├─> Wraps: turnitin_upload_engine
        ├─> Browser: Playwright (persistent)
        ├─> Logging: Callback to dashboard UI
        └─> No separate UI window
```

## User Experience Changes

### Before:
1. Run `python app.py` - Dashboard opens
2. Click "Open OJS Login" - Browser opens, terminal shows message
3. Login in browser
4. **Return to terminal, press ENTER** ← MANUAL STEP
5. Click "Start OJS Download" - Subprocess runs
6. Terminal shows logs, UI shows static message
7. Click "Start Turnitin Upload" - **Second UI window opens** ← CONFUSING
8. Two windows now open (main dashboard + Turnitin UI)

### After:
1. Run `python app.py` - Dashboard opens
2. Click "Open OJS Login" - Browser opens
3. Login in browser
4. **Workflow continues automatically** (no ENTER needed)
5. Click "Start OJS Download" - Direct function call
6. **Dashboard terminal shows real-time logs** ← LIVE FEEDBACK
7. Click "Start Turnitin Upload" - Direct function call
8. **One dashboard window throughout** ← CLEAN UX

## Technical Benefits

1. **No Terminal Dependency**
   - Workflows don't block on terminal input
   - Can be packaged as standalone app

2. **Single UI Window**
   - Consistent user experience
   - No window management confusion

3. **Real-Time Logging**
   - Engine progress visible in dashboard
   - Better debugging and monitoring

4. **Direct Function Calls**
   - Faster execution (no subprocess overhead)
   - Better error handling
   - Shared state possible

5. **Browser Reuse**
   - Login once, use multiple times
   - Faster subsequent workflows

## Preserved Features

1. **Excel Integration**
   - OJS downloads still logged to "download" sheet
   - Turnitin results still logged to "turnitin" sheet
   - Same column structure maintained

2. **Turnitin Automation Logic**
   - Complex filtering logic untouched
   - Feedback Studio workflow preserved
   - Date-based polling strategy intact

3. **Profile Management**
   - `profiles/ojs-profile` - OJS login persistence
   - `profiles/turnitin-profile` - Turnitin login persistence

4. **Storage Structure**
   - `storage/ojs-downloads/` - Article downloads
   - `storage/turnitin-reports/` - Similarity reports
   - `storage/screenshots/` - Debug screenshots

## Files Modified

- ✅ `app.py` - Complete rewrite
- ✅ `engines/ojs_download_engine.py` - Refactored for UI control
- ✅ `engines/turnitin_engine_api.py` - New API wrapper
- ✅ `engines/__init__.py` - New package init
- ⚠️ `engines/turnitin_upload_engine.py` - Unchanged (working)

## Files Removed

- ❌ `engines/turnitin_upload_engine_backup.py` - Backup file (deleted)

## Testing Recommendations

1. **OJS Workflow:**
   - Click "Open OJS Login"
   - Login manually in browser
   - Click "Start OJS Download" (no ENTER needed)
   - Verify logs appear in dashboard terminal
   - Check downloaded files in `storage/ojs-downloads/`
   - Verify Excel logging in "download" sheet

2. **Turnitin Workflow:**
   - Click "Open Turnitin Login"
   - Login manually in browser
   - Click "Start Turnitin Upload"
   - Verify no second window opens
   - Verify logs appear in dashboard terminal
   - Check reports in `storage/turnitin-reports/`
   - Verify Excel logging in "turnitin" sheet

3. **End-to-End:**
   - Run OJS workflow first
   - Then run Turnitin workflow
   - Verify Excel bridge works (Turnitin reads from OJS downloads)
   - Verify only ONE dashboard window throughout

## Known Limitations

1. **Full Workflow**: Still marked as KIV (Keep In View)
   - Sequential execution not automated
   - User must run OJS then Turnitin manually

2. **Browser Management**: Browsers stay open
   - No auto-close on completion
   - User must close manually if desired

3. **Error Recovery**: Limited retry logic
   - Workflow stops on error
   - User must restart manually

## Future Enhancements

1. Implement full workflow (OJS → Turnitin automatic)
2. Add progress bars to UI
3. Add browser close buttons to dashboard
4. Improve error recovery with retries
5. Add workflow history/status panel

## Conclusion

The refactoring successfully achieves all stated goals:
- ✅ All workflows controlled from main dashboard UI
- ✅ No terminal prompts or input() dependencies
- ✅ No separate UI windows for Turnitin
- ✅ Real-time logging to dashboard terminal
- ✅ Direct function calls (no subprocess)
- ✅ Existing automation logic preserved
- ✅ Folder structure maintained

The application is now a proper desktop application with integrated UI control.
