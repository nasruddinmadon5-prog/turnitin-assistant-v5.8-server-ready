# Turnitin Assistant v5.5.2 - Persistent Profile/Session Reuse

## CHANGELOG

**v5.5.2 - Fixed Turnitin persistent profile/session reuse so login is preserved across runs.**

## Root Cause

The Turnitin browser profile was not being persisted correctly across runs because:

1. **Profile path inconsistency**: `turnitin_upload_engine.py` defined its own `PROFILE_DIR` pointing to `<project>/profiles/turnitin-profile` but didn't import from `config.py`'s shared constant.

2. **Context closure after workflow**: `turnitin_engine_api.py` closed the browser context in both the `finally` block after each workflow run. This discarded any session data that was established during the run.

3. **Missing browser args**: `launch_persistent_context()` did not use `--disable-blink-features=AutomationControlled` or `--start-maximized` args, and didn't specify `channel="chrome"`.

4. **No storage state saving**: After manual login, the session was not explicitly saved via `context.storage_state()`.

5. **No stale lock cleanup**: Lock files from previous crashed runs could prevent the profile from being used.

## Changes Made

### `config.py`
- Added `TURNITIN_PROFILE_PATH = PROJECT_ROOT / "profiles" / "turnitin-profile"` as the canonical profile path constant
- Added `OJS_PROFILE_PATH` for consistency
- Updated version to 5.5.2

### `engines/turnitin_upload_engine.py`
- **Profile path**: Imports `TURNITIN_PROFILE_PATH` from `config` with fallback to local path
- **`_ensure_browser()`**: Added diagnostics logging before launch:
  - Profile path exists check
  - Cookies DB exists check  
  - Local Storage exists check
- **`_ensure_browser()`**: Added stale lock file cleanup before launching:
  - SingletonLock, SingletonCookie, SingletonSocket, Default/LOCK
  - Only removes stale locks (no Chrome process using profile)
- **`_ensure_browser()`**: Updated `launch_persistent_context()` with:
  - `channel="chrome"` (uses system Chrome if available)
  - `args=["--disable-blink-features=AutomationControlled", "--start-maximized", "--no-sandbox"]`
  - These prevent Turnitin from detecting automation and ensure full window
- **`_wait_for_dashboard_or_manual_login()`**: Added:
  - Current URL logging after opening Turnitin
  - "Session reused!" message when Quick Submit link found
  - `context.storage_state()` save after manual login

### `engines/turnitin_engine_api.py`
- **`run_turnitin_upload_workflow()`**: 
  - **CRITICAL FIX**: Removed `context.close()` from the `finally` block
  - Browser context is now kept alive after workflow completes
  - Storage state saved after each workflow in `finally`
- **`close_turnitin_browser()`**: Added storage state save before closing browser
- **Version header updated to v5.5.2**

### `app.py`
- Updated version reference from 5.5.1 to 5.5.2
- Updated console banner to reflect v5.5.2 changes

## Technical Notes

- The persistent context (`launch_persistent_context`) uses the same profile folder path every time
- Cookies, Local Storage, Session Storage are never deleted or cleared
- Lock files are only removed if stale (no active Chrome process using the profile)
- Browser context stays open between workflow runs to preserve session
- On app shutdown, `atexit` handler calls `close_turnitin_browser()` which saves state then closes
- First run: user may need to login manually once
- Second run onward: Turnitin opens already logged in automatically

## Files Changed

- `app.py`
- `config.py`
- `engines/turnitin_engine_api.py`
- `engines/turnitin_upload_engine.py`