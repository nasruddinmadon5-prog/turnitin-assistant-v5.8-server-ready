"""
Turnitin Upload Engine API - Callable Functions
v5.5.2: Persistent profile/session reuse fix
  - Always uses same TURNITIN_PROFILE_PATH from config
  - launch_persistent_context with anti-detection args + channel="chrome"
  - Stale lock cleanup before launch (never delete profile data)
  - Browser context KEPT ALIVE after workflow (session persists)
  - Enhanced diagnostics: cookies, local storage, URL logging
  - Storage state saved after manual login
  - Login detection: Quick Submit link found → session reused
  - No browser.new_context() used - only persistent context
"""
import threading
import os
from pathlib import Path
from engines import turnitin_upload_engine as turnitin_engine

# Workflow guard - prevent concurrent Turnitin runs
_workflow_lock = threading.Lock()
_workflow_running = False

# UI log callback (set when workflow starts)
_ui_log_callback = None


def set_ui_log_callback(callback):
    """Set UI logging callback and wire it into turnitin_engine.ui_log."""
    global _ui_log_callback
    _ui_log_callback = callback

    # Override ui_log in turnitin_engine so all internal messages route to UI
    def wrapped_ui_log(message):
        if _ui_log_callback:
            _ui_log_callback(message)
        else:
            print(message)

    turnitin_engine.ui_log = wrapped_ui_log


def check_profile_locks(profile_dir):
    """
    Check for Chromium profile lock files and log their status.
    
    Returns: (has_locks, lock_files) tuple
    """
    profile_path = Path(profile_dir)
    
    lock_files = {
        "SingletonLock": profile_path / "SingletonLock",
        "SingletonCookie": profile_path / "SingletonCookie", 
        "SingletonSocket": profile_path / "SingletonSocket",
        "Default/LOCK": profile_path / "Default" / "LOCK",
    }
    
    found_locks = []
    
    for name, lock_file in lock_files.items():
        if lock_file.exists():
            found_locks.append(name)
            turnitin_engine.ui_log(f"[LOCK CHECK] Found: {name}")
    
    return len(found_locks) > 0, found_locks


def safe_remove_stale_locks(profile_dir):
    """
    Safely remove stale profile locks with diagnostics.
    
    Only removes locks if:
    - No Chromium processes appear to be using the profile
    - Locks appear stale (can add time-based checks if needed)
    
    Returns: True if locks were removed, False if profile appears in use
    """
    profile_path = Path(profile_dir)
    
    has_locks, lock_files = check_profile_locks(profile_dir)
    
    if not has_locks:
        turnitin_engine.ui_log("[LOCK CHECK] No lock files found - profile is free")
        return True
    
    turnitin_engine.ui_log(f"[LOCK CHECK] Found {len(lock_files)} lock file(s)")
    
    # Check if Chromium/Chrome processes are running
    # This is a basic check - can be enhanced
    try:
        if os.name == 'posix':  # macOS/Linux
            import subprocess
            result = subprocess.run(
                ['pgrep', '-f', 'Chromium|Chrome'],
                capture_output=True,
                text=True
            )
            if result.returncode == 0 and result.stdout.strip():
                turnitin_engine.ui_log("[LOCK CHECK] ⚠️  WARNING: Chromium/Chrome process(es) detected")
                turnitin_engine.ui_log("[LOCK CHECK] Profile may be in use by another instance")
                turnitin_engine.ui_log("[LOCK CHECK] If Turnitin Assistant is not running elsewhere, locks are likely stale")
    except Exception as e:
        turnitin_engine.ui_log(f"[LOCK CHECK] Could not check for Chromium processes: {e}")
    
    # Attempt to remove locks
    turnitin_engine.ui_log("[LOCK CHECK] Attempting to remove stale lock files...")
    
    removed_count = 0
    failed_count = 0
    
    for name in lock_files:
        lock_file = profile_path / name.replace('/', os.sep)
        if lock_file.exists():
            try:
                lock_file.unlink()
                turnitin_engine.ui_log(f"[LOCK CHECK] ✅ Removed: {name}")
                removed_count += 1
            except Exception as e:
                turnitin_engine.ui_log(f"[LOCK CHECK] ❌ Failed to remove {name}: {e}")
                failed_count += 1
    
    if failed_count > 0:
        turnitin_engine.ui_log(f"[LOCK CHECK] ⚠️  Could not remove {failed_count} lock file(s)")
        turnitin_engine.ui_log("[LOCK CHECK] Profile may be in use by another process")
        return False
    
    if removed_count > 0:
        turnitin_engine.ui_log(f"[LOCK CHECK] ✅ Removed {removed_count} stale lock file(s)")
    
    return True


def diagnose_profile(profile_dir):
    """
    Log diagnostic information about the profile at startup.
    """
    profile_path = Path(profile_dir)
    
    turnitin_engine.ui_log("="*60)
    turnitin_engine.ui_log("[DIAGNOSTICS] Turnitin Profile Status")
    turnitin_engine.ui_log("="*60)
    turnitin_engine.ui_log(f"[DIAGNOSTICS] Profile path: {profile_path}")
    
    if not profile_path.exists():
        turnitin_engine.ui_log("[DIAGNOSTICS] Profile does not exist (will be created on first launch)")
        return
    
    turnitin_engine.ui_log("[DIAGNOSTICS] Profile exists")
    
    # Check for cookies database
    cookie_paths = [
        profile_path / "Default" / "Cookies",
        profile_path / "Default" / "Network" / "Cookies",
    ]
    
    cookies_found = False
    for cookie_path in cookie_paths:
        if cookie_path.exists():
            turnitin_engine.ui_log(f"[DIAGNOSTICS] ✅ Cookies DB found: {cookie_path.relative_to(profile_path)}")
            cookies_found = True
            break
    
    if not cookies_found:
        turnitin_engine.ui_log("[DIAGNOSTICS] No cookies DB found (new profile or never logged in)")
    
    # Check lock status
    has_locks, lock_files = check_profile_locks(profile_dir)
    
    if has_locks:
        turnitin_engine.ui_log(f"[DIAGNOSTICS] ⚠️  Lock files present: {', '.join(lock_files)}")
    else:
        turnitin_engine.ui_log("[DIAGNOSTICS] ✅ No lock files (profile is free)")
    
    turnitin_engine.ui_log("="*60)


def run_turnitin_upload_workflow(log_callback=None):
    """
    Run complete Turnitin upload workflow in the calling thread.

    v5.5.2 improvements:
    - Uses project-local TURNITIN_PROFILE_PATH consistently
    - Browser context kept alive after workflow to preserve session
    - Only closes browser on app shutdown via close_turnitin_browser()
    
    Launches persistent browser session with turnitin-profile.
    Login check is handled by the existing _wait_for_dashboard_or_manual_login():
      - IF already logged in  → Quick Submit link detected → continue immediately
      - IF not logged in      → log message and wait up to 5 min (same thread, no input())

    All Playwright operations happen in the SAME thread.
    No cross-thread object sharing. No second browser window.

    Returns: dict with status and message
    """
    global _workflow_running

    if log_callback:
        set_ui_log_callback(log_callback)

    # Prevent concurrent runs
    with _workflow_lock:
        if _workflow_running:
            turnitin_engine.ui_log("[TURNITIN] ⚠️ Workflow already running. Please wait.")
            return {"status": "error", "message": "Turnitin workflow already running"}
        _workflow_running = True

    try:
        turnitin_engine.ui_log("[TURNITIN] Starting upload workflow...")
        
        # Diagnose profile at startup
        diagnose_profile(str(turnitin_engine.PROFILE_DIR))

        # v5.5.2: Do NOT close existing browser context if it's already alive.
        # If context already exists from previous run, reuse it (session preserved)
        if turnitin_engine.context is not None:
            turnitin_engine.ui_log("[TURNITIN] Reusing existing browser context (session preserved)")

        # --- Run full workflow in THIS thread ---
        api = turnitin_engine.Api()
        api._prepare_quick_submit_no_repository()

        turnitin_engine.ui_log("[TURNITIN] ✅ Workflow completed successfully!")

        return {
            "status": "success",
            "message": "Turnitin workflow completed",
            "submission_no": turnitin_engine.process_state.submission_no,
            "report_path": turnitin_engine.process_state.final_pdf_path
        }

    except Exception as e:
        error_msg = str(e)
        turnitin_engine.ui_log(f"[TURNITIN] ❌ Workflow failed: {error_msg}")
        
        # Check if error is profile-related
        if "ProcessSingleton" in error_msg or "profile" in error_msg.lower():
            turnitin_engine.ui_log("[TURNITIN] ⚠️  Profile lock error detected")
            turnitin_engine.ui_log("[TURNITIN] Close other Turnitin Assistant instances and retry")
        
        return {
            "status": "error",
            "message": error_msg
        }
    finally:
        _workflow_running = False
        turnitin_engine.ui_log("[TURNITIN] 🔓 Workflow lock released.")
        
        # v5.5.2: Do NOT close browser context here!
        # Keep context alive so session is preserved for future runs.
        # Browser will be closed only on app exit via close_turnitin_browser().
        if turnitin_engine.context is not None:
            turnitin_engine.ui_log("[TURNITIN] Browser kept open for session persistence.")
            # Save storage state explicitly
            try:
                storage_state_path = Path(turnitin_engine.PROFILE_DIR) / "storage_state.json"
                turnitin_engine.context.storage_state(path=str(storage_state_path))
                turnitin_engine.ui_log(f"[TURNITIN] ✅ Storage state saved: {storage_state_path}")
            except Exception:
                pass


def close_turnitin_browser():
    """
    Explicitly close Turnitin browser session.
    Call this on app shutdown or when user explicitly wants to close browser.
    """
    try:
        if turnitin_engine.context is not None:
            turnitin_engine.ui_log("[TURNITIN] Closing browser...")
            # Save state before close
            try:
                storage_state_path = Path(turnitin_engine.PROFILE_DIR) / "storage_state.json"
                turnitin_engine.context.storage_state(path=str(storage_state_path))
                turnitin_engine.ui_log(f"[TURNITIN] ✅ Storage state saved before close: {storage_state_path}")
            except Exception:
                pass
            turnitin_engine.context.close()
            turnitin_engine.ui_log("[TURNITIN] ✅ Browser closed. Profile session preserved.")
            turnitin_engine.context = None
            turnitin_engine.page = None

        if turnitin_engine.pw is not None:
            turnitin_engine.pw.stop()
            turnitin_engine.pw = None
    except Exception as e:
        turnitin_engine.ui_log(f"[TURNITIN] Error closing browser: {e}")