"""
Configuration file for Turnitin Assistant v5.8-server-ready
Clean, centralized configuration with stable profile paths

v5.8: Added python-dotenv support, env var path overrides, cross-platform APP_DATA_DIR.
      On Linux/VPS, APP_DATA_DIR uses ~/.turnitin-assistant/ instead of macOS Library path.

v5.7: Added AI rate limit / anti-spam config, model switching, cache config.

v5.6.6: Added .env loading for AI verifier config (gated, disabled by default)
"""
import os
import sys
from pathlib import Path
import shutil

# Project root
PROJECT_ROOT = Path(__file__).resolve().parent

# ── v5.8: Load .env using python-dotenv if available ──
try:
    from dotenv import load_dotenv
    _dotenv_path = PROJECT_ROOT / ".env"
    if _dotenv_path.exists():
        load_dotenv(dotenv_path=_dotenv_path, override=True)
except ImportError:
    # Fallback: manual parsing (legacy behavior)
    _env_file = PROJECT_ROOT / ".env"
    if _env_file.exists():
        try:
            with open(_env_file) as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        key, val = line.split("=", 1)
                        os.environ[key.strip()] = val.strip().strip('"').strip("'")
        except Exception:
            pass

# ── v5.8: Platform-specific APP_DATA_DIR ──
# macOS: ~/Library/Application Support/TurnitinAssistant
# Linux/VPS: ~/.turnitin-assistant/
_IS_MAC = sys.platform == "darwin"
if _IS_MAC:
    APP_DATA_DIR = Path.home() / "Library" / "Application Support" / "TurnitinAssistant"
else:
    APP_DATA_DIR = Path.home() / ".turnitin-assistant"

STABLE_PROFILES_DIR = APP_DATA_DIR / "profiles"
STABLE_TURNITIN_PROFILE_DIR = STABLE_PROFILES_DIR / "turnitin-profile"
STABLE_OJS_PROFILE_DIR = STABLE_PROFILES_DIR / "ojs-profile"

# === TURNITIN PROFILE PATH (persistent, shared across engines) ===
# v5.5.2: Always use project-local profile path consistently.
# All Turnitin browser actions use this same path.
TURNITIN_PROFILE_PATH = PROJECT_ROOT / "profiles" / "turnitin-profile"
OJS_PROFILE_PATH = PROJECT_ROOT / "profiles" / "ojs-profile"

# Old profile paths (inside project, for migration)
OLD_PROFILES_DIR = PROJECT_ROOT / "profiles"
OLD_TURNITIN_PROFILE_DIR = OLD_PROFILES_DIR / "turnitin-profile"
OLD_OJS_PROFILE_DIR = OLD_PROFILES_DIR / "ojs-profile"

# ── v5.8: Data paths with env var overrides ──
DATA_DIR = PROJECT_ROOT / "data"
_master_excel_env = os.environ.get("MASTER_EXCEL_PATH")
if _master_excel_env:
    MASTER_EXCEL = Path(_master_excel_env)
else:
    MASTER_EXCEL = DATA_DIR / "master update.xlsx"

# Storage paths
STORAGE_DIR = PROJECT_ROOT / "storage"
_ojs_downloads_env = os.environ.get("OJS_DOWNLOADS_DIR")
if _ojs_downloads_env:
    OJS_DOWNLOADS_DIR = Path(_ojs_downloads_env)
else:
    OJS_DOWNLOADS_DIR = STORAGE_DIR / "ojs-downloads"

_turnitin_reports_env = os.environ.get("TURNITIN_REPORTS_DIR")
if _turnitin_reports_env:
    TURNITIN_REPORTS_DIR = Path(_turnitin_reports_env)
else:
    TURNITIN_REPORTS_DIR = STORAGE_DIR / "turnitin-reports"

_screening_reports_env = os.environ.get("SCREENING_REPORTS_DIR")
if _screening_reports_env:
    SCREENING_REPORTS_DIR = Path(_screening_reports_env)
else:
    SCREENING_REPORTS_DIR = STORAGE_DIR / "template_screening_reports"

SCREENSHOTS_DIR = STORAGE_DIR / "screenshots"

# UI paths
UI_DIR = PROJECT_ROOT / "ui"
UI_INDEX = UI_DIR / "index.html"

# Engine paths
ENGINES_DIR = PROJECT_ROOT / "engines"
OJS_ENGINE = ENGINES_DIR / "ojs_download_engine.py"
TURNITIN_ENGINE = ENGINES_DIR / "turnitin_upload_engine.py"
OJS_REPORT_UPLOAD_ENGINE = ENGINES_DIR / "ojs_report_upload_engine.py"
TEMPLATE_SCREENING_ENGINE = ENGINES_DIR / "template_screening_engine.py"

# Template Screening paths
TEMPLATE_RULES_DIR = PROJECT_ROOT / "config" / "template_rules"
TEMPLATE_SCREENING_CONFIG = {
    "rules_file": TEMPLATE_RULES_DIR / "jurnal_teknologi_rules.json",
    "reports_dir": str(SCREENING_REPORTS_DIR),
    "master_excel": str(MASTER_EXCEL),
    "conversion_enabled": True,
    "screening_sheet": "screening",
    "download_sheet": "download",
}


def migrate_profile_if_needed(old_path, new_path, profile_name):
    """
    Migrate profile from old location to new stable location.
    Only migrates if:
    - Old profile exists and has content
    - New profile does not exist or is empty
    """
    if not old_path.exists():
        return False
    
    # Check if old profile has meaningful content
    old_has_content = False
    try:
        # Check for cookies database or other session files
        if (old_path / "Default" / "Cookies").exists():
            old_has_content = True
        elif (old_path / "Default" / "Network" / "Cookies").exists():
            old_has_content = True
        elif any(old_path.rglob("*")):  # Has any files
            old_has_content = True
    except Exception:
        pass
    
    if not old_has_content:
        print(f"[MIGRATION] Old {profile_name} profile exists but is empty, skipping migration")
        return False
    
    # Check if new profile already has content
    new_has_content = False
    if new_path.exists():
        try:
            if (new_path / "Default" / "Cookies").exists():
                new_has_content = True
            elif (new_path / "Default" / "Network" / "Cookies").exists():
                new_has_content = True
            elif any(new_path.rglob("*")):  # Has any files
                new_has_content = True
        except Exception:
            pass
    
    if new_has_content:
        print(f"[MIGRATION] New {profile_name} profile already has content, skipping migration")
        return False
    
    # Perform migration
    try:
        print(f"[MIGRATION] Migrating {profile_name} profile:")
        print(f"[MIGRATION]   FROM: {old_path}")
        print(f"[MIGRATION]   TO:   {new_path}")
        
        # Ensure parent directory exists
        new_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Copy entire profile directory
        if new_path.exists():
            shutil.rmtree(new_path)
        shutil.copytree(old_path, new_path)
        
        print(f"[MIGRATION] ✅ {profile_name} profile migrated successfully")
        print(f"[MIGRATION] Old profile NOT deleted (kept for safety)")
        return True
    except Exception as e:
        print(f"[MIGRATION] ❌ Failed to migrate {profile_name} profile: {e}")
        return False


def ensure_directories():
    """Ensure all required directories exist"""
    from pathlib import Path
    reports_dir = Path(TEMPLATE_SCREENING_CONFIG["reports_dir"])
    dirs = [
        DATA_DIR,
        OJS_DOWNLOADS_DIR,
        TURNITIN_REPORTS_DIR,
        SCREENSHOTS_DIR,
        STABLE_PROFILES_DIR,
        STABLE_TURNITIN_PROFILE_DIR,
        STABLE_OJS_PROFILE_DIR,
        UI_DIR,
        ENGINES_DIR,
        TEMPLATE_RULES_DIR,
        reports_dir,
        reports_dir / "json",
        reports_dir / "pdf",
        reports_dir / "combined",
        reports_dir / "converted_pdf",
        reports_dir / "crops",
        reports_dir / "ai_cache",
        reports_dir / "pass",
        reports_dir / "needs_review",
        reports_dir / "reject",
    ]
    for directory in dirs:
        directory.mkdir(parents=True, exist_ok=True)
    
    print("[CONFIG] All directories ensured (v5.7 folder structure)")
    
    # Perform profile migrations if needed
    migrate_profile_if_needed(
        OLD_TURNITIN_PROFILE_DIR,
        STABLE_TURNITIN_PROFILE_DIR,
        "Turnitin"
    )
    migrate_profile_if_needed(
        OLD_OJS_PROFILE_DIR,
        STABLE_OJS_PROFILE_DIR,
        "OJS"
    )


# OJS Configuration
OJS_CONFIG = {
    "submissions_url": "https://journals.utm.my/jurnalteknologi/submissions",
    "profile_dir": str(STABLE_OJS_PROFILE_DIR),
    "download_dir": str(OJS_DOWNLOADS_DIR),
    "master_excel": str(MASTER_EXCEL),
    "max_scan_pages": 50,
    "max_download": 1,
    "items_per_page": 30,
    "stuck_seconds": 60,
    "human_slow_mo": 900
}

# Turnitin Configuration
TURNITIN_CONFIG = {
    "login_url": "https://www.turnitin.com/login_page.asp",
    "home_url": "https://www.turnitin.com/t_home.asp",
    "profile_dir": str(STABLE_TURNITIN_PROFILE_DIR),
    "screenshot_dir": str(SCREENSHOTS_DIR),
    "download_dir": str(TURNITIN_REPORTS_DIR),
    "master_excel": str(MASTER_EXCEL),
    "journal_short": "JTI",
    "default_timeout": 15000,
    "navigation_timeout": 30000,
    "manual_login_timeout": 300000,
    "human_mode": "test",  # options: "test", "human"
    "exclude_keywords": [
        "universiti teknologi malaysia",
        "university teknologi malaysia",
        "university technology malaysia",
        "utm.my",
        "journals.utm.my",
        "penerbit utm",
        "utm press"
    ]
}

# OJS Report Upload Configuration
OJS_UPLOAD_CONFIG = {
    "submissions_url": "https://journals.utm.my/jurnalteknologi/submissions",
    "profile_dir": str(STABLE_OJS_PROFILE_DIR),
    "master_excel": str(MASTER_EXCEL),
    "stuck_seconds": 60,
    "human_slow_mo": 900,
    "turnitin_sheet": "turnitin",
    "upload_sheet": "upload",
}

# ── v5.7: .env loading for AI config ──
_env_file = PROJECT_ROOT / ".env"
if _env_file.exists():
    try:
        with open(_env_file) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    os.environ[key.strip()] = val.strip().strip('"').strip("'")
    except Exception:
        pass

# ── v5.7 AI Configuration ──
# Master switch
USE_AI_VERIFIER = os.environ.get("USE_AI_VERIFIER", "false").lower() == "true"

# API Key
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

# Model selection
AI_TEXT_MODEL = os.environ.get("AI_TEXT_MODEL", "gpt-4.1-mini")
AI_VISION_MODEL = os.environ.get("AI_VISION_MODEL", "gpt-4.1-mini")

# Rate limiting & anti-spam
AI_REQUEST_DELAY_SECONDS = float(os.environ.get("AI_REQUEST_DELAY_SECONDS", "2"))
AI_MAX_CALLS_PER_SUBMISSION = int(os.environ.get("AI_MAX_CALLS_PER_SUBMISSION", "8"))
AI_MAX_CALLS_PER_RUN = int(os.environ.get("AI_MAX_CALLS_PER_RUN", "80"))
AI_TIMEOUT_SECONDS = int(os.environ.get("AI_TIMEOUT_SECONDS", "30"))
AI_CACHE_ENABLED = os.environ.get("AI_CACHE_ENABLED", "true").lower() == "true"

# Snippet size
AI_VERIFIER_MAX_SNIPPET_CHARS = int(os.environ.get("AI_VERIFIER_MAX_SNIPPET_CHARS", "800"))

# Redundant fallback for USE_AI_VISION_VERIFIER (mapped to master switch)
USE_AI_VISION_VERIFIER = os.environ.get("USE_AI_VISION_VERIFIER", "false").lower() == "true"

# Application Configuration
APP_CONFIG = {
    "app_name": "Turnitin Assistant v5.7",
    "version": "5.7",
    "window_width": 1400,
    "window_height": 900
}

if __name__ == "__main__":
    ensure_directories()
    print("\n[CONFIG] Project Structure:")
    print(f"  Root: {PROJECT_ROOT}")
    print(f"  Data: {DATA_DIR}")
    print(f"  Storage: {STORAGE_DIR}")
    print(f"  Stable Profiles: {STABLE_PROFILES_DIR}")
    print(f"  Old Profiles: {OLD_PROFILES_DIR}")
    print(f"  UI: {UI_DIR}")
    print(f"  Engines: {ENGINES_DIR}")
    print("\n[CONFIG] Stable Profile Paths:")
    print(f"  Turnitin: {STABLE_TURNITIN_PROFILE_DIR}")
    print(f"  OJS: {STABLE_OJS_PROFILE_DIR}")
    print("\n[CONFIG] AI Config:")
    print(f"  USE_AI_VERIFIER: {USE_AI_VERIFIER}")
    print(f"  AI_TEXT_MODEL: {AI_TEXT_MODEL}")
    print(f"  AI_VISION_MODEL: {AI_VISION_MODEL}")
    print(f"  AI_REQUEST_DELAY_SECONDS: {AI_REQUEST_DELAY_SECONDS}")
    print(f"  AI_MAX_CALLS_PER_SUBMISSION: {AI_MAX_CALLS_PER_SUBMISSION}")
    print(f"  AI_MAX_CALLS_PER_RUN: {AI_MAX_CALLS_PER_RUN}")
    print(f"  AI_CACHE_ENABLED: {AI_CACHE_ENABLED}")