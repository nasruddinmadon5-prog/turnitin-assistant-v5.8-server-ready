# Academic Auto Bridge v5.1

## Quick Start

```bash
cd "/Users/nasruddinmadon/turnitin-assistant-v5.1"
python3 app.py
```

## What Changed in v5.1

### Window Title
- Updated from "Academic Auto Bridge v5" to "Academic Auto Bridge v5.1"

### Improved Button Flow
The main dashboard now has clearer messaging about what each button does:

1. **Open OJS Login** - Opens browser, waits for you to login
2. **Start OJS Download** - Runs OJS automation (will prompt in terminal)
3. **Open Turnitin Login** - Opens browser for login
4. **Start Turnitin Upload** - Opens Turnitin automation UI window
5. **Start Full Workflow** - KIV (placeholder message)

### How Buttons Work

#### OJS Workflow:
1. Click "Open OJS Login" (message appears but button doesn't activate the engine)
2. Click "Start OJS Download" - This opens the OJS browser and starts automation
   - You'll see terminal prompts to press ENTER after logging in
   - Follow the terminal instructions

#### Turnitin Workflow:
1. Click "Open Turnitin Login" (informational message)
2. Click "Start Turnitin Upload" - This opens a separate Turnitin UI window
   - The Turnitin engine uses its own webview window
   - Follow instructions in that window

### Original Engine Code Preserved

The original v5 engine code remains intact:
- `engines/ojs_download_engine.py` - Original OJS automation
- `engines/turnitin_upload_engine.py` - Original Turnitin automation

This ensures stability and preserves all the complex functionality.

### Safe Duplication

The original v5 folder is untouched:
- Original: `/Users/nasruddinmadon/turnitin-assistant-v5`
- New v5.1: `/Users/nasruddinmadon/turnitin-assistant-v5.1`

You can always go back to v5 if needed.

## Button Behavior Details

### ✅ Active Buttons

**1. Open OJS Login**
- Shows informational message
- Tells you to click "Start OJS Download" to actually begin

**2. Start OJS Download**  
- Opens OJS browser with persistent profile
- You must login manually when browser opens
- Press ENTER in terminal when instructed
- Downloads one Article Text file
- Updates Excel `download` sheet

**3. Open Turnitin Login**
- Shows informational message
- Tells you to click "Start Turnitin Upload" to actually begin

**4. Start Turnitin Upload**
- Opens separate Turnitin UI window (the original Turnitin interface)
- Follow instructions in that window to login and upload
- Reads from Excel `download` sheet
- Uploads to Turnitin, filters sources, downloads report
- Updates Excel `turnitin` sheet

### ⏸️ KIV Button

**5. Start Full Workflow**
- Shows KIV message
- Not yet implemented
- Use buttons 2 and 4 manually instead

## Known Behavior

### Terminal Interaction Still Required

The original engines require terminal interaction in some cases:
- **OJS:** Press ENTER after logging in (as instructed in terminal)
- **Turnitin:** Uses separate UI window (original behavior)

This was kept intentionally to preserve the working v5 code without breaking anything.

### Original v5 vs v5.1

| Feature | v5 | v5.1 |
|---------|-----|------|
| Window Title | "Academic Auto Bridge v5" | "Academic Auto Bridge v5.1" |
| Button Messages | Basic | Improved with instructions |
| Full Workflow Button | Shows both engines | Shows KIV message |
| Engine Code | Original | Original (preserved) |
| Stability | Proven | Same (code unchanged) |

## Rollback Instructions

If you need to go back to v5:

```bash
cd "/Users/nasruddinmadon/turnitin-assistant-v5"
python3 app.py
```

The v5 folder is completely untouched and ready to use.

## Files Modified in v5.1

1. **app.py** - Updated window title, improved button messages, KIV for full workflow
2. **README_v5.1.md** - This file (documentation)
3. **CHANGES_v5.1.md** - Detailed change log

**Original engine files are UNCHANGED from v5** - This ensures stability.

## Support

- Check terminal output for progress/errors
- OJS engine: Follow terminal prompts
- Turnitin engine: Follow UI window prompts
- Excel file must not be open in another program

---

**Version:** 5.1  
**Date:** May 19, 2026  
**Status:** Stable (uses original v5 engine code)
