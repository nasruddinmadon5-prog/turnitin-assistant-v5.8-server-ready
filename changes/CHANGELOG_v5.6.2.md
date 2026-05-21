# CHANGELOG v5.6.2

**Release Date:** 2026-05-20

## Summary

Fixed misleading "SCREENING FAILED" / "Screening Error" status shown in the UI when the workflow actually completed successfully. Separated **workflow status** (engine process) from **paper evaluation result** (screening score).

## Changes

### 1. UI Status Panel (Right Column)
- **Before:** `Screening Error` / `SCREENING FAILED`
- **After:** `Screening Completed` (in green) when workflow runs normally
- Added new **Screening Summary** panel showing:
  - `X Passed` (green)
  - `X Warning` (yellow/orange)
  - `X Failed` (red)
- Added dedicated `SCREENING` status tile alongside OJS / TURNITIN / OJS UPLOAD

### 2. System Output Log Wording
- **Before:** `[SCREENING] SUB-XXXXX: SCREENING FAILED, score 43`
- **After:** `[SCREENING] SUB-XXXXX: SCREENING FAILED (Score: 43)`
- Added inline comment clarifying this is a **paper evaluation result** NOT an engine failure

### 3. Status Logic Separation
- `SCREENING ERROR` now **only** shows for:
  - Engine crash / runtime exception
  - PDF parsing failure
  - DOCX to PDF conversion failure
  - File not found / unsupported format
- `SCREENING COMPLETED` now shows when:
  - Engine completed normally (even if some papers scored low)
  - No fatal/engine errors occurred

### 4. Version Labels Updated
- All visible version labels: `v5.6.1` → `v5.6.2`
  - `ui/index.html` (title, header, footer)
  - `app.py` (docstring, print banner, window title)
  - `config.py` (APP_CONFIG)
  - `engines/template_screening_engine.py` (docstring, PDF report footer)

### 5. Architecture Note
- Not modified: pipeline structure, automation flow, engine sequence, report generation architecture
- Only improved: wording, status mapping logic, workflow vs evaluation separation

## Files Modified

| File | Changes |
|------|---------|
| `ui/index.html` | Version labels, new SCREENING status tile, Screening Summary panel, `updateScreeningStatusFromLog()` rewritten |
| `app.py` | Version labels, docstring update |
| `config.py` | APP_CONFIG version update |
| `engines/template_screening_engine.py` | Version labels, screening log message wording, PDF report footer |

## Migration Notes

No migration required. This is a UI/logic-only update. Existing Excel data and screening reports remain unchanged.