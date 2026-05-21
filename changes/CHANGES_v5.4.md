# Turnitin Assistant v5.4 Stabilization Patch

## Overview

Version 5.4 is a targeted stabilization patch that standardizes Turnitin submission metadata and download filenames for consistency and reliability.

---

## Key Changes from v5.3.1

### 1. ✅ Direct Login URL (Already Stable)
- System uses `https://www.turnitin.com/login_page.asp` directly
- Avoids homepage redirects and intermediate pages
- Login form fields are verified before proceeding

### 2. ✅ Stabilized Filters & Settings Flow (Already Stable)
- Filters panel verification with real checkbox state checking
- Retry logic with max 3 attempts for Exclude Quotes and Exclude Bibliography
- Wait for overlay disappearance before applying changes
- Comprehensive logging for filter actions

### 3. ✅ Auto Exclude UTM Sources (Already Stable)
- Automatic detection of UTM sources containing:
  - "student paper"
  - "submitted to Universiti Teknologi Malaysia"
  - "Universiti Teknologi Malaysia"
- Exclusion workflow with verification
- Continues normally if no UTM sources found

### 4. ✅ Stabilized Current View Download (Already Stable)
- Race-style handling for three possible outcomes:
  - **Outcome A**: Direct PDF download
  - **Outcome B**: Classic report popup (newreport_classic.asp)
  - **Outcome C**: Recovered UUID PDF from downloads folder
- Non-blocking detection of all outcomes
- Retry logic with max 3 attempts

### 5. ✅ Corrupted Download Detection (Already Stable)
- PDF validation using %PDF- header check
- File extension verification
- File size validation
- Automatic retry on invalid downloads (max 3 retries)
- Detailed logging of validation results

### 6. ✅ **NEW: Standardized Download Filename**

**OLD FORMAT (v5.3.1):**
```
turnitin-report_26791_JTI_Desulfurization of Heavy Naphtha...pdf
```

**NEW FORMAT (v5.4):**
```
turnitin-report_26791.pdf
```

**Implementation:**
```python
# Line 344 in turnitin_upload_engine.py
# v5.4: Generate turnitin_name: ONLY submission_no (no title anymore)
turnitin_name = submission_no_str
```

**Benefits:**
- Cleaner filenames
- No truncation issues
- Easier file management
- Consistent naming across all submissions

### 7. ✅ **NEW: Standardized Turnitin Submission Metadata**

**OLD VALUES (v5.3.1):**
- First Name: Parsed from author name (e.g., "John")
- Last Name: Parsed from author name (e.g., "Doe")
- Submission Title: Full article title

**NEW VALUES (v5.4):**
- **First Name**: `Jurnal Teknologi`
- **Last Name**: `Editor Team`
- **Submission Title**: Submission number only (e.g., `26791`)

**Implementation:**
```python
# Lines 2061-2066 in turnitin_upload_engine.py
# v5.4: Use standardized metadata for Turnitin submission
first_name = "Jurnal Teknologi"
last_name  = "Editor Team"
title      = submission["submission_no"]
```

**Benefits:**
- Consistent author attribution across all submissions
- Simplified submission titles for easy identification
- Professional journal-level attribution
- Eliminates author name parsing issues

### 8. ✅ **NEW: Synced Excel `turnitin_name` Field**

The Excel `turnitin_name` field in the `turnitin` sheet now contains the same value as the Turnitin submission title: **the submission number only**.

**Excel Schema (unchanged):**
- Column A: Timestamp
- Column B: Submission No
- **Column C: Turnitin Name** (now contains submission_no only)
- Column D: Turnitin Score
- Column E: Turnitin Report Path
- Column F: Status

**Example:**
| Submission No | Turnitin Name |
|---------------|---------------|
| 26791         | 26791         |
| 26792         | 26792         |

---

## Technical Details

### File Modified
- `engines/turnitin_upload_engine.py`

### Lines Changed

**Change 1: turnitin_name Generation (Line 344)**
```python
# BEFORE (v5.3.1):
safe_title = title_str if title_str else submission_no_str
turnitin_name = f"{submission_no_str}_JTI_{safe_title}"

# AFTER (v5.4):
turnitin_name = submission_no_str
```

**Change 2: Upload Metadata (Lines 2061-2066)**
```python
# BEFORE (v5.3.1):
# Parsed author name and full article title

# AFTER (v5.4):
first_name = "Jurnal Teknologi"
last_name  = "Editor Team"
title      = submission["submission_no"]
```

### Download Filename Logic
The download filename uses `turnitin_name` from `process_state`:
```python
# Line 1148-1149
if process_state.turnitin_name:
    safe_name = sanitize_filename(process_state.turnitin_name)
    filename   = f"turnitin-report_{safe_name}.pdf"
```

Result: `turnitin-report_26791.pdf`

---

## Stability Features (Inherited from v5.3.1)

### Retry Mechanisms
- Download attempts: 3 retries
- Filter operations: 3 retries
- Element detection: Progressive backoff
- Page state verification before actions

### Download Handling
- Multiple selector fallback strategies
- Playwright locator click (primary)
- JavaScript DOM click (fallback)
- Text-based search (last resort)

### PDF Validation
```python
with open(save_path, "rb") as f:
    header = f.read(5)
if header != b"%PDF-":
    # Invalid - retry download
```

### Logging
- Comprehensive stage tracking
- Excel logging with status mapping
- Screenshot capture on failures
- Process metadata preservation

---

## Migration Notes

### For Users Upgrading from v5.3.1

1. **Filenames will change**: New downloads will use the simplified format `turnitin-report_<submission_no>.pdf`

2. **Turnitin interface changes**: Papers will appear as:
   - Author: "Jurnal Teknologi Editor Team"
   - Title: Submission number (e.g., "26791")

3. **Excel data**: The `turnitin_name` column will contain submission numbers only

4. **No action required**: The system automatically uses the new format for all new submissions

### Backward Compatibility

- Old v5.3.1 files remain unchanged
- Excel sheets are compatible (same schema)
- Download folder structure unchanged
- Profile data preserved

---

## Testing Recommendations

1. **Test Upload**: Verify Turnitin shows "Jurnal Teknologi Editor Team" as author
2. **Test Download**: Confirm filename is `turnitin-report_<submission_no>.pdf`
3. **Test Excel Logging**: Verify `turnitin_name` contains submission number only
4. **Test PDF Validation**: Confirm %PDF- header check works correctly
5. **Test UTM Exclusion**: Verify UTM sources are detected and excluded

---

## Version History

- **v5.4**: Standardized metadata and filenames (2026-05-19)
- **v5.3.1**: Stable release with race-style download handling
- **v5.3**: Enhanced Feedback Studio controller
- **v5.2**: UI alignment improvements
- **v5.1**: Major refactoring and Excel bridge implementation

---

## Support

For issues or questions about v5.4:
1. Check the comprehensive logging output
2. Review screenshots in `storage/screenshots/`
3. Verify Excel `turnitin` sheet for process status
4. Check download folder for PDF files

---

## Summary

Version 5.4 is a focused patch that:
- ✅ Simplifies download filenames to `turnitin-report_<submission_no>.pdf`
- ✅ Standardizes Turnitin author metadata to "Jurnal Teknologi Editor Team"
- ✅ Uses submission number as Turnitin submission title
- ✅ Syncs Excel `turnitin_name` field with submission number
- ✅ Maintains all stability features from v5.3.1
- ✅ Requires no user intervention or configuration changes

All existing stability features (direct login, filter verification, UTM exclusion, race-style download, PDF validation) remain active and unchanged.
