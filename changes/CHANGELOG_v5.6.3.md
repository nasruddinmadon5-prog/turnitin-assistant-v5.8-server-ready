# CHANGELOG v5.6.3

**Release Date:** 2026-05-20

## Summary

Combined the screening report PDF and the original article PDF into a **single merged PDF file**. No more split files — the output is one complete document: screening report pages first, followed by the original article pages.

## Changes

### 1. Merged PDF Output
- **Before:** Two separate files
  - `template_screening_report_XXXXX.pdf` (screening analysis only)
  - Original article PDF (kept separately)
- **After:** Single combined file
  - `template_screening_combined_XXXXX.pdf`
  - **Part 1:** Template Screening Report (analysis, scores, findings, conclusion)
  - **Part 2:** Original Article (full manuscript pages appended)

### 2. How Merging Works
- Report PDF is generated first using reportlab (as before)
- Original article PDF (or converted DOCX→PDF) is appended using PyMuPDF (`fitz`)
- If PyMuPDF is unavailable, falls back to report-only PDF
- If article PDF file is missing, outputs report-only PDF with a warning log

### 3. Version Labels Updated
- All version labels: `v5.6.2` → `v5.6.3`
  - `ui/index.html` (title, header, terminal init, footer)
  - `app.py` (docstring, startup banner, window title)
  - `config.py` (APP_CONFIG)
  - `engines/template_screening_engine.py` (docstring, PDF report footer)
  - `changes/CHANGELOG_v5.6.3.md`

## Files Modified

| File | Changes |
|------|---------|
| `engines/template_screening_engine.py` | `generate_screening_pdf_report()` rewritten: temporary report → merge with article → single combined PDF |
| `ui/index.html` | Version labels |
| `app.py` | Version labels, docstring, startup banner |
| `config.py` | `APP_CONFIG` version update |

## Files Added

| File | Description |
|------|-------------|
| `changes/CHANGELOG_v5.6.3.md` | This changelog |

## Architecture Note
- The pipeline structure, automation flow, engine sequence remain unchanged
- Only the report generation function was modified
- The Excel column `template_screening_report_pdf` now stores the path to the combined PDF
- Filename format changed from `template_screening_report_*` to `template_screening_combined_*`

## Dependencies
- Requires `PyMuPDF` (fitz) for PDF merging — already used by the screening engine for PDF parsing
- Falls back gracefully if unavailable

## Migration Notes
- No migration required for existing data
- New screening runs will produce combined PDFs automatically
- Old separate report PDFs remain in their existing locations