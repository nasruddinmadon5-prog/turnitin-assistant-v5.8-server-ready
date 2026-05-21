# Turnitin Assistant v5

**Clean, restructured version with dual-engine architecture**

## 🎯 Overview

Automated workflow for downloading manuscripts from OJS (Open Journal Systems) and uploading them to Turnitin with similarity report filtering.

## 📁 Structure

```
turnitin-assistant-v5/
├── app.py                          # Main dashboard application
├── config.py                       # Centralized configuration
├── requirements.txt                # Python dependencies
├── README.md                       # This file
│
├── ui/
│   └── index.html                  # Obsidian-themed dashboard UI
│
├── engines/
│   ├── ojs_download_engine.py      # OJS manuscript download engine
│   └── turnitin_upload_engine.py   # Turnitin upload & report engine
│
├── data/
│   └── master update.xlsx          # Central Excel tracking file
│                                    # - Sheet 'download': OJS downloads
│                                    # - Sheet 'turnitin': Turnitin results
│
├── storage/
│   ├── ojs-downloads/              # Downloaded manuscripts
│   ├── turnitin-reports/           # Similarity reports (PDFs)
│   └── screenshots/                # Debug screenshots
│
└── profiles/
    ├── ojs-profile/                # OJS browser profile
    └── turnitin-profile/           # Turnitin browser profile
```

## 🚀 Setup

### 1. Install Dependencies

```bash
cd /Users/nasruddinmadon/turnitin-assistant-v2/turnitin-assistant-v5

# Install Python packages
pip3 install -r requirements.txt

# Install Playwright browser
python3 -m playwright install chromium
```

### 2. Verify Structure

```bash
python3 config.py
```

This will create all required directories and show the project structure.

## 🎮 Usage

### Run Dashboard

```bash
python3 app.py
```

This opens the dual-panel dashboard with:
- **Left Panel**: OJS Download Engine
- **Right Panel**: Turnitin Upload Engine

### Workflow Options

**Option 1: Manual Step-by-Step**
1. Click "Open OJS Login" → Login manually → Start download
2. Click "Open Turnitin Login" → Login manually → Start upload

**Option 2: Full Automation**
- Click "Start Full Workflow" → OJS downloads first, then Turnitin uploads

### Run Engines Standalone

**OJS Engine (Terminal)**
```bash
python3 engines/ojs_download_engine.py
```

**Turnitin Engine (Terminal)**
```bash
python3 engines/turnitin_upload_engine.py
```

## 📊 Excel Format

### Sheet: `download` (OJS outputs)
| Column | Description |
|--------|-------------|
| A | timestamp |
| B | submission_no |
| C | author_name |
| D | title |
| E | download path |

### Sheet: `turnitin` (Turnitin outputs)
| Column | Description |
|--------|-------------|
| A | timestamp |
| B | submission_no |
| C | turnitin_name |
| D | turnitin_score |
| E | turnitin_report_path |
| F | status (BERJAYA/TAK BERJAYA) |

## 🔧 Configuration

Edit `config.py` to customize:
- Excel file path
- Download directories
- OJS/Turnitin URLs
- Timeout values
- UTM exclusion keywords

## 📝 Engine Features

### OJS Download Engine
- Auto-login with browser profile persistence
- Filters: Active submissions in "Submission" stage
- Bottom-to-top scanning (newest first)
- Excel deduplication (skips already downloaded)
- Logs to `download` sheet

### Turnitin Upload Engine
- Reads from `download` sheet automatically
- Auto-uploads with parsed author/title
- Waits for similarity report generation
- Filters UTM sources automatically
- Downloads filtered PDF reports
- Logs to `turnitin` sheet with BERJAYA/TAK BERJAYA status

## 🎨 UI Theme

**Obsidian Operations** - Dark theme optimized for data density and rapid operations.

## ⚠️ Important Notes

1. **Do NOT rename folder/phase structure** - Clean v5 structure only
2. **One UI only** - No manual upload, no sidebar versions
3. **Two engines** - Kept separate, not merged yet
4. **Browser profiles** - Separate for OJS and Turnitin
5. **Excel bridge** - Turnitin reads from OJS download sheet

## 🔄 Migration from v4

Key changes:
- No more `fasa1/fasa2/fasa3` folders
- Renamed Python files to meaningful names
- Centralized config in `config.py`
- Clean path structure under `storage/` and `profiles/`
- Single dashboard UI (Obsidian theme)

## 📞 Support

Report issues or request changes via the project maintainer.

---

**Version**: 5.0.0  
**Last Updated**: 2026-05-19
