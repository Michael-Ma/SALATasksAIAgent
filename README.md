# CAR.org Report Workflow System

Complete automated workflow for downloading buyer's guide reports and generating updated monthly reports using AI.

## ✅ CORE FILES

- **`workflow_startup.py`** - Main interactive workflow script
- **`final_working_downloader.py`** - Power BI county/city image downloader
- **`simple_report_updater.py`** - AI-powered monthly report generator
- **`monthly report sample.txt`** - Sample report template
- **`README.md`** - This documentation

## 🚀 Quick Start

### Prerequisites
```bash
pip install -r requirements.txt
```

### Run Workflow
```bash
python workflow_startup.py
```

**Interactive 5-step guided workflow:**
1. Download PNG files from CAR.org Marketing and SharePoint
2. Prepare automated Power BI downloader
3. Run Power BI downloader (downloads county/city images)
4. MLSL BI portal tasks (bilingual instructions)
5. Generate AI-updated monthly report

## 📁 Directory Organization

Creates organized monthly directories:
```
~/Downloads/SALA Report January 2025/
├── 1Santa Clara.png, 2San Mateo.png, 3Alameda.png, 4San Francisco.png (manual)
├── 1.png, 2.png, 3.png, 4.png (automated county files)
├── 1(1).png, 1(2).png, 2(1).png... (automated city files)
├── 6(1).png, 6(2).png, 6(3).png... (MLSL reports)
└── monthly_report_updated.txt (AI-generated report)
```

## 🎯 Target Counties

1. **Santa Clara County** → `1*.png`
2. **San Mateo County** → `2*.png`  
3. **Alameda County** → `3*.png`
4. **San Francisco County** → `4*.png`

## 🤖 AI Report Features

- **Auto month increment**: 7月 → 8月
- **Data extraction**: Analyzes PNG images for latest market data
- **Smart percentages**: Uses "4%" not "4.0%", "与上月持平" for no change
- **Chinese language**: Proper terms (上升/下降/增加/减少)
- **Diff preview**: Shows changes before replacing sample file

## 🌐 Bilingual Support

Step 5 includes both English and Chinese instructions (English + 中文) for MLSL BI portal tasks.

## ✅ Success Rate

- **Dual naming support** for manual (1Santa Clara.png) and automated (1.png) files
- **100% automation** for data download and report generation
- **Organized workflow** with progress tracking
- **Error resilient** with fallback options
- **User-friendly** interactive bilingual navigation

## 🎉 Complete Solution

End-to-end workflow from initial data download to final AI-generated monthly report, all organized in monthly directories with bilingual support.