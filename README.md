# CAR.org Report Workflow System

Complete automated workflow for downloading buyer's guide reports from CAR.org (California Association of Realtors) Power BI dashboards and generating updated monthly reports using AI.

---

## Setup Guide (First-Time Users)

Follow these steps if you've never used this tool before.

### Step 1: Install Python

**Mac:**
1. Open **Terminal** (press `Cmd + Space`, type "Terminal", press Enter)
2. Check if Python is installed by typing: `python3 --version`
3. If not installed, download from https://www.python.org/downloads/
4. Run the installer and follow the prompts

**Windows:**
1. Download Python from https://www.python.org/downloads/
2. Run the installer
3. **IMPORTANT:** Check the box "Add Python to PATH" during installation
4. Click "Install Now"

### Step 2: Install Google Chrome

The automation requires Google Chrome browser:
- Download from https://www.google.com/chrome/
- Install and make sure it's up to date

### Step 3: Download This Project

1. Download this project folder to your computer
2. Remember where you saved it (e.g., `~/Documents/my agent` or `C:\Users\YourName\Documents\my agent`)

### Step 4: Install Required Packages

1. Open **Terminal** (Mac) or **Command Prompt** (Windows)
2. Navigate to the project folder:
   ```bash
   # Mac example:
   cd ~/Documents/my\ agent

   # Windows example:
   cd C:\Users\YourName\Documents\my agent
   ```
3. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```
   Or if that doesn't work:
   ```bash
   pip3 install -r requirements.txt
   ```

### Step 5: Set Up OpenAI API Key (Required for AI Report Generation)

1. Go to https://platform.openai.com/api-keys
2. Sign up or log in to your OpenAI account
3. Click "Create new secret key"
4. Copy the key (starts with `sk-...`)
5. Set it as an environment variable:

**Mac (add to ~/.zshrc or ~/.bash_profile):**
```bash
echo 'export OPENAI_API_KEY="sk-your-key-here"' >> ~/.zshrc
source ~/.zshrc
```

**Windows (Command Prompt as Administrator):**
```cmd
setx OPENAI_API_KEY "sk-your-key-here"
```
Then restart your Command Prompt.

### Step 6: Run the Workflow

```bash
python workflow_startup.py
```
Or on Mac:
```bash
python3 workflow_startup.py
```

---

## Quick Start (Returning Users)

If you've already completed the setup above:

```bash
cd /path/to/project
python workflow_startup.py
```

---

## Core Files

| File | Description |
|------|-------------|
| `workflow_startup.py` | Main interactive 4-step workflow script (Chinese UI) |
| `final_working_downloader.py` | Manual Power BI county/city image downloader |
| `final_working_downloader_auto.py` | Fully automated Power BI downloader with auto county selection |
| `auto_download_county_file_sharepoint.py` | Auto-download county PNGs from SharePoint |
| `convert_zip_pdfs_to_png.py` | Convert MLSL ZIP PDFs to PNG format |
| `simple_report_updater.py` | AI-powered monthly report generator |
| `monthly report sample.txt` | Sample report template |

**Interactive 4-step guided workflow:**
1. MLSL BI portal tasks (login, reports, email, conversion)
2. Download PNG files from SharePoint for 4 counties
3. Run Power BI downloader (manual or auto mode)
4. Generate AI-updated monthly report

## Directory Organization

Creates organized monthly directories:
```
~/Downloads/SALA Report January 2026/
├── Santa Clara.png, San Mateo.png, Alameda.png, San Francisco.png (SharePoint)
├── 1.png, 2.png, 3.png, 4.png (automated county files from Power BI)
├── 1(1).png, 1(2).png, 2(1).png... (automated city files from Power BI)
├── 6(1).png, 6(2).png, 6(3).png... (MLSL reports)
└── monthly_report_updated.txt (AI-generated report)
```

## Target Counties

| Order | County | Files |
|-------|--------|-------|
| 1 | Santa Clara | `1.png`, `1(1).png`, `1(2).png`... |
| 2 | San Mateo | `2.png`, `2(1).png`, `2(2).png`... |
| 3 | Alameda | `3.png`, `3(1).png`, `3(2).png`... |
| 4 | San Francisco | `4.png`, `4(1).png`, `4(2).png`... |

## Automated Downloader

The `final_working_downloader_auto.py` script provides fully automated county selection:

```bash
# Run with default counties (Santa Clara, San Mateo, Alameda, San Francisco)
python final_working_downloader_auto.py

# Customize target counties via environment variable
SALA_COUNTIES="Santa Clara,San Mateo" python final_working_downloader_auto.py

# Enable debug mode
SALA_DEBUG=1 python final_working_downloader_auto.py
```

Features:
- Automatic Power BI slicer interaction
- Search input and listbox navigation
- Scrollable list handling
- SharePoint image link extraction
- Robust retry logic

## AI Report Features

- **Auto month increment**: 7月 → 8月
- **Data extraction**: Analyzes PNG images for latest market data
- **Smart percentages**: Uses "4%" not "4.0%", "与上月持平" for no change
- **Chinese language**: Proper terms (上升/下降/增加/减少)
- **Diff preview**: Shows changes before replacing sample file
- **Backup creation**: Creates `.backup` file before replacing

## Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `SALA_WORK_DIR` | Custom work directory | `~/Downloads/SALA Report {Month Year}` |
| `SALA_COUNTIES` | Comma-separated county list | Santa Clara, San Mateo, Alameda, San Francisco |
| `SALA_DEBUG` | Enable debug output (1/true/yes) | disabled |
| `OPENAI_API_KEY` | Required for AI report generation | - |

## Workflow Details

All steps default to **automatic mode** for a streamlined experience.

### Step 1: MLSL BI Portal
- Auto-opens MLSL BI portal for login
- User downloads ZIP report PDFs manually
- Auto-converts PDFs to PNG (skip option available)

### Step 2: SharePoint Download
- Auto-downloads 4 county PNG files from CAR.org SharePoint
- Falls back to manual browser download if auto-download fails

### Step 3: Power BI Downloader
- **Default: Auto mode** - Script automatically selects counties via UI automation
- Manual mode available if needed (press 'm' when prompted)
- Downloads both county-level and city-level images

### Step 4: AI Report Generation
- Auto-launches AI report updater
- Analyzes downloaded PNG files
- Generates updated monthly report in Chinese
- Shows diff preview before replacing sample file

## Technical Notes

- Uses Selenium WebDriver for browser automation
- Handles Power BI iframes and dynamic content
- SharePoint download via direct URL extraction
- Supports both authenticated and public SharePoint links

---

## Troubleshooting

### "command not found: python"
Try using `python3` instead of `python`:
```bash
python3 workflow_startup.py
```

### "pip: command not found"
Try using `pip3` instead of `pip`:
```bash
pip3 install -r requirements.txt
```

### Chrome browser opens but automation fails
1. Make sure Chrome is updated to the latest version
2. The script will auto-download the correct ChromeDriver

### "OPENAI_API_KEY not set" error
Make sure you've set up the API key (see Step 5 in Setup Guide), then restart your terminal.

### Files not downloading to the right folder
Check the work directory shown at the start of the workflow. By default, files go to:
```
~/Downloads/SALA Report {Month Year}/
```

### "ModuleNotFoundError: No module named 'selenium'"
Run the package installation again:
```bash
pip install -r requirements.txt
```

### Power BI page doesn't load correctly
1. Wait a few seconds for the page to fully load
2. If using auto mode and it fails, try manual mode (press 'm')

---

## Need Help?

If you encounter issues not covered above:
1. Check that all prerequisites are installed
2. Make sure you're in the correct project directory
3. Try running the workflow again - some network issues are temporary