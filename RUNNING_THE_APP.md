# Running the Enhanced RTF Tables App

## Quick Start

```bash
cd /Users/nikolasreppert/Documents/DATA_from_PDF
Rscript -e "shiny::runApp('rtf_tables_app.R')"
```

Or from R/RStudio:
```R
setwd('/Users/nikolasreppert/Documents/DATA_from_PDF')
shiny::runApp('rtf_tables_app.R')
```

## What Happens on Startup

1. **Dependency Check** (~5-10 seconds)
   - Verifies all R packages installed
   - Checks Python 3 availability
   - Checks for `pdfplumber` and `openpyxl`
   - Installs missing packages automatically

2. **App Launch**
   - All tabs available and functional
   - File list populated from `extracted_tables/` folder
   - Dependency status shown in "PDFs laden" tab

## New Features to Explore

### Help Tab
- First tab "Hilfe" contains complete user guide
- No internet needed - all help is built-in

### PDF Loading Tab
- Second tab "PDFs laden" 
- Upload PDFs directly from the app
- No command-line needed
- Automatic extraction to Excel format

### Dependency Display
- Check status of all required packages
- Install status shown with ✓ or ✗ symbols
- Automatic installation if needed

## File Organization

```
DATA_from_PDF/
├── rtf_tables_app.R           ← Main app (UPDATED)
├── extract_tables.py          ← PDF extraction script
├── ENHANCEMENT_SUMMARY.md     ← Feature documentation
├── RUNNING_THE_APP.md         ← This file
├── extracted_tables/          ← Auto-created, contains extracted Excel files
│   └── page_index.csv         ← Maps tables to PDF pages
├── muster/                    ← Pattern/template storage
├── translations_custom.csv    ← Custom word translations
└── formatting_rules_de.csv    ← Custom number formatting
```

## Common Tasks

### Extract Tables from PDF
1. Go to "PDFs laden" tab
2. Click "Choose File" → Select PDF(s)
3. Click "Tabellen extrahieren"
4. Wait for completion message
5. Go to "Tabellen & Vorschau" to see extracted tables

### Export Single Table as RTF
1. Select table in "Tabellen & Vorschau"
2. Choose style and language
3. Click "Aktive Tabelle (RTF)"
4. File downloads automatically

### Export Multiple Tables
1. Go to "Batch-Export" tab
2. Select multiple tables using checkboxes
3. Choose style and language
4. Click "ZIP herunterladen"
5. Receive ZIP with all selected tables

### Save Your Settings
1. Go to "Muster" tab
2. Enter pattern name
3. Click "Muster speichern"
4. Load anytime with "Muster laden"

## Troubleshooting

### App Won't Start
```bash
# Check R packages
Rscript -e "pacman::p_load(shiny, readxl, flextable, officer, dplyr, stringr, DT, tools, htmltools, zip, colourpicker)"

# Check Python
python3 --version

# Check Python packages
python3 -c "import pdfplumber; print('OK')"
python3 -c "import openpyxl; print('OK')"
```

### PDFs Won't Extract
1. Check "PDFs laden" tab for dependency status
2. Ensure extract_tables.py is in the same directory as rtf_tables_app.R
3. Check PDF is valid and readable
4. Look for error messages in status box

### Python Packages Missing
- Automatically installed on first run
- If failed, manual install: `pip3 install pdfplumber openpyxl`

## Browser Requirements
- Modern browser (Chrome, Firefox, Safari, Edge)
- JavaScript enabled
- No special plugins needed

## Performance Tips
- Large PDFs (100+ pages) may take 1-2 minutes
- Multiple file uploads process sequentially
- Browser can handle large file lists (1000+ tables)

## Getting Help
- Click "Hilfe" tab in the app for comprehensive documentation
- Check ENHANCEMENT_SUMMARY.md for feature details
- Review comments in rtf_tables_app.R source code
