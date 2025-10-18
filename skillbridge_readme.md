# SkillBridge Data Extractor

Automated tool to extract program data from [SkillBridge](https://skillbridge.osd.mil/locations.htm) for specified job families and export to Excel with location consolidation.

## 📋 Overview

This tool automatically:
- ✅ Extracts program listings from SkillBridge website
- ✅ Filters by job family (Sales and Related, Business and Financial Operations)
- ✅ Consolidates multiple locations for the same organization into single entries
- ✅ Exports to formatted Excel files and CSV
- ✅ Creates timestamped backups
- ✅ Handles updates without creating duplicates

## 🎯 Features

### Smart Extraction
- **Automatic dropdown detection** - Finds and selects the correct job family filter
- **Dynamic data loading** - Waits for JavaScript content to fully load
- **Multiple extraction methods** - Uses DataTables API with HTML fallback
- **Error handling** - Continues even if one job family fails

### Location Consolidation
Instead of multiple rows:
```
ABC Company | New York
ABC Company | Los Angeles  
ABC Company | Chicago
```

You get one row:
```
ABC Company | Chicago | Los Angeles | New York
```

### Professional Output
- Formatted Excel files with:
  - Color-coded headers (blue background, white text)
  - Auto-sized columns
  - Text wrapping enabled
  - Frozen header row
- CSV backups for easy importing
- Timestamped versions for history tracking

## 🔧 Requirements

### Software
- **Python 3.7+** - [Download Python](https://www.python.org/downloads/)
- **Google Chrome Browser** - [Download Chrome](https://www.google.com/chrome/)
- **ChromeDriver** (auto-installed by script)

### Python Packages
```bash
pip install selenium pandas openpyxl
```

Or use the requirements file:
```bash
pip install -r requirements.txt
```

## 📦 Installation

### Step 1: Install Python
- **Windows**: Download from [python.org](https://www.python.org/downloads/) and run installer
- **Mac**: `brew install python3` or download from python.org
- **Linux**: `sudo apt-get install python3 python3-pip`

### Step 2: Install Dependencies
Open terminal/command prompt and run:
```bash
pip install selenium pandas openpyxl
```

### Step 3: Verify Chrome Installation
Make sure Google Chrome is installed and up to date.

### Step 4: Download the Script
Save `web.py` (or whatever you named the script) to your desired folder.

## 🚀 Usage

### Basic Usage

1. **Open Terminal/Command Prompt**
   - Windows: Press `Win + R`, type `cmd`, press Enter
   - Mac: Press `Cmd + Space`, type `terminal`, press Enter
   - Linux: Press `Ctrl + Alt + T`

2. **Navigate to Script Folder**
   ```bash
   cd path/to/your/script
   ```

3. **Run the Script**
   ```bash
   python skillbridge_scraper.py
   ```

4. **Choose Run Mode**
   ```
   Select option (1 or 2, default=2): 2
   ```
   - **Option 1**: Headless mode (runs in background)
   - **Option 2**: Visible mode (watch it work - recommended for first run)

### What Happens Next

The script will:
1. ✅ Open Chrome browser
2. ✅ Navigate to SkillBridge website
3. ✅ Detect available job family options
4. ✅ Extract "Sales and Related" programs
5. ✅ Extract "Business and Financial Operations" programs
6. ✅ Consolidate duplicate locations
7. ✅ Save Excel and CSV files
8. ✅ Create timestamped backups
9. ✅ Show completion summary

### Expected Runtime
- First run: 2-3 minutes (including page load and data extraction)
- Subsequent runs: 1-2 minutes

## 📁 Output Files

All files are saved in the `output` folder:

### Main Files (Always Updated)
```
output/
├── SkillBridge_Sales_and_Related.xlsx
├── SkillBridge_Sales_and_Related.csv
├── SkillBridge_Business_and_Financial_Operations.xlsx
└── SkillBridge_Business_and_Financial_Operations.csv
```

### Backup Files (Timestamped)
```
output/
├── SkillBridge_Sales_and_Related_20251018_094322.xlsx
├── SkillBridge_Business_and_Financial_Operations_20251018_094322.xlsx
├── screenshot_Sales_and_Related.png
└── screenshot_Business_and_Financial_Operations.png
```

### Excel File Structure

| Column | Description |
|--------|-------------|
| Partner/Program/Agency | Organization name |
| Program Name | Name of the program |
| Location | All locations (consolidated with \|) |
| City | Cities (consolidated) |
| State | States (consolidated) |
| Job Family | Sales and Related OR Business and Financial Operations |
| Duration | Program duration |
| Contact Info | Contact details |
| Links | Application/program links |
| Last Updated | Extraction timestamp |

## 🔄 Updating Your Data

### Manual Update
Simply run the script whenever you want fresh data:
```bash
python skillbridge_scraper.py
```

The script will:
- Fetch latest data from SkillBridge
- Merge with existing files
- Remove duplicates
- Keep the most recent version

### Scheduled Updates

#### Windows (Task Scheduler)
1. Open Task Scheduler
2. Click "Create Basic Task"
3. Set trigger (e.g., "Weekly on Monday at 9:00 AM")
4. Action: "Start a program"
5. Program: `C:\Path\To\Python\python.exe`
6. Arguments: `C:\Path\To\web.py`
7. Click Finish

#### Mac/Linux (Cron)
```bash
# Edit crontab
crontab -e

# Add line to run every Monday at 9 AM
0 9 * * 1 cd /path/to/script && /usr/bin/python3 web.py
```

## ⚙️ Customization

### Change Job Families
Edit the `job_family_mappings` in `main()`:
```python
job_family_mappings = {
    'Healthcare': 'Healthcare Practitioners and Technical',
    'IT': 'Computer and Mathematical'
}
```

### Change Output Directory
Modify in `run_extraction()`:
```python
extractor.run_extraction(job_family_mappings, output_dir='my_data')
```

### Adjust Wait Times
If pages load slowly, increase wait times:
```python
time.sleep(5)  # Change to 10 for slower connections
```

### Run Completely Silent (Headless)
Edit the default in `main()`:
```python
headless = True  # Always run in background
```

## 🛠️ Troubleshooting

### Issue: "ChromeDriver not found"
**Solution:**
```bash
# Automatic installation
pip install webdriver-manager

# Then update script to use it (add at top):
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

# In setup_driver(), replace:
self.driver = webdriver.Chrome(options=chrome_options)
# With:
service = Service(ChromeDriverManager().install())
self.driver = webdriver.Chrome(service=service, options=chrome_options)
```

### Issue: "No data extracted"
**Possible causes:**
1. **Website is down** - Check if skillbridge.osd.mil is accessible in your browser
2. **Slow internet** - Increase wait times in the script
3. **Website structure changed** - Check debug files:
   - `debug_page_source.html`
   - `debug_initial_page.png`
   - `screenshot_*.png`

**Solution:**
- Run in visible mode (option 2) to watch what happens
- Check console output for specific error messages
- Verify website works in regular browser

### Issue: Module Not Found
```bash
# Install missing module
pip install selenium
pip install pandas
pip install openpyxl
```

### Issue: Chrome Version Mismatch
**Solution:**
1. Update Chrome browser to latest version
2. ChromeDriver will auto-update on next run

### Issue: "Consolidation error"
The script will still save your data, just without location consolidation. The CSV file will have all raw data.

## 📊 Sample Console Output

```
╔══════════════════════════════════════════════════════════════╗
║     SKILLBRIDGE DATA EXTRACTOR - DROPDOWN VERSION            ║
║     Extracts: Sales & Business/Financial Operations          ║
╚══════════════════════════════════════════════════════════════╝

🛠️  Setting up Chrome driver...
✅ Chrome driver setup complete

STEP 1: Checking available dropdown options...
======================================================================
✅ Found 25 job family options:
   - Sales and Related
   - Business and Financial Operations
   [...]

STEP 2: Verifying target job families...
======================================================================
✅ Found 'Sales and Related' as: 'Sales and Related'
✅ Found 'Business and Financial Operations' as: 'Business and Financial Operations'

STEP 3: Starting extraction...
======================================================================

======================================================================
  EXTRACTING: Sales and Related
======================================================================
🎯 Selecting industry: Sales and Related
✅ Selected: Sales and Related
⏳ Waiting for data to load...
✅ Data loaded - 13 visible rows
📊 Extracting data for Sales and Related...
✅ Extracted 10 rows via DataTables API

🔄 Consolidating locations...
   Before: 10 rows
   After: 8 rows
✅ Consolidated 2 duplicates

💾 Saving to output\SkillBridge_Sales_and_Related.xlsx...
✅ Saved 8 programs

======================================================================
  EXTRACTING: Business and Financial Operations
======================================================================
🎯 Selecting industry: Business and Financial Operations
✅ Selected: Business and Financial Operations
⏳ Waiting for data to load...
✅ Data loaded - 25 visible rows
📊 Extracting data...
✅ Extracted 20 rows via DataTables API

🔄 Consolidating locations...
   Before: 20 rows
   After: 15 rows
✅ Consolidated 5 duplicates

💾 Saving to output\SkillBridge_Business_and_Financial_Operations.xlsx...
✅ Saved 15 programs

======================================================================
  EXTRACTION COMPLETE!
======================================================================
✅ Sales and Related: 8 programs
✅ Business and Financial Operations: 15 programs

📁 Files saved to: C:\Users\YourName\Desktop\Python\output
======================================================================

✨ ALL DONE! Check the 'output' folder for your files.
```

## 📝 Important Notes

### Compliance
- ✅ Respects website terms of use
- ✅ Only accesses publicly available data
- ✅ Includes rate limiting (5-second delays)
- ✅ No aggressive scraping
- ✅ User-agent identifies as legitimate browser

### Data Accuracy
- Data is as current as the SkillBridge website
- Always verify critical information on the official website
- Timestamps show when data was extracted
- The tool does NOT modify or interpret data

### Duplicate Handling
- Automatic duplicate removal based on:
  - Organization name
  - Program name
  - Job family
- Most recent data is kept when duplicates found
- Location fields are consolidated (not removed)

## 🔍 Advanced Usage

### Extract Specific Job Family Only
Edit `job_family_mappings` to include only what you want:
```python
job_family_mappings = {
    'Sales and Related': 'Sales and Related'
}
```

### Programmatic Usage
```python
from web import SkillBridgeExtractor

# Create extractor
extractor = SkillBridgeExtractor(headless=True)

# Custom extraction
job_families = {
    'IT Programs': 'Computer and Mathematical',
    'Healthcare': 'Healthcare Practitioners and Technical'
}

extractor.run_extraction(job_families, output_dir='my_output')
```

### Export to Different Formats
The data is in pandas DataFrame format, so you can easily export to other formats:
```python
# After extraction, in the script:
df.to_json('output.json', orient='records')
df.to_html('output.html')
df.to_parquet('output.parquet')
```

## 🆘 Getting Help

### Debug Files
When issues occur, check these files:
- `debug_page_source.html` - Full HTML of the page
- `debug_initial_page.png` - Screenshot of initial page
- `screenshot_*.png` - Screenshots during extraction

### Console Output
The console provides detailed information:
- What was found on the page
- Each step of the extraction process
- Specific error messages
- Summary statistics

### Common Questions

**Q: How often should I run this?**  
A: Weekly or monthly, depending on how often SkillBridge updates their listings.

**Q: Can I extract all job families at once?**  
A: Yes! Just add them all to the `job_family_mappings` dictionary.

**Q: Will this work if SkillBridge changes their website?**  
A: The script is adaptive and tries multiple strategies. Minor changes should be handled automatically.

**Q: Can I run this on a server?**  
A: Yes! Use headless mode and ensure Chrome/ChromeDriver are installed.

**Q: Is there a limit to how much data I can extract?**  
A: No artificial limits. The script extracts all available data for the selected job families.

## 📞 Support

If you encounter issues:
1. Check the console output for error messages
2. Review debug files (HTML and screenshots)
3. Verify SkillBridge website is accessible
4. Ensure Chrome and Python are up to date
5. Check that all dependencies are installed

## 📄 Files in This Package

```
skillbridge-extractor/
├── web.py                          # Main extraction script
├── requirements.txt                # Python dependencies
├── README.md                       # This file
└── output/                         # Created on first run
    ├── *.xlsx                      # Excel output files
    ├── *.csv                       # CSV backup files
    └── screenshot_*.png            # Debug screenshots
```

## 🔄 Version History

**Current Version: 2.0**
- ✅ Automatic dropdown detection
- ✅ Fuzzy matching for job family names
- ✅ Location consolidation
- ✅ DataTables API integration
- ✅ Multiple extraction fallback methods
- ✅ Comprehensive error handling
- ✅ Debug file generation

**Version 1.0**
- Initial release
- Basic extraction functionality

## 📜 License

This tool is provided as-is for personal and educational use. Always comply with SkillBridge's terms of service when using this tool.

## 🙏 Acknowledgments

- Data source: [SkillBridge.osd.mil](https://skillbridge.osd.mil/)
- Built with: Selenium, Pandas, OpenPyXL

---

**Happy Extracting! 🚀**

For questions or issues, please review the troubleshooting section or check the debug files generated during execution.