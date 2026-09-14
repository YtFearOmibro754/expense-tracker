# Expense Tracker - Installation & Setup

## System Requirements

- **Python 3.7 or higher** (usually pre-installed on Mac/Linux)
- **pip** (Python package manager - comes with Python)
- **A spreadsheet application** (Excel, Google Sheets, LibreOffice, etc.)

## Platform-Specific Installation

### Windows Installation

**Step 1: Check if Python is installed**
```bash
python --version
```

If you see a version number, Python is installed. If not:
1. Download from [python.org](https://www.python.org/downloads/)
2. Run the installer
3. **IMPORTANT**: Check "Add Python to PATH"
4. Click Install

**Step 2: Install openpyxl**
```bash
pip install openpyxl
```

**Step 3: Generate the Excel file**
```bash
python create_tracker.py
```

**Step 4: Open the file**
- Find `Expense_Tracker.xlsx` in your current folder
- Double-click to open in Excel
- Or right-click → Open With → Choose your app

### Mac Installation

**Step 1: Check if Python is installed**
```bash
python3 --version
```

You should see Python 3.x.x. If not:
1. Download from [python.org](https://www.python.org/downloads/)
2. Run the installer
3. Follow the installation wizard

Or use Homebrew:
```bash
brew install python3
```

**Step 2: Install openpyxl**
```bash
pip3 install openpyxl
```

**Step 3: Generate the Excel file**
```bash
python3 create_tracker.py
```

**Step 4: Open the file**
- Find `Expense_Tracker.xlsx` in your current folder
- Double-click to open (will use Numbers, Excel, or Google Sheets)
- Or right-click → Open With → Choose your app

### Linux Installation

**Step 1: Check if Python is installed**
```bash
python3 --version
```

If not installed:
```bash
# Ubuntu/Debian
sudo apt-get install python3 python3-pip

# Fedora
sudo dnf install python3 python3-pip

# Arch
sudo pacman -S python python-pip
```

**Step 2: Install openpyxl**
```bash
pip3 install openpyxl
```

**Step 3: Generate the Excel file**
```bash
python3 create_tracker.py
```

**Step 4: Open the file**
```bash
# Using LibreOffice Calc
libreoffice --calc Expense_Tracker.xlsx

# Or just open with your default app
xdg-open Expense_Tracker.xlsx
```

---

## One-Line Installation (All Platforms)

**Windows:**
```bash
pip install openpyxl && python create_tracker.py
```

**Mac/Linux:**
```bash
pip3 install openpyxl && python3 create_tracker.py
```

---

## What Gets Created

After running the script, you'll have:

```
Expense_Tracker.xlsx
├── Expenses Sheet
│   ├── Date
│   ├── Category
│   ├── Description
│   ├── Amount
│   ├── Payment Method
│   ├── Tags
│   └── Notes
├── Summary Sheet
│   ├── Total Expenses (auto-calculated)
│   └── By Category (auto-calculated)
└── Categories Sheet
    └── Pre-defined expense categories
```

---

## Verification

After running the script, you should see:
```
============================================================
✓ SUCCESS! Expense Tracker created
============================================================

File: Expense_Tracker.xlsx
Location: /path/to/your/folder/Expense_Tracker.xlsx

Features:
  ✓ Expenses sheet with sample data
  ✓ Auto-calculating Summary sheet
  ✓ Category breakdown with formulas
  ✓ Professional formatting and styling
  ✓ Frozen header rows for easy navigation
  ✓ Pre-formatted currency and date columns

Next Steps:
  1. Find Expense_Tracker.xlsx in your folder
  2. Open with Excel, Google Sheets, or LibreOffice
  3. Start tracking your expenses!

See QUICK_START.md for detailed usage instructions
============================================================
```

---

## Troubleshooting

### Windows: "'python' is not recognized"

**Solution 1**: Use `py` instead
```bash
py -m pip install openpyxl
py create_tracker.py
```

**Solution 2**: Reinstall Python
1. Uninstall Python
2. Reinstall from [python.org](https://www.python.org/downloads/)
3. **CHECK** "Add Python to PATH" during installation
4. Restart your terminal
5. Try again

### Mac: "command not found: python3"

```bash
# Install using Homebrew
brew install python3

# Or download from python.org and run the installer
```

### "ModuleNotFoundError: openpyxl"

```bash
# Reinstall openpyxl
pip install --upgrade openpyxl
# or
pip3 install --upgrade openpyxl
```

### File not created

1. Check that you're in the correct directory
   ```bash
   ls  # Mac/Linux
   dir # Windows
   ```

2. Verify Python and pip installation
   ```bash
   python --version  # Windows
   python3 --version # Mac/Linux
   ```

3. Try running with full path
   ```bash
   python3 /full/path/to/create_tracker.py
   ```

4. Check file permissions
   ```bash
   chmod +x create_tracker.py  # Mac/Linux
   ```

### Formulas not working in Excel

1. Make sure automatic calculation is enabled
   - **Excel**: File → Options → Formulas → Calculation Options → Automatic
   - **Google Sheets**: Should auto-calculate (no action needed)
   - **LibreOffice**: Tools → Options → LibreOffice Calc → Calculate → Automatic

2. Try pressing Ctrl+Shift+F9 (Windows) or Cmd+Shift+F9 (Mac) to recalculate

### File won't open

1. Try opening with Google Sheets (free, online)
   - Upload to [sheets.google.com](https://sheets.google.com)
   - Works on all devices

2. Install LibreOffice (free, open-source)
   - Download from [libreoffice.org](https://www.libreoffice.org/)
   - Works on Windows, Mac, Linux

3. Try OnlyOffice (free, online)
   - Works directly in browser
   - No installation needed

---

## Opening the File

### Windows
- **Default**: Double-click → Opens in Excel
- **Alternatives**: 
  - Right-click → Open With → Google Sheets
  - Right-click → Open With → LibreOffice Calc

### Mac
- **Default**: Double-click → Opens in Numbers or Excel
- **Alternatives**:
  - Right-click → Open With → Google Sheets
  - Right-click → Open With → LibreOffice Calc
  - Drag to Excel icon in Dock

### Linux
- **Terminal**: `libreoffice --calc Expense_Tracker.xlsx`
- **GUI**: Double-click (opens default app)
- **Alternatives**: Right-click → Open With → Choose app

---

## Cloud Setup (Recommended for Sync)

### Google Sheets (Free)
1. Go to [drive.google.com](https://drive.google.com)
2. Click "New" → "File Upload"
3. Select `Expense_Tracker.xlsx`
4. Right-click → "Open With" → "Google Sheets"
5. Now accessible from any device!

### OneDrive (Windows)
1. Save file to OneDrive folder
2. Access from phone, tablet, or another computer
3. Automatically backed up

### Dropbox
1. Save file to Dropbox folder
2. Access from any device
3. Version history available

### iCloud Drive (Mac)
1. Save file to iCloud Drive folder
2. Access from any Apple device
3. Automatically synced

---

## Next Steps

1. ✓ Python installed
2. ✓ openpyxl installed
3. ✓ Script run successfully
4. ✓ File created
5. → Open the file and start using it!

See [QUICK_START.md](QUICK_START.md) for usage instructions.

---

## Getting Help

If you run into issues:
1. Check this file (INSTALL.md)
2. Review QUICK_START.md for usage tips
3. Create an issue on GitHub with:
   - Your operating system (Windows 10/11, Mac OS, Linux version)
   - Python version (`python --version`)
   - Error message (exact text)
   - Steps to reproduce the problem

---

## Success! 🎉

You now have a fully functional expense tracker. See QUICK_START.md for usage instructions.
