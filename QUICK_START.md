# Expense Tracker - Quick Start Guide

## What You Get

A **fully functional Excel-based expense tracker** that's ready to use immediately on Windows, Mac, and Linux.

## Installation (60 seconds)

### Step 1: Install Python (if not already installed)
- Download from [python.org](https://www.python.org/downloads/)
- Ensure Python 3.7+ is installed
- **Mac Tip**: Python should be pre-installed. Open Terminal and type `python3 --version`
- **Windows Tip**: During installation, check "Add Python to PATH"

### Step 2: Install openpyxl

**Windows:**
```bash
pip install openpyxl
```

**Mac:**
```bash
pip3 install openpyxl
```

**Linux:**
```bash
pip3 install openpyxl
```

### Step 3: Generate Your Excel File

**Windows:**
```bash
python create_tracker.py
```

**Mac:**
```bash
python3 create_tracker.py
```

**Linux:**
```bash
python3 create_tracker.py
```

### Step 4: Open & Use
- Find `Expense_Tracker.xlsx` in your folder
- **Windows**: Double-click the file → Opens in Excel
- **Mac**: Double-click the file → Opens in Numbers or Excel
- **Alternative**: Upload to Google Sheets (works everywhere)
- Start tracking expenses!

---

## Features

### 📊 Pre-Built Sheets
1. **Expenses** - Main tracking sheet with sample data
2. **Summary** - Auto-calculating totals by category
3. **Categories** - Reference list of expense types

### 💰 Auto-Calculating Features
- Total expense sum (auto-updates as you add expenses)
- Category-wise breakdown with formulas
- Currency formatting ($)
- Date formatting (MM/DD/YYYY)

### 🎨 Professional Formatting
- Color-coded headers (blue with white text)
- Frozen header rows for easy scrolling
- Borders and proper alignment
- Properly sized columns

### 📝 Sample Data Included
- Pre-filled with realistic expense examples
- Shows how to structure your entries
- Easy to delete and replace

---

## How to Use

### Adding a New Expense
1. Open `Expense_Tracker.xlsx`
2. Go to "Expenses" sheet
3. Find the next empty row
4. Enter:
   - **Date**: When you spent the money (01/20/2024)
   - **Category**: Type from the Categories sheet
   - **Description**: What you bought
   - **Amount**: How much you spent (123.45)
   - **Payment Method**: Credit Card, Cash, Bank Transfer, etc.
   - **Tags**: Keywords like "food", "work", etc. (optional)
   - **Notes**: Any additional details (optional)

### Viewing Your Summary
1. Click on "Summary" sheet
2. See your **Total Expenses** automatically calculated
3. See spending breakdown by category

### Example Entry
```
Date: 01/20/2024
Category: Groceries
Description: Weekly grocery shopping
Amount: 87.50
Payment Method: Credit Card
Tags: food, weekly
Notes: Whole Foods
```

---

## Platform-Specific Notes

### Windows
- **File Opens With**: Microsoft Excel (default)
- **Alternative Apps**: Google Sheets, LibreOffice, OnlyOffice
- **Terminal**: Use `cmd` or PowerShell
- **Command**: `python create_tracker.py`

### Mac
- **File Opens With**: Numbers or Microsoft Excel
- **Alternative Apps**: Google Sheets, LibreOffice, OnlyOffice
- **Terminal**: Use Terminal app (Applications > Utilities > Terminal)
- **Command**: `python3 create_tracker.py`
- **Tip**: If you get "command not found", install from python.org

### Linux
- **File Opens With**: LibreOffice Calc (or any spreadsheet app)
- **Alternative Apps**: Google Sheets, OnlyOffice
- **Terminal**: Use your default terminal
- **Command**: `python3 create_tracker.py`

---

## Customization

### Add New Categories
1. Go to "Categories" sheet
2. Add your custom categories in column A
3. Use them in the Expenses sheet

### Modify Payment Methods
- Just type any payment method you use in the "Payment Method" column
- Excel doesn't restrict values - it's fully flexible

### Track Different Date Ranges
- Modify the Summary formulas to filter by date
- Example: `=SUMIF(Expenses!B:B,"Groceries",Expenses!D:D)` filters by category

### Add More Columns
- Right-click on any column and insert new ones
- Add custom fields as needed (Income, Location, Card Used, etc.)

---

## Troubleshooting

### Windows: "'python' is not recognized"
```bash
# Try this instead:
pip install openpyxl
python create_tracker.py
```
If still fails, reinstall Python and check "Add Python to PATH"

### Mac: "command not found: python3"
```bash
# Install Python from python.org or use Homebrew:
brew install python3

# Then try:
pip3 install openpyxl
python3 create_tracker.py
```

### "pip install failed"
- Windows: Use `pip install openpyxl`
- Mac/Linux: Use `pip3 install openpyxl`

### Excel file won't open
- Try opening with Google Sheets (free, online)
- Or LibreOffice (free, open-source)
- Or OnlyOffice (free, online)

### Formulas not calculating
- Make sure Automatic Calculation is enabled
- **Excel**: Formulas → Calculation Options → Automatic
- **Google Sheets**: Should auto-calculate
- **LibreOffice**: Tools → Options → LibreOffice Calc → Calculate

### Can't find the file
- Run the script from a visible location (Desktop or Documents)
- Check your Downloads folder
- Open your file manager and search for "Expense_Tracker.xlsx"

---

## What's Next?

### Basic Users
✓ Done! Start tracking expenses with the Excel file

### Advanced Users
- Modify formulas to create custom reports
- Add pivot tables for deeper analysis
- Create charts for visualization
- Export to CSV for backups

### Developers
- Web app coming soon (React + Node.js)
- Database support for cloud sync
- Mobile app support

---

## Tips & Tricks

📱 **Use Everywhere**
- Excel Desktop (Windows/Mac)
- Google Sheets (all devices)
- OnDrive/Dropbox for cloud backup
- Tablet or phone via cloud apps

📊 **Better Tracking**
- Enter expenses daily for accuracy
- Use consistent category names
- Add descriptive notes for reference
- Review Summary sheet weekly

💾 **Backup Your File**
- Keep a copy in cloud storage (Google Drive, OneDrive, etc.)
- Save backup copies monthly
- Export to CSV for long-term storage

---

## Support & Feedback

- **Found a bug?** Create an issue on GitHub
- **Need a feature?** Submit a request
- **Want to contribute?** We welcome pull requests!

---

**You're all set! Start tracking expenses now.** 🎉
