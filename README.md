# Expense Tracker

A fully functional, easy-to-use expense tracking system. Works on Windows, Mac, and Linux.

## 🚀 Quick Start

### 30-Second Setup
```bash
# Windows
pip install openpyxl && python create_tracker.py

# Mac/Linux
pip3 install openpyxl && python3 create_tracker.py
```

Then open `Expense_Tracker.xlsx` and start tracking!

## 📖 Documentation

- **[QUICK_START.md](QUICK_START.md)** - User guide with tips and tricks
- **[INSTALL.md](INSTALL.md)** - Detailed installation for Windows, Mac, Linux
- **[PROJECT.md](PROJECT.md)** - Project overview and features

## ✨ Features

✅ **Works Offline** - No internet needed  
✅ **No Installation Required** - Just Python (pre-installed on most systems)  
✅ **Accessible Format** - Works with Excel, Google Sheets, LibreOffice  
✅ **Professional Design** - Color-coded, formatted, easy to read  
✅ **Auto-Calculating** - Totals update automatically  
✅ **Fully Customizable** - Add categories, modify columns, extend as needed  
✅ **Zero Learning Curve** - If you've used a spreadsheet, you can use this  
✅ **Cross-Platform** - Windows, Mac, and Linux  

## 📊 What You Get

### Expenses Sheet
- Track date, category, description, amount, payment method, tags, notes
- Sample data included for reference
- Add/edit/delete rows as needed

### Summary Sheet
- Total expenses (auto-calculated)
- Breakdown by category (auto-calculated)
- Always shows current totals

### Categories Sheet
- Pre-defined categories (Groceries, Utilities, Transportation, etc.)
- Easy to customize
- Reference list for consistency

## 🎯 Use Cases

- Track personal monthly expenses
- Monitor budget spending
- Categorize expenses for analysis
- Generate expense reports
- Export data for tax purposes
- Share with accountants or family
- Create financial records

## 💻 System Requirements

- **Python 3.7+** (usually pre-installed)
- **pip** (Python package manager)
- **Any spreadsheet app** (Excel, Google Sheets, LibreOffice, etc.)

## 🔧 Installation

### Windows
```bash
pip install openpyxl
python create_tracker.py
```

### Mac
```bash
pip3 install openpyxl
python3 create_tracker.py
```

### Linux
```bash
pip3 install openpyxl
python3 create_tracker.py
```

For detailed instructions, see [INSTALL.md](INSTALL.md)

## 📱 Compatibility

| Platform | Python | Excel | Google Sheets | LibreOffice | Notes |
|----------|--------|-------|----------------|-------------|-------|
| **Windows** | ✅ | ✅ | ✅ | ✅ | Fully supported |
| **Mac** | ✅ | ✅ | ✅ | ✅ | Fully supported |
| **Linux** | ✅ | - | ✅ | ✅ | Fully supported |
| **Mobile** | - | ✅ | ✅ | ⚠️ | Via cloud (Google Sheets recommended) |

## 📖 How to Use

### Adding an Expense

1. Open `Expense_Tracker.xlsx`
2. Go to the "Expenses" sheet
3. Find the next empty row
4. Fill in:
   - **Date**: 01/20/2024
   - **Category**: Groceries (or your category)
   - **Description**: Weekly grocery shopping
   - **Amount**: 87.50
   - **Payment Method**: Credit Card
   - **Tags**: food, weekly (optional)
   - **Notes**: Whole Foods (optional)
5. Press Enter
6. The Summary sheet updates automatically!

### Viewing Summary

1. Click on "Summary" sheet
2. See **Total Expenses** (auto-calculated)
3. See spending by category (auto-calculated)

## 🎨 Features Explained

### Auto-Calculating Formulas
- **Total Expenses**: `=SUM(Expenses!D2:D1000)` - adds all amounts
- **Category Totals**: `=SUMIF(Expenses!B:B,"Category",Expenses!D:D)` - sums by category
- Updates automatically as you add expenses

### Professional Formatting
- Blue headers with white text
- Currency formatting ($)
- Date formatting (MM/DD/YYYY)
- Frozen header rows
- Proper column widths
- Cell borders for clarity

### Easy Customization
- Add new categories in Categories sheet
- Add new columns as needed
- Modify formulas for custom calculations
- Change colors and formatting
- Export to CSV/PDF

## 🔒 Data & Privacy

- Your data stays on your device
- No cloud upload (unless you choose to)
- No tracking or analytics
- Fully open-source and transparent
- Use with confidence

## 🚀 Getting Started

1. **Install Python** (if needed)
   - Check: `python3 --version`
   - Download from [python.org](https://www.python.org/downloads/)

2. **Install openpyxl**
   ```bash
   pip3 install openpyxl  # Mac/Linux
   pip install openpyxl   # Windows
   ```

3. **Generate your tracker**
   ```bash
   python3 create_tracker.py  # Mac/Linux
   python create_tracker.py   # Windows
   ```

4. **Open and use**
   - Find `Expense_Tracker.xlsx`
   - Open with Excel, Google Sheets, or LibreOffice
   - Start tracking!

## 📚 Documentation

- **README.md** (this file) - Overview and features
- **QUICK_START.md** - User guide with examples and tips
- **INSTALL.md** - Detailed platform-specific installation
- **PROJECT.md** - Project roadmap and future plans

## 🔮 Future Plans

- Web dashboard (React + Node.js)
- Data visualization (charts, graphs)
- Mobile app support
- Cloud sync and backup
- Budget alerts and notifications
- Receipt scanning (OCR)
- Multi-user support
- CSV/PDF export automation

## 🤝 Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Commit (`git commit -m 'Add amazing feature'`)
5. Push (`git push origin feature/amazing-feature`)
6. Open a Pull Request

## 📝 License

MIT License - see LICENSE file for details

## ❓ FAQ

**Q: Do I need to pay for this?**  
A: No, it's completely free and open-source.

**Q: Can I use this on my phone?**  
A: Yes! Upload to Google Sheets for access on any device.

**Q: Will my data be stored in the cloud?**  
A: Only if you upload to Google Sheets or another service. Local files stay on your device.

**Q: Can I add more columns?**  
A: Absolutely! It's a spreadsheet - fully customizable.

**Q: Is my data secure?**  
A: Local files are as secure as your computer. Use caution with cloud services.

**Q: How often should I back up?**  
A: Daily or weekly depending on usage. Keep copies in cloud storage.

**Q: What if I have multiple years of data?**  
A: Create a new sheet for each year or use the cloud version for unlimited rows.

**Q: Can I share with others?**  
A: Yes! Upload to Google Sheets and share the link.

**Q: What if formulas break?**  
A: The file includes explanations. Recreate by running the script again.

## 🐛 Reporting Issues

Found a bug? Please create an issue with:
- Your operating system
- Python version (`python --version`)
- What you were trying to do
- Error message (if any)
- Steps to reproduce

## 💬 Questions?

- Check the documentation first
- Create an issue with your question
- Look for existing solutions

## ⭐ Show Your Support

If this helps you, please consider:
- ⭐ Starring the repository
- 🔄 Sharing with friends
- 📢 Recommending to others
- 🐛 Reporting bugs
- 💡 Suggesting features

---

## Quick Links

- **[Quick Start](QUICK_START.md)** - Get started in 60 seconds
- **[Installation Guide](INSTALL.md)** - Detailed setup instructions
- **[Project Overview](PROJECT.md)** - Full project details
- **[GitHub Issues](../../issues)** - Report problems or suggest features

---

**Ready to start?** Run `python3 create_tracker.py` now! 🎉
