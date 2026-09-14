#!/usr/bin/env python3
"""
Expense Tracker - Simple Excel-based solution
Creates an accessible Excel file for tracking expenses
Works on Windows, Mac, and Linux
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import sys

def create_expense_tracker():
    """Create a fully functional Excel expense tracker"""
    
    try:
        # Create workbook
        wb = openpyxl.Workbook()
        
        # Remove default sheet
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        
        # Create sheets
        expenses_sheet = wb.create_sheet('Expenses', 0)
        summary_sheet = wb.create_sheet('Summary', 1)
        categories_sheet = wb.create_sheet('Categories', 2)
        
        # Define styles
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)
        summary_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        summary_font = Font(bold=True, size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        currency_format = '$#,##0.00'
        date_format = 'mm/dd/yyyy'
        
        # ===== EXPENSES SHEET =====
        expenses_headers = ['Date', 'Category', 'Description', 'Amount', 'Payment Method', 'Tags', 'Notes']
        
        for col, header in enumerate(expenses_headers, 1):
            cell = expenses_sheet.cell(row=1, column=col)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        # Set column widths
        expenses_sheet.column_dimensions['A'].width = 12
        expenses_sheet.column_dimensions['B'].width = 15
        expenses_sheet.column_dimensions['C'].width = 25
        expenses_sheet.column_dimensions['D'].width = 12
        expenses_sheet.column_dimensions['E'].width = 15
        expenses_sheet.column_dimensions['F'].width = 15
        expenses_sheet.column_dimensions['G'].width = 20
        
        # Add sample data
        sample_data = [
            ['1/15/2024', 'Groceries', 'Weekly grocery shopping', 85.50, 'Credit Card', 'food', ''],
            ['1/16/2024', 'Utilities', 'Electric bill', 120.00, 'Bank Transfer', 'bills', ''],
            ['1/17/2024', 'Entertainment', 'Movie tickets', 30.00, 'Cash', 'entertainment', ''],
            ['1/18/2024', 'Transportation', 'Gas', 45.75, 'Credit Card', 'transport', ''],
            ['1/19/2024', 'Dining', 'Lunch with friends', 25.00, 'Credit Card', 'food', ''],
        ]
        
        for row_idx, data in enumerate(sample_data, 2):
            for col_idx, value in enumerate(data, 1):
                cell = expenses_sheet.cell(row=row_idx, column=col_idx)
                cell.value = value
                cell.border = border
                cell.alignment = Alignment(horizontal='left', vertical='center')
                
                # Format date and amount
                if col_idx == 1:  # Date column
                    cell.number_format = date_format
                    cell.alignment = Alignment(horizontal='center')
                elif col_idx == 4:  # Amount column
                    cell.number_format = currency_format
                    cell.alignment = Alignment(horizontal='right')
        
        # Freeze header row
        expenses_sheet.freeze_panes = 'A2'
        
        # ===== SUMMARY SHEET =====
        summary_sheet['A1'] = 'EXPENSE SUMMARY'
        summary_sheet['A1'].font = Font(bold=True, size=14)
        summary_sheet.merge_cells('A1:B1')
        
        summary_sheet['A3'] = 'Total Expenses:'
        summary_sheet['A3'].font = summary_font
        summary_sheet['B3'] = '=SUM(Expenses!D2:D1000)'
        summary_sheet['B3'].number_format = currency_format
        summary_sheet['B3'].font = Font(bold=True, size=11)
        
        summary_sheet['A5'] = 'By Category'
        summary_sheet['A5'].font = Font(bold=True, size=12)
        
        # Category breakdown headers
        summary_sheet['A6'] = 'Category'
        summary_sheet['B6'] = 'Total'
        for cell in ['A6', 'B6']:
            summary_sheet[cell].fill = summary_fill
            summary_sheet[cell].font = summary_font
            summary_sheet[cell].border = border
        
        categories = ['Groceries', 'Utilities', 'Entertainment', 'Transportation', 'Dining']
        for idx, category in enumerate(categories, 7):
            summary_sheet[f'A{idx}'] = category
            summary_sheet[f'B{idx}'] = f'=SUMIF(Expenses!B:B,A{idx},Expenses!D:D)'
            summary_sheet[f'B{idx}'].number_format = currency_format
            summary_sheet[f'B{idx}'].border = border
            summary_sheet[f'A{idx}'].border = border
        
        summary_sheet.column_dimensions['A'].width = 20
        summary_sheet.column_dimensions['B'].width = 15
        
        # ===== CATEGORIES SHEET =====
        categories_sheet['A1'] = 'Available Categories'
        categories_sheet['A1'].font = Font(bold=True, size=12)
        
        category_list = [
            'Groceries',
            'Utilities',
            'Transportation',
            'Entertainment',
            'Dining',
            'Healthcare',
            'Shopping',
            'Subscriptions',
            'Education',
            'Other'
        ]
        
        for idx, cat in enumerate(category_list, 2):
            cell = categories_sheet.cell(row=idx, column=1)
            cell.value = cat
            cell.border = border
        
        categories_sheet.column_dimensions['A'].width = 20
        
        # Save file
        filename = 'Expense_Tracker.xlsx'
        wb.save(filename)
        
        # Success message
        print("\n" + "="*60)
        print("✓ SUCCESS! Expense Tracker created")
        print("="*60)
        print(f"\nFile: {filename}")
        print(f"Location: {os.path.abspath(filename)}")
        print("\nFeatures:")
        print("  ✓ Expenses sheet with sample data")
        print("  ✓ Auto-calculating Summary sheet")
        print("  ✓ Category breakdown with formulas")
        print("  ✓ Professional formatting and styling")
        print("  ✓ Frozen header rows for easy navigation")
        print("  ✓ Pre-formatted currency and date columns")
        print("\nNext Steps:")
        print("  1. Find Expense_Tracker.xlsx in your folder")
        print("  2. Open with Excel, Google Sheets, or LibreOffice")
        print("  3. Start tracking your expenses!")
        print("\nSee QUICK_START.md for detailed usage instructions")
        print("="*60 + "\n")
        return True
        
    except ImportError:
        print("\n" + "="*60)
        print("ERROR: openpyxl not installed")
        print("="*60)
        print("\nTo fix, run one of these commands:\n")
        print("  Windows: pip install openpyxl")
        print("  Mac:     pip3 install openpyxl")
        print("  Linux:   pip3 install openpyxl")
        print("\nThen run this script again.")
        print("="*60 + "\n")
        return False
    except Exception as e:
        print("\n" + "="*60)
        print(f"ERROR: {str(e)}")
        print("="*60)
        print("\nPlease check:")
        print("  1. Python 3.7+ is installed")
        print("  2. openpyxl is installed (pip install openpyxl)")
        print("  3. You have write permissions in this directory")
        print("="*60 + "\n")
        return False

if __name__ == '__main__':
    success = create_expense_tracker()
    sys.exit(0 if success else 1)