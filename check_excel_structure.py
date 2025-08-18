#!/usr/bin/env python3
"""
Check the exact Excel structure to understand the header layout.
"""

import openpyxl
import pandas as pd

def check_excel_structure():
    wb = openpyxl.load_workbook('use4.xlsx', read_only=True, data_only=True)
    ws = wb['MultiTicker']
    
    print("🔍 CHECKING EXCEL STRUCTURE")
    print("=" * 80)
    
    # Check first 20 rows across first 10 columns
    for row_num in range(1, 21):
        print(f"\nRow {row_num:2d}:", end="")
        for col in range(1, 11):  # A to J
            cell = ws.cell(row=row_num, column=col)
            value = str(cell.value)[:15] if cell.value else "None"
            print(f" {value:<15}", end="")
    
    wb.close()

if __name__ == "__main__":
    check_excel_structure()