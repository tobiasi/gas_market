# -*- coding: utf-8 -*-
"""
Check the actual headers for columns 32, 36, 38 in the original sheet
"""

import pandas as pd

def check_empty_headers():
    """Check what the original headers actually are"""
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    print("🔍 CHECKING EMPTY HEADERS")
    print("=" * 50)
    
    problem_cols = [32, 36, 38]
    
    for col in problem_cols:
        print(f"\nColumn {col}:")
        print(f"  Row 9:  '{target_data.iloc[9, col]}'")
        print(f"  Row 10: '{target_data.iloc[10, col]}'")
        print(f"  Row 11: '{target_data.iloc[11, col]}'")
        
        # What we should name it
        row9 = str(target_data.iloc[9, col]) if pd.notna(target_data.iloc[9, col]) else ""
        row10 = str(target_data.iloc[10, col]) if pd.notna(target_data.iloc[10, col]) else ""
        
        row9_clean = row9.replace('nan', '').strip()
        row10_clean = row10.replace('nan', '').strip()
        
        if row10_clean:
            correct_name = row10_clean
        elif row9_clean:
            correct_name = row9_clean
        else:
            correct_name = ""
        
        print(f"  Correct header: '{correct_name}'")
        
        # Check for duplicate handling
        if col == 32 and correct_name == "Austria":
            print(f"  Note: This is a duplicate 'Austria' column (col 28 is Export Austria, col 32 is Production Austria)")

if __name__ == "__main__":
    check_empty_headers()