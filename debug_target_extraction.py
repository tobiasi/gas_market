# -*- coding: utf-8 -*-
"""
Debug target data extraction to understand why values are coming back as 0
"""

import pandas as pd
import numpy as np

def debug_target_extraction():
    """Debug the target data extraction"""
    
    print("🔍 DEBUGGING TARGET DATA EXTRACTION")
    print("=" * 80)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    # Read the target data
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    print(f"📊 Target data shape: {target_data.shape}")
    
    # Look at the header row (row 10)
    print(f"\n📋 HEADER ROW (row 10):")
    header_row = 10
    for j in range(min(20, target_data.shape[1])):
        val = target_data.iloc[header_row, j]
        print(f"   Column {j:2d}: {val}")
    
    # Look at data rows (starting from row 11)
    print(f"\n📊 DATA ROWS (starting from row 11):")
    data_start_row = 11
    
    for i in range(data_start_row, min(data_start_row + 10, target_data.shape[0])):
        print(f"\nRow {i}:")
        row_values = []
        for j in range(min(20, target_data.shape[1])):
            val = target_data.iloc[i, j]
            if pd.notna(val):
                if isinstance(val, (int, float)):
                    row_values.append(f"Col{j}={val:.2f}")
                else:
                    row_values.append(f"Col{j}={str(val)[:10]}")
        
        if row_values:
            print(f"   {', '.join(row_values[:10])}")
            if len(row_values) > 10:
                print(f"   ... and {len(row_values) - 10} more values")
        else:
            print("   All NaN or empty")
    
    # Specifically check the target columns we found in our previous analysis
    key_columns = [2, 3, 4, 8, 12, 13, 14, 15]  # France, Belgium, Italy, Germany, Total, Industrial, LDZ, Gas-to-Power
    
    print(f"\n🎯 CHECKING KEY COLUMNS:")
    print(f"Looking for data in columns: {key_columns}")
    
    # Find the first row with actual data
    found_data_row = None
    for i in range(data_start_row, min(data_start_row + 50, target_data.shape[0])):
        has_data = False
        for col in key_columns:
            val = target_data.iloc[i, col]
            if pd.notna(val) and isinstance(val, (int, float)) and abs(val) > 1:
                has_data = True
                break
        
        if has_data:
            found_data_row = i
            break
    
    if found_data_row is not None:
        print(f"\n✅ FOUND DATA AT ROW {found_data_row}:")
        for col in key_columns:
            val = target_data.iloc[found_data_row, col]
            header = target_data.iloc[header_row, col]
            print(f"   Column {col:2d} ({header}): {val}")
        
        # Check if this matches our target values
        italy_val = target_data.iloc[found_data_row, 4]  # Italy is column 4
        total_val = target_data.iloc[found_data_row, 12] # Total is column 12
        
        print(f"\n🎯 VALUES CHECK:")
        print(f"   Italy (Col 4): {italy_val} (target: ~151.47)")
        print(f"   Total (Col 12): {total_val} (target: ~767.69)")
        
        # Return the correct row index to use
        return found_data_row
        
    else:
        print(f"\n❌ NO DATA FOUND in first 50 rows")
        
        # Let's check if data exists later in the file
        print(f"\n🔍 SCANNING ENTIRE FILE FOR DATA:")
        for i in range(0, target_data.shape[0], 100):  # Check every 100th row
            for col in key_columns:
                val = target_data.iloc[i, col]
                if pd.notna(val) and isinstance(val, (int, float)) and abs(val) > 10:
                    print(f"   Found data at row {i}, col {col}: {val}")
                    return i
        
        return None

if __name__ == "__main__":
    data_row = debug_target_extraction()