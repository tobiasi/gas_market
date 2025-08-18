#!/usr/bin/env python3
"""
Check actual data values in supply columns for 2016-10-01.
"""

import pandas as pd
import sys
sys.path.append("C:/development/commodities")

def check_actual_values():
    print("🔍 CHECKING ACTUAL DATA VALUES FOR 2016-10-01")
    print("=" * 80)
    
    # Load Excel file
    df_full = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    data_start = 19 - 1  # Row 19 (0-based = 18)
    
    # Get the first data row (2016-10-01)
    data_row_idx = data_start + 0  # First row of data
    
    print(f"📊 Checking data row {data_row_idx + 1} (should be 2016-10-01)")
    
    # Check key supply columns
    supply_columns = [
        {'name': 'MAB_to_Austria', 'indices': [3, 4, 5], 'cols': ['D', 'E', 'F']},
        {'name': 'Austria_Production', 'indices': [7, 8], 'cols': ['H', 'I']},
        {'name': 'Germany_Production', 'indices': [22, 23, 24, 25, 26, 27, 28], 'cols': ['W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC']},
        {'name': 'Norway_to_Europe', 'indices': [33], 'cols': ['AH']},  # Just check one Norway column first
        {'name': 'LNG_to_Belgium', 'indices': [31], 'cols': ['AF']},
        {'name': 'Slovenia_to_Austria', 'indices': [34], 'cols': ['AI']},  # Actually AI based on our finding
    ]
    
    print("   Route                    Excel   Index   Raw Value    Numeric    Sum")
    print("   " + "-" * 70)
    
    for supply_col in supply_columns:
        print(f"\n   {supply_col['name']}:")
        total_sum = 0
        
        for col_name, col_idx in zip(supply_col['cols'], supply_col['indices']):
            if col_idx < len(df_full.columns):
                raw_value = df_full.iloc[data_row_idx, col_idx]
                
                try:
                    numeric_value = pd.to_numeric(raw_value, errors='coerce')
                    if pd.notna(numeric_value):
                        total_sum += numeric_value
                        print(f"      {col_name:8} {col_idx:6d} {str(raw_value):>12} {numeric_value:>8.2f}")
                    else:
                        print(f"      {col_name:8} {col_idx:6d} {str(raw_value):>12} {'NaN':>8}")
                except Exception as e:
                    print(f"      {col_name:8} {col_idx:6d} {str(raw_value):>12} ERROR")
            else:
                print(f"      {col_name:8} {col_idx:6d} {'OUT_OF_BOUNDS':>25}")
        
        print(f"      TOTAL: {total_sum:.2f}")
    
    # Now also check a broader scan for ANY non-zero values in the first row
    print(f"\n🔍 SCANNING FOR ANY NON-ZERO VALUES IN FIRST DATA ROW:")
    non_zero_found = []
    
    for col_idx in range(2, min(50, len(df_full.columns))):  # Scan first 50 columns
        raw_value = df_full.iloc[data_row_idx, col_idx]
        try:
            numeric_value = pd.to_numeric(raw_value, errors='coerce')
            if pd.notna(numeric_value) and numeric_value != 0:
                excel_col = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
                non_zero_found.append((excel_col, col_idx, numeric_value))
        except:
            pass
    
    if non_zero_found:
        print(f"   ✅ Found {len(non_zero_found)} non-zero values:")
        for excel_col, col_idx, value in non_zero_found[:10]:  # Show first 10
            print(f"      {excel_col}: {value:.2f}")
    else:
        print(f"   ❌ No non-zero values found in first 50 columns")

if __name__ == "__main__":
    check_actual_values()