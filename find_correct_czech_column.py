#!/usr/bin/env python3
"""
Find the exact correct column for Czech_and_Poland that gives 58.41.
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd

def find_correct_czech_column():
    print("🔍 FINDING CORRECT Czech_and_Poland Column")
    print("=" * 60)
    
    # Load MultiTicker
    multiticker_df = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    # Get criteria headers
    criteria_row2 = multiticker_df.iloc[14, 2:400].values  # Subcategory row
    
    # Get first data row
    data_row = multiticker_df.iloc[19, 2:400].values
    
    print("🔍 All Czech_and_Poland matches with their values:")
    print("   Excel Col | Index | Value     | Difference from 58.41")
    print("   " + "-" * 55)
    
    target_value = 58.41
    best_match = None
    best_diff = float('inf')
    
    # Check all Czech_and_Poland columns
    for i in range(len(criteria_row2)):
        if i >= len(data_row):
            continue
            
        subcategory = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
        
        if subcategory == "Czech and Poland":
            value = data_row[i] if pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
            value = float(value)
            
            # Calculate Excel column name
            col_idx = i + 2  # Add 2 for offset (A=0, B=1, C=2...)
            if col_idx < 26:
                excel_col = chr(65 + col_idx)
            else:
                excel_col = chr(65 + (col_idx // 26) - 1) + chr(65 + (col_idx % 26))
            
            diff = abs(value - target_value)
            
            print(f"   {excel_col:9} | {i:5d} | {value:8.2f} | {diff:6.2f}")
            
            if diff < best_diff:
                best_diff = diff
                best_match = {
                    'excel_col': excel_col,
                    'index': i,
                    'value': value,
                    'diff': diff
                }
    
    print("   " + "-" * 55)
    
    if best_match:
        print(f"✅ BEST MATCH: Column {best_match['excel_col']} = {best_match['value']:.2f}")
        print(f"   Difference from target: {best_match['diff']:.2f}")
        
        if best_match['diff'] < 0.1:
            print(f"   🎯 EXACT MATCH FOUND!")
        else:
            print(f"   ⚠️  No exact match - closest is {best_match['diff']:.2f} away")
            
        # Now check what the Daily sheet column AB actually maps to
        print(f"\n🔍 Checking Daily sheet column AB mapping:")
        daily_df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
        
        # AB = column index 27 (A=0, B=1, ..., AB=27)
        ab_daily_value = daily_df.iloc[12, 27]  # Row 13 in Daily sheet (data row)
        print(f"   Daily sheet AB row 13 value: {ab_daily_value}")
        
        return best_match
    else:
        print("❌ No Czech_and_Poland matches found")
        return None

if __name__ == "__main__":
    result = find_correct_czech_column()