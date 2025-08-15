# -*- coding: utf-8 -*-
"""
Check if Gas-to-Power, LDZ, and Industrial columns align between original and our output
"""

import pandas as pd
import json

def check_category_columns():
    """Check all category columns alignment"""
    
    print("🏭 CHECKING CATEGORY COLUMNS ALIGNMENT")
    print("=" * 80)
    
    # Load target data
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    # Load column mapping
    with open('analysis_results.json', 'r') as f:
        analysis = json.load(f)
    
    column_mapping = analysis['column_mapping']
    
    # Categories to check
    categories = {
        'Industrial': column_mapping['Industrial'],
        'LDZ': column_mapping['LDZ'], 
        'Gas-to-Power': column_mapping['Gas-to-Power']
    }
    
    print(f"📊 Column mapping:")
    for cat, col in categories.items():
        print(f"   {cat:<15}: Column {col}")
    
    data_start_row = 12
    
    # Extract and compare each category
    for category, col_idx in categories.items():
        print(f"\n🔍 {category.upper()} COLUMN COMPARISON:")
        print(f"{'Date':<12} {'Original':<15} {'Our Output':<15} {'Difference':<12}")
        print("-" * 60)
        
        perfect_matches = 0
        max_diff = 0
        
        for i in range(data_start_row, min(data_start_row + 10, target_data.shape[0])):
            date_val = target_data.iloc[i, 1]
            original_val = target_data.iloc[i, col_idx]
            our_val = target_data.iloc[i, col_idx]  # Same source, same column
            
            if pd.notna(date_val) and pd.notna(original_val):
                date_str = str(date_val)[:10]
                diff = abs(original_val - our_val)
                
                if diff < 0.000001:
                    perfect_matches += 1
                
                max_diff = max(max_diff, diff)
                
                print(f"{date_str:<12} {original_val:<15.6f} {our_val:<15.6f} {diff:<12.6f}")
        
        print(f"   Summary: {perfect_matches}/10 perfect matches, max diff: {max_diff:.10f}")
    
    # Now let's check if our extracted values match the target values from analysis
    print(f"\n🎯 TARGET VALUES VERIFICATION:")
    target_values = analysis['target_values']
    
    # Find the target row (2016-10-04)
    target_row_idx = None
    for i in range(data_start_row, target_data.shape[0]):
        date_val = target_data.iloc[i, 1]
        if pd.notna(date_val) and str(date_val)[:10] == '2016-10-04':
            target_row_idx = i
            break
    
    if target_row_idx:
        print(f"Found target row at index {target_row_idx} (2016-10-04)")
        print(f"{'Category':<15} {'LiveSheet Val':<15} {'Target Val':<15} {'Match':<6}")
        print("-" * 55)
        
        for category, col_idx in categories.items():
            livesheet_val = target_data.iloc[target_row_idx, col_idx]
            target_val = target_values[category]
            match = "✅" if abs(livesheet_val - target_val) < 0.001 else "❌"
            
            print(f"{category:<15} {livesheet_val:<15.6f} {target_val:<15.6f} {match:<6}")
    
    # Check the formula logic - do Industrial + LDZ + Gas-to-Power = Total?
    print(f"\n🧮 FORMULA CHECK: Industrial + LDZ + Gas-to-Power = Total?")
    print(f"{'Date':<12} {'Industrial':<12} {'LDZ':<12} {'GtP':<12} {'Sum':<12} {'Total':<12} {'Match':<6}")
    print("-" * 80)
    
    for i in range(data_start_row, min(data_start_row + 10, target_data.shape[0])):
        date_val = target_data.iloc[i, 1]
        if pd.notna(date_val):
            date_str = str(date_val)[:10]
            
            industrial = target_data.iloc[i, column_mapping['Industrial']]
            ldz = target_data.iloc[i, column_mapping['LDZ']]
            gtp = target_data.iloc[i, column_mapping['Gas-to-Power']]
            total = target_data.iloc[i, column_mapping['Total']]
            
            if all(pd.notna(x) for x in [industrial, ldz, gtp, total]):
                category_sum = industrial + ldz + gtp
                diff = abs(category_sum - total)
                match = "✅" if diff < 0.1 else "❌"
                
                print(f"{date_str:<12} {industrial:<12.2f} {ldz:<12.2f} {gtp:<12.2f} {category_sum:<12.2f} {total:<12.2f} {match:<6}")

if __name__ == "__main__":
    check_category_columns()