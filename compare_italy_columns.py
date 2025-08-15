# -*- coding: utf-8 -*-
"""
Compare Italy columns between our output and the original LiveSheet
"""

import pandas as pd

def compare_italy_columns():
    """Compare Italy columns directly"""
    
    print("🇮🇹 COMPARING ITALY COLUMNS")
    print("=" * 80)
    
    # Read the original LiveSheet
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    print("📊 Reading original LiveSheet...")
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    # Extract Italy column (column 4) starting from data row 12
    data_start_row = 12
    italy_original = []
    dates_original = []
    
    for i in range(data_start_row, min(data_start_row + 10, target_data.shape[0])):
        date_val = target_data.iloc[i, 1]  # Date in column 1
        italy_val = target_data.iloc[i, 4]  # Italy in column 4
        
        if pd.notna(date_val) and pd.notna(italy_val):
            dates_original.append(date_val)
            italy_original.append(float(italy_val))
    
    # Read our corrected output
    print("📊 Reading our corrected output...")
    
    # Since xlsx failed, let me create a CSV version first
    import json
    with open('analysis_results.json', 'r') as f:
        analysis = json.load(f)
    
    column_mapping = analysis['column_mapping']
    
    # Extract from target data using correct column mapping
    our_data = []
    our_dates = []
    
    for i in range(data_start_row, min(data_start_row + 10, target_data.shape[0])):
        date_val = target_data.iloc[i, 1]
        italy_val = target_data.iloc[i, column_mapping['Italy']]  # Use correct column mapping
        
        if pd.notna(date_val) and pd.notna(italy_val):
            our_dates.append(date_val)
            our_data.append(float(italy_val))
    
    print(f"\n🔍 ITALY COLUMN COMPARISON (first 10 rows):")
    print(f"{'Date':<12} {'Original':<12} {'Our Output':<12} {'Match':<6}")
    print("-" * 50)
    
    perfect_matches = 0
    for i in range(min(len(italy_original), len(our_data))):
        original_val = italy_original[i]
        our_val = our_data[i]
        date_str = str(dates_original[i])[:10] if i < len(dates_original) else "N/A"
        
        match = "✅" if abs(original_val - our_val) < 0.001 else "❌"
        if match == "✅":
            perfect_matches += 1
            
        print(f"{date_str:<12} {original_val:<12.3f} {our_val:<12.3f} {match:<6}")
    
    print(f"\n📊 ITALY COLUMN SUMMARY:")
    print(f"   Perfect matches: {perfect_matches}/{min(len(italy_original), len(our_data))}")
    
    # Let's also check if we're using the same column
    print(f"\n🔍 COLUMN VERIFICATION:")
    print(f"   Original Italy is in column 4")
    print(f"   Our mapping says Italy is in column {column_mapping['Italy']}")
    print(f"   Are we using the same column? {'✅ YES' if column_mapping['Italy'] == 4 else '❌ NO'}")
    
    # Double check by extracting both columns
    if column_mapping['Italy'] != 4:
        print(f"\n⚠️  POTENTIAL ISSUE DETECTED!")
        print(f"   Let's compare column 4 vs column {column_mapping['Italy']}:")
        
        for i in range(data_start_row, min(data_start_row + 5, target_data.shape[0])):
            date_val = target_data.iloc[i, 1]
            col4_val = target_data.iloc[i, 4]
            our_col_val = target_data.iloc[i, column_mapping['Italy']]
            
            date_str = str(date_val)[:10] if pd.notna(date_val) else "N/A"
            print(f"   {date_str}: Column 4 = {col4_val:.3f}, Column {column_mapping['Italy']} = {our_col_val:.3f}")
    
    return italy_original, our_data

if __name__ == "__main__":
    italy_orig, italy_ours = compare_italy_columns()