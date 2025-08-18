# -*- coding: utf-8 -*-
"""
Double-check that our supply columns exactly match the live balance sheet
"""

import pandas as pd

def verify_supply_columns_match():
    """Verify our supply extraction matches the original exactly"""
    
    print("🔍 VERIFYING SUPPLY COLUMNS MATCH LIVE BALANCE SHEET")
    print("=" * 80)
    
    # Read original LiveSheet
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    original_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    # Read our master output
    master_file = 'European_Gas_Market_Master.xlsx'
    our_supply = pd.read_excel(master_file, sheet_name='Supply')
    
    print(f"📊 Original sheet shape: {original_data.shape}")
    print(f"📊 Our supply tab shape: {our_supply.shape}")
    
    # Check columns 17-38 mapping
    supply_start_col = 17
    supply_end_col = 38
    
    print(f"\n📋 COLUMN-BY-COLUMN COMPARISON (Columns {supply_start_col}-{supply_end_col}):")
    print(f"{'Col':<4} {'Original Header':<25} {'Our Header':<25} {'Match':<6}")
    print("-" * 70)
    
    # Build mapping from original
    original_headers = {}
    for j in range(supply_start_col, supply_end_col + 1):
        row10 = str(original_data.iloc[10, j]) if pd.notna(original_data.iloc[10, j]) else ""
        row10_clean = row10.replace('nan', '').strip()
        original_headers[j] = row10_clean
    
    # Compare with our headers
    our_headers = list(our_supply.columns)[1:]  # Skip 'Date' column
    
    mismatches = 0
    for i, (col_idx, original_header) in enumerate(original_headers.items()):
        if i < len(our_headers):
            our_header = our_headers[i]
            match = "✅" if original_header == our_header else "❌"
            if original_header != our_header:
                mismatches += 1
            print(f"{col_idx:<4} {original_header[:24]:<25} {our_header[:24]:<25} {match:<6}")
        else:
            print(f"{col_idx:<4} {original_header[:24]:<25} {'MISSING':<25} ❌")
            mismatches += 1
    
    print(f"\n📊 HEADER COMPARISON SUMMARY:")
    print(f"   Total columns checked: {len(original_headers)}")
    print(f"   Mismatches: {mismatches}")
    print(f"   Accuracy: {((len(original_headers) - mismatches) / len(original_headers) * 100):.1f}%")
    
    # Now check actual data values for a few key dates
    print(f"\n📊 DATA VALUES COMPARISON (Sample dates):")
    
    test_dates = ['2016-10-01', '2016-10-04', '2016-10-10']
    data_start_row = 12
    
    for test_date in test_dates:
        print(f"\n📅 Date: {test_date}")
        
        # Find the row in original data
        original_row_idx = None
        for i in range(data_start_row, min(data_start_row + 50, original_data.shape[0])):
            date_val = original_data.iloc[i, 1]
            if pd.notna(date_val) and str(date_val)[:10] == test_date:
                original_row_idx = i
                break
        
        # Find the row in our data
        our_row = our_supply[our_supply['Date'].dt.strftime('%Y-%m-%d') == test_date]
        
        if original_row_idx is not None and not our_row.empty:
            our_row = our_row.iloc[0]
            
            # Compare key supply columns
            key_columns = ['Norway', 'Russia (Nord Stream)', 'LNG', 'Algeria', 'Netherlands', 'Total']
            
            print(f"   {'Column':<20} {'Original':<12} {'Our Value':<12} {'Diff':<10} {'Match':<6}")
            print(f"   {'-'*60}")
            
            for col_name in key_columns:
                if col_name in our_row.index:
                    # Find corresponding column in original
                    original_col_idx = None
                    for j, header in original_headers.items():
                        if header == col_name:
                            original_col_idx = j
                            break
                    
                    if original_col_idx is not None:
                        original_val = original_data.iloc[original_row_idx, original_col_idx]
                        our_val = our_row[col_name]
                        
                        if pd.notna(original_val) and pd.notna(our_val):
                            diff = abs(original_val - our_val)
                            match = "✅" if diff < 0.001 else "❌"
                            print(f"   {col_name:<20} {original_val:<12.3f} {our_val:<12.3f} {diff:<10.6f} {match:<6}")
                        else:
                            print(f"   {col_name:<20} {'NaN':<12} {'NaN':<12} {'---':<10} ❓")
    
    # Final verification - check if we missed any important columns
    print(f"\n🔍 CHECKING FOR MISSED COLUMNS:")
    
    # Look at all columns beyond 38 to see if we missed important supply data
    important_beyond_38 = []
    for j in range(39, min(50, original_data.shape[1])):
        row9 = str(original_data.iloc[9, j]) if pd.notna(original_data.iloc[9, j]) else ""
        row10 = str(original_data.iloc[10, j]) if pd.notna(original_data.iloc[10, j]) else ""
        
        # Check if these look like important supply columns
        supply_keywords = ['supply', 'import', 'export', 'production', 'total']
        if any(keyword in (row9 + row10).lower() for keyword in supply_keywords):
            sample_val = original_data.iloc[15, j] if pd.notna(original_data.iloc[15, j]) else "N/A"
            important_beyond_38.append((j, row9, row10, sample_val))
    
    if important_beyond_38:
        print(f"   ⚠️  Found {len(important_beyond_38)} potentially important columns beyond column 38:")
        for col_idx, row9, row10, sample in important_beyond_38:
            print(f"   Col {col_idx}: {row9} | {row10} (Sample: {sample})")
    else:
        print(f"   ✅ No important supply columns found beyond column 38")
    
    return mismatches == 0

if __name__ == "__main__":
    match_perfect = verify_supply_columns_match()
    
    if match_perfect:
        print(f"\n✅ VERIFICATION PASSED: Supply columns match perfectly!")
    else:
        print(f"\n❌ VERIFICATION FAILED: Some discrepancies found")