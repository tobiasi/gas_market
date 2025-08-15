# -*- coding: utf-8 -*-
"""
Analyze target structure from "Daily historic data by category" tab
"""

import pandas as pd

def analyze_target_structure():
    """Analyze the target structure we need to replicate"""
    
    print(f"\n{'='*80}")
    print("🎯 ANALYZING TARGET STRUCTURE")
    print(f"{'='*80}")
    
    # Read the target CSV
    target = pd.read_csv('The_daily_we_want_to_replicate.csv')
    
    print(f"📊 Target shape: {target.shape}")
    print(f"📋 First few columns: {list(target.columns[:15])}")
    
    # Analyze the structure
    print(f"\n🔍 STRUCTURE ANALYSIS:")
    print("Row 0 (Categories):", target.iloc[0, :15].tolist())
    print("Row 1 (Countries):", target.iloc[1, :15].tolist()) 
    print("Row 2 (Subcategories):", target.iloc[2, :15].tolist())
    
    # Check Italy specifically
    print(f"\n🇮🇹 ITALY ANALYSIS:")
    italy_col_idx = None
    for i, country in enumerate(target.iloc[1]):
        if country == 'Italy':
            italy_col_idx = i
            break
    
    if italy_col_idx:
        print(f"Italy column index: {italy_col_idx}")
        italy_values = target.iloc[3:8, italy_col_idx].astype(float)
        print(f"Italy first 5 values: {italy_values.tolist()}")
        print(f"Italy average: {italy_values.mean():.2f}")
    
    # Find column positions for totals
    print(f"\n📊 TOTALS ANALYSIS:")
    total_idx, industrial_idx, ldz_idx, gtp_idx = None, None, None, None
    
    for i, country in enumerate(target.iloc[1]):
        if country == 'Total':
            total_idx = i
        elif country == 'Industrial':
            industrial_idx = i
        elif country == 'LDZ':
            ldz_idx = i
        elif country == 'Gas-to-Power':
            gtp_idx = i
    
    if all(idx is not None for idx in [total_idx, industrial_idx, ldz_idx, gtp_idx]):
        print(f"Column positions - Total:{total_idx}, Industrial:{industrial_idx}, LDZ:{ldz_idx}, GtP:{gtp_idx}")
        
        # Check if they sum correctly
        for row_idx in range(3, 8):  # First 5 data rows
            total_val = float(target.iloc[row_idx, total_idx])
            industrial_val = float(target.iloc[row_idx, industrial_idx])
            ldz_val = float(target.iloc[row_idx, ldz_idx])
            gtp_val = float(target.iloc[row_idx, gtp_idx])
            
            sum_categories = industrial_val + ldz_val + gtp_val
            difference = abs(total_val - sum_categories)
            
            print(f"Row {row_idx}: Total={total_val:.2f}, Sum={sum_categories:.2f}, Diff={difference:.6f}")
    
    # Check our output
    print(f"\n📋 OUR OUTPUT ANALYSIS:")
    try:
        our_output = pd.read_csv('our_daily.csv')
        print(f"Our shape: {our_output.shape}")
        
        # Get first data row values
        if our_output.shape[0] > 1:
            our_row = our_output.iloc[1]  # Skip header
            our_italy = our_row['Italy'] if 'Italy' in our_row.index else 'Not found'
            our_total = our_row['Total'] if 'Total' in our_row.index else 'Not found'
            our_industrial = our_row['Industrial'] if 'Industrial' in our_row.index else 'Not found'
            our_ldz = our_row['LDZ'] if 'LDZ' in our_row.index else 'Not found'
            our_gtp = our_row['Gas-to-Power'] if 'Gas-to-Power' in our_row.index else 'Not found'
            
            print(f"Our Italy: {our_italy}")
            print(f"Our Total: {our_total}")
            print(f"Our Industrial: {our_industrial}")
            print(f"Our LDZ: {our_ldz}")
            print(f"Our GtP: {our_gtp}")
            
            if all(isinstance(val, (int, float)) for val in [our_total, our_industrial, our_ldz, our_gtp]):
                our_sum = our_industrial + our_ldz + our_gtp
                our_diff = abs(our_total - our_sum)
                print(f"Our sum check: {our_sum:.2f} vs {our_total:.2f}, diff: {our_diff:.6f}")
        
    except Exception as e:
        print(f"Could not read our output: {e}")
    
    # Compare Italy values
    print(f"\n🎯 ITALY COMPARISON:")
    if italy_col_idx and 'our_italy' in locals() and isinstance(our_italy, (int, float)):
        target_italy = float(target.iloc[3, italy_col_idx])  # First data row
        print(f"Target Italy: {target_italy:.2f}")
        print(f"Our Italy: {our_italy:.2f}")
        print(f"Difference: {abs(target_italy - our_italy):.2f}")
        
        if abs(target_italy - our_italy) > 0.1:
            print("❌ Italy values don't match!")
            print("💡 Check:")
            print("   1. Normalization factors")
            print("   2. Which series are included for Italy")
            print("   3. Aggregation logic")
        else:
            print("✅ Italy values match!")

if __name__ == "__main__":
    analyze_target_structure()