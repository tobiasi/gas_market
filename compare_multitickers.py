# -*- coding: utf-8 -*-
"""
DIRECT COMPARISON: Compare the two MultiTicker files
"""

import pandas as pd
import numpy as np

def compare_multitickers():
    """Direct comparison of the two MultiTicker files"""
    
    print(f"\n{'='*80}")
    print("📊 MULTITICKER FILES COMPARISON")
    print(f"{'='*80}")
    
    try:
        target_mt = pd.read_csv('the_multiticker_we_are_trying_to_replicate.csv', low_memory=False)
        our_mt = pd.read_csv('our_multiticker.csv', low_memory=False)
        
        print(f"📋 FILE OVERVIEW:")
        print(f"   Target shape: {target_mt.shape}")
        print(f"   Our shape: {our_mt.shape}")
        print(f"   Difference: {our_mt.shape[1] - target_mt.shape[1]} columns, {our_mt.shape[0] - target_mt.shape[0]} rows")
        
        # Compare data starting from row where actual values begin (after headers)
        data_start_row = 5  # Skip header rows
        
        # Find common columns (by position, since names might differ)
        min_cols = min(target_mt.shape[1], our_mt.shape[1])
        min_rows = min(target_mt.shape[0], our_mt.shape[0])
        
        print(f"\n🔍 DATA COMPARISON (starting from row {data_start_row}):")
        
        # Compare first 10 data rows across first 10 columns
        differences_found = 0
        total_comparisons = 0
        
        print(f"{'Row':<5} {'Col':<5} {'Target':<15} {'Ours':<15} {'Difference':<15}")
        print("-" * 70)
        
        for row in range(data_start_row, min(data_start_row + 10, min_rows)):
            for col in range(min(10, min_cols)):
                target_val = target_mt.iloc[row, col]
                our_val = our_mt.iloc[row, col]
                
                try:
                    target_num = float(target_val) if pd.notna(target_val) and target_val != '' else 0
                    our_num = float(our_val) if pd.notna(our_val) and our_val != '' else 0
                    diff = abs(our_num - target_num)
                    
                    total_comparisons += 1
                    
                    if diff > 0.001:  # Significant difference
                        differences_found += 1
                        print(f"{row:<5} {col:<5} {target_num:<15.3f} {our_num:<15.3f} {diff:<15.3f}")
                        
                        if differences_found >= 20:  # Limit output
                            break
                            
                except (ValueError, TypeError):
                    # Non-numeric data, compare as strings
                    if str(target_val) != str(our_val):
                        differences_found += 1
                        target_str = str(target_val)[:12] + "..." if len(str(target_val)) > 12 else str(target_val)
                        our_str = str(our_val)[:12] + "..." if len(str(our_val)) > 12 else str(our_val)
                        print(f"{row:<5} {col:<5} {target_str:<15} {our_str:<15} {'TEXT_DIFF':<15}")
            
            if differences_found >= 20:
                break
        
        print(f"\n📈 SUMMARY:")
        print(f"   Comparisons made: {total_comparisons}")
        print(f"   Differences found: {differences_found}")
        
        if differences_found == 0:
            print(f"   ✅ DATA APPEARS IDENTICAL in sampled area")
        elif differences_found < total_comparisons * 0.1:
            print(f"   ⚠️  MINOR DIFFERENCES found ({differences_found/total_comparisons*100:.1f}%)")
        else:
            print(f"   ❌ SIGNIFICANT DIFFERENCES found ({differences_found/total_comparisons*100:.1f}%)")
        
        # Check if the files have the same number of series
        if target_mt.shape[1] != our_mt.shape[1]:
            print(f"\n📊 COLUMN COUNT DIFFERENCE:")
            print(f"   Target has {target_mt.shape[1]} series")
            print(f"   We have {our_mt.shape[1]} series")
            print(f"   Difference: {our_mt.shape[1] - target_mt.shape[1]} series")
            
            if our_mt.shape[1] > target_mt.shape[1]:
                print(f"   → We have EXTRA series")
            else:
                print(f"   → We have MISSING series")
        
        # Check if files have same date range
        if target_mt.shape[0] != our_mt.shape[0]:
            print(f"\n📅 ROW COUNT DIFFERENCE:")
            print(f"   Target has {target_mt.shape[0]} rows")
            print(f"   We have {our_mt.shape[0]} rows") 
            print(f"   Difference: {our_mt.shape[0] - target_mt.shape[0]} rows")
            
            if our_mt.shape[0] > target_mt.shape[0]:
                print(f"   → We have MORE data (longer time series)")
            else:
                print(f"   → We have LESS data (shorter time series)")
        
        return differences_found, total_comparisons
        
    except Exception as e:
        print(f"❌ Error in comparison: {e}")
        import traceback
        traceback.print_exc()
        return None, None

if __name__ == "__main__":
    diffs, total = compare_multitickers()