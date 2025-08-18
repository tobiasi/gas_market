#!/usr/bin/env python3
"""
Debug EXACT matches found for Czech_and_Poland to eliminate extra 3.37.
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd

def debug_exact_czech_matches():
    print("🔍 DEBUGGING EXACT Czech_and_Poland MATCHES")
    print("=" * 70)
    
    # Load MultiTicker
    multiticker_df = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    max_col = min(400, len(multiticker_df.columns))
    
    # Get criteria headers
    criteria_row1 = multiticker_df.iloc[13, 2:max_col].values  # Category
    criteria_row2 = multiticker_df.iloc[14, 2:max_col].values  # Subcategory  
    criteria_row3 = multiticker_df.iloc[15, 2:max_col].values  # Third level
    
    # Get first data row (2016-10-02)
    data_row = multiticker_df.iloc[19, 2:max_col].values
    
    # Target criteria for Czech_and_Poland
    target_category = "Import"
    target_subcategory = "Czech and Poland" 
    target_third_level = "Germany"
    
    print(f"🎯 Target SUMIFS criteria:")
    print(f"   Category: '{target_category}'")
    print(f"   Subcategory: '{target_subcategory}'")
    print(f"   Third Level: '{target_third_level}'")
    print(f"   Target value: 58.41")
    
    print(f"\n🔍 ALL EXACT MATCHES found:")
    print("   Excel | Index | Category | Subcategory     | Third Level | Value    | Include?")
    print("   " + "-" * 85)
    
    exact_matches = []
    total_sum = 0.0
    
    for i in range(len(criteria_row1)):
        if i >= len(data_row):
            continue
            
        c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
        c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
        c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
        
        # Check for EXACT match
        c1_match = c1 == target_category
        c2_match = c2 == target_subcategory
        c3_match = c3 == target_third_level
        
        if c1_match and c2_match and c3_match:
            value = data_row[i] if pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
            value = float(value)
            
            # Calculate Excel column
            col_idx = i + 2  # Add offset
            if col_idx < 26:
                excel_col = chr(65 + col_idx)
            else:
                first_letter = chr(65 + (col_idx // 26) - 1)
                second_letter = chr(65 + (col_idx % 26))
                excel_col = first_letter + second_letter
            
            exact_matches.append({
                'excel_col': excel_col,
                'index': i,
                'category': c1,
                'subcategory': c2,
                'third_level': c3,
                'value': value
            })
            
            total_sum += value
            
            print(f"   {excel_col:5} | {i:5d} | {c1:8} | {c2:15} | {c3:11} | {value:8.2f} | ?")
    
    print("   " + "-" * 85)
    print(f"   TOTAL SUM: {total_sum:.2f}")
    print(f"   TARGET: 58.41")
    print(f"   EXCESS: {total_sum - 58.41:.2f}")
    
    if len(exact_matches) == 0:
        print("   ❌ No exact matches found!")
        return
    
    print(f"\n🎯 ANALYSIS:")
    print(f"   Found {len(exact_matches)} exact matches")
    print(f"   Need to eliminate {total_sum - 58.41:.2f} worth of matches")
    
    # Sort matches by value to see which ones to potentially exclude
    exact_matches.sort(key=lambda x: x['value'], reverse=True)
    
    print(f"\n📊 Matches sorted by value (highest first):")
    for i, match in enumerate(exact_matches):
        print(f"   {i+1}. {match['excel_col']}: {match['value']:.2f}")
    
    # Try different combinations to get exactly 58.41
    print(f"\n🔧 FINDING CORRECT COMBINATION:")
    
    # Strategy 1: Use only the largest match closest to 58.41
    if exact_matches:
        best_single = min(exact_matches, key=lambda x: abs(x['value'] - 58.41))
        print(f"   Best single match: {best_single['excel_col']} = {best_single['value']:.2f} (diff: {abs(best_single['value'] - 58.41):.2f})")
    
    # Strategy 2: Try combinations that sum to ~58.41
    target = 58.41
    print(f"\n   Trying combinations that sum to ~{target:.2f}:")
    
    # Check all possible combinations
    from itertools import combinations
    
    best_combo = None
    best_diff = float('inf')
    
    for r in range(1, len(exact_matches) + 1):
        for combo in combinations(exact_matches, r):
            combo_sum = sum(m['value'] for m in combo)
            diff = abs(combo_sum - target)
            
            if diff < best_diff:
                best_diff = diff
                best_combo = combo
            
            if diff < 0.1:  # Very close match
                cols = [m['excel_col'] for m in combo]
                print(f"   ✅ {'+'.join(cols)} = {combo_sum:.2f} (diff: {diff:.2f})")
    
    if best_combo:
        cols = [m['excel_col'] for m in best_combo]
        combo_sum = sum(m['value'] for m in best_combo)
        print(f"\n🎯 BEST COMBINATION: {'+'.join(cols)} = {combo_sum:.2f}")
        print(f"   This should be used for Czech_and_Poland SUMIFS")
        
        return best_combo
    
    return exact_matches

if __name__ == "__main__":
    result = debug_exact_czech_matches()