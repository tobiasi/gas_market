#!/usr/bin/env python3
"""
Debug Czech_and_Poland SUMIFS to find overcounting issue.
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd
import numpy as np

def debug_czech_poland_sumifs():
    print("🔍 DEBUGGING Czech_and_Poland SUMIFS")
    print("=" * 60)
    
    # Load Daily sheet to check exact criteria for column AB
    daily_df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
    
    # Column AB = 27 (A=0, B=1, ..., AB=27)
    ab_col_idx = 27
    
    print("📊 Daily Sheet Column AB Criteria:")
    category = str(daily_df.iloc[9, ab_col_idx]).strip()
    subcategory = str(daily_df.iloc[10, ab_col_idx]).strip()  
    third_level = str(daily_df.iloc[11, ab_col_idx]).strip()
    
    print(f"   Category: '{category}'")
    print(f"   Subcategory: '{subcategory}'")
    print(f"   Third Level: '{third_level}'")
    
    # Load MultiTicker and check criteria headers
    multiticker_df = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    # Extract criteria headers (wide scan)
    max_col = min(400, len(multiticker_df.columns))
    criteria_row1 = multiticker_df.iloc[13, 2:max_col].values  # Category
    criteria_row2 = multiticker_df.iloc[14, 2:max_col].values  # Subcategory  
    criteria_row3 = multiticker_df.iloc[15, 2:max_col].values  # Third level
    
    # Get first data row (2016-10-02)
    data_row = multiticker_df.iloc[19, 2:max_col].values
    
    print(f"\n🔍 Scanning MultiTicker for matches:")
    print("   Col | Category        | Subcategory     | Third Level     | Value    | Match?")
    print("   " + "-" * 75)
    
    total_sum = 0.0
    match_count = 0
    matches_found = []
    
    # Search for matches
    for i in range(len(criteria_row1)):
        if i >= len(data_row):
            continue
            
        c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
        c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
        c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
        
        # Check for exact match with Czech_and_Poland criteria
        c1_match = c1 == category
        c2_match = c2 == subcategory
        c3_match = c3 == third_level
        
        all_match = c1_match and c2_match and c3_match
        
        # Also check for potential partial matches that might be causing issues
        partial_match = False
        if "czech" in c2.lower() and "poland" in c2.lower():
            partial_match = True
        
        if all_match or partial_match:
            value = data_row[i] if pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
            excel_col = chr(65 + i + 2) if i + 2 < 26 else f"A{chr(65 + i + 2 - 26)}"
            
            match_type = "EXACT" if all_match else "PARTIAL"
            if all_match:
                total_sum += float(value)
                match_count += 1
                matches_found.append((excel_col, c1, c2, c3, value))
            
            print(f"   {excel_col:3} | {c1:15} | {c2:15} | {c3:15} | {value:8.2f} | {match_type}")
    
    print("   " + "-" * 75)
    print(f"   TOTAL EXACT MATCHES: {match_count}")
    print(f"   TOTAL SUM: {total_sum:.2f}")
    print(f"   EXPECTED: 58.41")
    print(f"   DIFFERENCE: {total_sum - 58.41:.2f}")
    
    if abs(total_sum - 58.41) > 0.1:
        print(f"\n❌ ISSUE CONFIRMED: Overcounting by {total_sum - 58.41:.2f}")
        
        print(f"\n🔍 DETAILED ANALYSIS:")
        if match_count > 1:
            print(f"   - Found {match_count} matches (should probably be 1)")
            print(f"   - Possible double-counting across multiple columns")
        
        print(f"\n🔍 ALL EXACT MATCHES:")
        for excel_col, c1, c2, c3, value in matches_found:
            print(f"   {excel_col}: {c1} | {c2} | {c3} = {value:.2f}")
            
    else:
        print(f"\n✅ SUMIFS working correctly!")
    
    # Also check what variations exist for Czech/Poland
    print(f"\n🔍 ALL Czech/Poland variations in MultiTicker:")
    print("   Col | Category        | Subcategory     | Third Level     | Value")
    print("   " + "-" * 70)
    
    for i in range(len(criteria_row1)):
        if i >= len(data_row):
            continue
            
        c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
        
        if "czech" in c2.lower() or "poland" in c2.lower():
            c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
            c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
            value = data_row[i] if pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
            excel_col = chr(65 + i + 2) if i + 2 < 26 else f"A{chr(65 + i + 2 - 26)}"
            
            print(f"   {excel_col:3} | {c1:15} | {c2:15} | {c3:15} | {value:8.2f}")

if __name__ == "__main__":
    debug_czech_poland_sumifs()