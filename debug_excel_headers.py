#!/usr/bin/env python3
"""
Debug script to check actual header values in Excel.
"""

import pandas as pd
import sys
sys.path.append("C:/development/commodities")

def debug_excel_headers():
    print("🔍 DEBUGGING EXCEL HEADERS")
    print("=" * 80)
    
    # Load Excel file
    df_full = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    # Check structure
    category_row = 14 - 1      # Row 14 (0-based = 13)  
    subcategory_row = 15 - 1   # Row 15 (0-based = 14)
    third_level_row = 16 - 1   # Row 16 (0-based = 15)
    
    print(f"Checking first 20 columns for supply-related data:")
    print("\nCol | Category              | Subcategory          | Third Level          ")
    print("----|----------------------|----------------------|----------------------")
    
    supply_found = 0
    for col_idx in range(2, 22):  # Columns C to V
        if col_idx >= len(df_full.columns):
            break
            
        category = df_full.iloc[category_row, col_idx]
        subcategory = df_full.iloc[subcategory_row, col_idx] 
        third_level = df_full.iloc[third_level_row, col_idx]
        
        # Look for supply-related content
        cat_str = str(category) if pd.notna(category) else ""
        sub_str = str(subcategory) if pd.notna(subcategory) else ""
        third_str = str(third_level) if pd.notna(third_level) else ""
        
        is_supply = any(keyword in cat_str.lower() for keyword in ['import', 'export', 'production'])
        
        if is_supply:
            excel_col = chr(65 + col_idx)  # A, B, C, etc.
            print(f"{excel_col:3} | {cat_str:20} | {sub_str:20} | {third_str:20}")
            supply_found += 1
    
    print(f"\n✅ Found {supply_found} supply-related columns in first 20")
    
    # Now check for specific routes we're looking for
    print(f"\n🎯 SEARCHING FOR TARGET ROUTES:")
    target_routes = [
        ('Slovakia', 'Austria'),
        ('Russia', 'Germany'), 
        ('Norway', 'Europe'),
        ('Netherlands', 'Netherlands'),
        ('GB', 'GB'),
        ('LNG', '*'),
        ('Algeria', 'Italy'),
        ('Libya', 'Italy')
    ]
    
    for target_sub, target_third in target_routes:
        print(f"\n🔍 Looking for {target_sub} → {target_third}:")
        matches = []
        
        for col_idx in range(2, min(100, len(df_full.columns))):  # Check first 100 cols
            category = str(df_full.iloc[category_row, col_idx]) if pd.notna(df_full.iloc[category_row, col_idx]) else ""
            subcategory = str(df_full.iloc[subcategory_row, col_idx]) if pd.notna(df_full.iloc[subcategory_row, col_idx]) else ""
            third_level = str(df_full.iloc[third_level_row, col_idx]) if pd.notna(df_full.iloc[third_level_row, col_idx]) else ""
            
            # Check for matches
            sub_match = target_sub.lower() in subcategory.lower()
            third_match = target_third == '*' or target_third.lower() in third_level.lower()
            supply_cat = any(keyword in category.lower() for keyword in ['import', 'export', 'production'])
            
            if sub_match and third_match and supply_cat:
                excel_col = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
                matches.append((excel_col, category, subcategory, third_level))
        
        if matches:
            for excel_col, cat, sub, third in matches:
                print(f"   ✅ {excel_col}: {cat} | {sub} | {third}")
        else:
            print(f"   ❌ No matches found")

if __name__ == "__main__":
    debug_excel_headers()