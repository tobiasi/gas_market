#!/usr/bin/env python3
"""
Find the missing 6.56 in supply routes to match 754.38 exactly.
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd

def find_missing_routes():
    print("🔍 FINDING MISSING 6.56 IN SUPPLY ROUTES")
    print("=" * 60)
    
    # Load Daily sheet to check ALL supply columns R through AI
    daily_df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
    
    # All supply columns from Daily sheet
    supply_columns = ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 
                     'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']
    
    print("📊 ALL Daily Sheet Supply Routes (columns R-AI):")
    print("   Col | Category        | Subcategory     | Third Level     | Row 13 Value")
    print("   " + "-" * 80)
    
    total_daily_value = 0.0
    
    for col_letter in supply_columns:
        # Convert to index
        if len(col_letter) == 1:
            col_idx = ord(col_letter) - ord('A')
        else:
            col_idx = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
        
        category = str(daily_df.iloc[9, col_idx]).strip()
        subcategory = str(daily_df.iloc[10, col_idx]).strip()  
        third_level = str(daily_df.iloc[11, col_idx]).strip()
        
        # Get the value from row 13 (first data row in Daily sheet)
        daily_value = daily_df.iloc[12, col_idx] if pd.notna(daily_df.iloc[12, col_idx]) else 0
        daily_value = float(daily_value)
        
        total_daily_value += daily_value
        
        print(f"   {col_letter:3} | {category:15} | {subcategory:15} | {third_level:15} | {daily_value:9.2f}")
    
    print("   " + "-" * 80)
    print(f"   TOTAL Daily Sheet Supply: {total_daily_value:.2f}")
    
    # Compare with our calculated routes
    our_routes = {
        'Slovakia': 108.87,
        'Russia (Nord Stream)': 121.08,
        'Norway': 206.67,
        'Netherlands': 79.69,
        'GB': 94.01,
        'LNG': 47.86,
        'Algeria': 25.33,
        'Libya': 10.98,
        'Spain': -8.43,
        'Czech and Poland': 61.78  # Our corrected value
    }
    
    our_total = sum(our_routes.values())
    print(f"\n   Our calculated total: {our_total:.2f}")
    print(f"   Daily sheet total: {total_daily_value:.2f}")
    print(f"   Difference: {total_daily_value - our_total:.2f}")
    
    # Check which routes we might be missing
    print(f"\n🔍 Routes we might be missing:")
    
    # All possible routes from Daily sheet
    all_routes = []
    for col_letter in supply_columns:
        if len(col_letter) == 1:
            col_idx = ord(col_letter) - ord('A')
        else:
            col_idx = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
        
        category = str(daily_df.iloc[9, col_idx]).strip()
        subcategory = str(daily_df.iloc[10, col_idx]).strip()  
        third_level = str(daily_df.iloc[11, col_idx]).strip()
        daily_value = float(daily_df.iloc[12, col_idx]) if pd.notna(daily_df.iloc[12, col_idx]) else 0
        
        route_key = f"{subcategory}"
        if route_key not in [key.split(' (')[0] for key in our_routes.keys()]:
            print(f"   Missing: {col_letter} - {subcategory} → {third_level} = {daily_value:.2f}")
    
    print(f"\n🎯 TARGET: Find routes that sum to ~6.56 to reach 754.38")

if __name__ == "__main__":
    find_missing_routes()