#!/usr/bin/env python3
"""
Find actual supply routes in Excel to correct our mapping.
"""

import pandas as pd
import sys
sys.path.append("C:/development/commodities")

def find_all_supply_routes():
    print("🔍 FINDING ALL ACTUAL SUPPLY ROUTES")
    print("=" * 80)
    
    # Load Excel file
    df_full = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    # Check structure
    category_row = 14 - 1      # Row 14 (0-based = 13)  
    subcategory_row = 15 - 1   # Row 15 (0-based = 14)
    third_level_row = 16 - 1   # Row 16 (0-based = 15)
    
    print(f"Scanning first 100 columns for ALL supply routes:")
    print("\nCol | Category              | Subcategory          | Third Level          ")
    print("----|----------------------|----------------------|----------------------")
    
    supply_routes = []
    for col_idx in range(2, min(100, len(df_full.columns))):  # Columns C onwards
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
            excel_col = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
            print(f"{excel_col:3} | {cat_str:20} | {sub_str:20} | {third_str:20}")
            
            supply_routes.append({
                'excel_col': excel_col,
                'col_idx': col_idx,
                'category': cat_str,
                'subcategory': sub_str, 
                'third_level': third_str,
                'route_name': f"{sub_str}_to_{third_str}" if cat_str.lower() == 'import' else f"{sub_str}_Production" if cat_str.lower() == 'production' else f"{sub_str}_Export_{third_str}"
            })
    
    print(f"\n✅ Found {len(supply_routes)} supply-related columns")
    
    # Group by categories
    imports = [r for r in supply_routes if 'import' in r['category'].lower()]
    production = [r for r in supply_routes if 'production' in r['category'].lower()]
    exports = [r for r in supply_routes if 'export' in r['category'].lower()]
    
    print(f"\n📊 SUPPLY ROUTE CATEGORIES:")
    print(f"   Imports: {len(imports)}")
    print(f"   Production: {len(production)}")
    print(f"   Exports: {len(exports)}")
    
    # Show unique subcategories for each type
    print(f"\n🔍 IMPORT SOURCES:")
    import_sources = list(set(r['subcategory'] for r in imports))
    for i, source in enumerate(sorted(import_sources), 1):
        print(f"   {i:2d}. {source}")
    
    print(f"\n🔍 PRODUCTION COUNTRIES:")
    prod_countries = list(set(r['subcategory'] for r in production))
    for i, country in enumerate(sorted(prod_countries), 1):
        print(f"   {i:2d}. {country}")
    
    print(f"\n🔍 EXPORT SOURCES:")
    export_sources = list(set(r['subcategory'] for r in exports))
    for i, source in enumerate(sorted(export_sources), 1):
        print(f"   {i:2d}. {source}")
    
    # Look for specific routes that should be in Excel
    print(f"\n🎯 SPECIFIC ROUTE SEARCH:")
    target_searches = [
        ('Slovakia', 'Austria'),
        ('Russia', 'Germany'),
        ('Algeria', 'Italy'),
        ('Libya', 'Italy'),
        ('Denmark', 'Germany'),
        ('Czech', 'Germany'),
        ('Poland', 'Germany')
    ]
    
    for search_sub, search_third in target_searches:
        matches = []
        for route in supply_routes:
            sub_match = search_sub.lower() in route['subcategory'].lower()
            third_match = search_third.lower() in route['third_level'].lower()
            if sub_match and third_match:
                matches.append(route)
        
        if matches:
            print(f"   {search_sub} → {search_third}:")
            for match in matches:
                print(f"      {match['excel_col']}: {match['category']} | {match['subcategory']} | {match['third_level']}")
        else:
            print(f"   {search_sub} → {search_third}: ❌ NOT FOUND")
    
    return supply_routes

if __name__ == "__main__":
    routes = find_all_supply_routes()