# -*- coding: utf-8 -*-
"""
Debug the supply column ordering issue
"""

import pandas as pd

def debug_supply_column_order():
    """Debug why column headers are misaligned"""
    
    print("🔧 DEBUGGING SUPPLY COLUMN ORDER")
    print("=" * 80)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    supply_start_col = 17
    supply_end_col = 38
    
    print(f"📋 CORRECT ORDER (from LiveSheet columns {supply_start_col}-{supply_end_col}):")
    
    correct_order = []
    for j in range(supply_start_col, supply_end_col + 1):
        row10 = str(target_data.iloc[10, j]) if pd.notna(target_data.iloc[10, j]) else ""
        row10_clean = row10.replace('nan', '').strip()
        
        if row10_clean:
            header_name = row10_clean
        else:
            header_name = f'Col_{j}'
        
        correct_order.append((j, header_name))
        print(f"   Col {j:2d}: {header_name}")
    
    # Now check how our master code is organizing them
    print(f"\n🔍 HOW OUR CODE IS ORGANIZING THEM:")
    
    # Simulate the logic from gas_market_master.py
    supply_columns = {}
    
    for j in range(supply_start_col, supply_end_col + 1):
        row9 = str(target_data.iloc[9, j]) if pd.notna(target_data.iloc[9, j]) else ""
        row10 = str(target_data.iloc[10, j]) if pd.notna(target_data.iloc[10, j]) else ""
        row11 = str(target_data.iloc[11, j]) if pd.notna(target_data.iloc[11, j]) else ""
        
        row9_clean = row9.replace('nan', '').strip()
        row10_clean = row10.replace('nan', '').strip()
        row11_clean = row11.replace('nan', '').strip()
        
        if row10_clean:
            header_name = row10_clean
        elif row9_clean:
            header_name = row9_clean
        elif row11_clean:
            header_name = row11_clean
        else:
            header_name = f'Col_{j}'
        
        supply_columns[j] = {
            'header': header_name,
            'category': row9_clean,
            'subcategory': row10_clean,
            'detail': row11_clean
        }
    
    # Group by category (this is where the reordering happens)
    imports = []
    production = []
    exports = []
    totals = []
    others = []
    
    for j, info in supply_columns.items():
        header = info['header']
        category = info['category'].lower()
        
        if 'import' in category:
            imports.append((j, header))
        elif 'production' in category:
            production.append((j, header))
        elif 'export' in category:
            exports.append((j, header))
        elif 'supply' in category or 'total' in header.lower():
            totals.append((j, header))
        else:
            others.append((j, header))
    
    print(f"\n📊 OUR CATEGORIZED ORDER:")
    print(f"   Imports ({len(imports)}):")
    for j, header in imports:
        print(f"     Col {j:2d}: {header}")
    
    print(f"   Production ({len(production)}):")
    for j, header in production:
        print(f"     Col {j:2d}: {header}")
    
    print(f"   Exports ({len(exports)}):")
    for j, header in exports:
        print(f"     Col {j:2d}: {header}")
    
    print(f"   Totals ({len(totals)}):")
    for j, header in totals:
        print(f"     Col {j:2d}: {header}")
    
    print(f"   Others ({len(others)}):")
    for j, header in others:
        print(f"     Col {j:2d}: {header}")
    
    # The problem: we're reordering by category, but we should preserve original order!
    print(f"\n❌ PROBLEM IDENTIFIED:")
    print(f"   We're reordering columns by category instead of preserving original order")
    print(f"   This causes header misalignment even though data values are correct")
    
    print(f"\n✅ SOLUTION:")
    print(f"   Keep columns in original order (17-38) without categorization reordering")
    
    return correct_order

if __name__ == "__main__":
    correct_order = debug_supply_column_order()