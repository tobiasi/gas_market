# -*- coding: utf-8 -*-
"""
List ALL supply columns from the Daily historic data by category sheet
"""

import pandas as pd

def list_all_supply_columns():
    """List all supply columns with their hierarchical structure"""
    
    print("⛽ ALL SUPPLY COLUMNS IN DAILY HISTORIC DATA BY CATEGORY")
    print("=" * 100)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    # Read the sheet
    data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    print(f"📊 Sheet shape: {data.shape}")
    print(f"📋 Analyzing columns 17-{data.shape[1]-1} (supply section)")
    
    # We know demand ends around column 15, supply starts around 17
    supply_start = 17
    
    print(f"\n📊 COMPLETE SUPPLY COLUMN STRUCTURE:")
    print(f"{'Col':<4} {'Row 9 (Category)':<25} {'Row 10 (Subcategory)':<25} {'Row 11 (Detail)':<25} {'Sample Value':<15}")
    print("-" * 100)
    
    supply_columns = {}
    categories = {}
    
    # Extract all supply columns
    for j in range(supply_start, data.shape[1]):
        row9 = str(data.iloc[9, j]) if pd.notna(data.iloc[9, j]) else ""
        row10 = str(data.iloc[10, j]) if pd.notna(data.iloc[10, j]) else ""
        row11 = str(data.iloc[11, j]) if pd.notna(data.iloc[11, j]) else ""
        
        # Get sample data value (from row 15 - our target row)
        sample_val = data.iloc[15, j] if j < data.shape[1] and pd.notna(data.iloc[15, j]) else ""
        sample_str = f"{sample_val:.2f}" if isinstance(sample_val, (int, float)) else str(sample_val)[:10]
        
        # Clean up the text
        row9_clean = row9.replace('nan', '').strip()
        row10_clean = row10.replace('nan', '').strip()
        row11_clean = row11.replace('nan', '').strip()
        
        # Only show columns with meaningful content
        if any(x for x in [row9_clean, row10_clean, row11_clean]):
            print(f"{j:<4} {row9_clean[:24]:<25} {row10_clean[:24]:<25} {row11_clean[:24]:<25} {sample_str:<15}")
            
            # Store for categorization
            supply_columns[j] = {
                'category': row9_clean,
                'subcategory': row10_clean, 
                'detail': row11_clean,
                'sample_value': sample_val
            }
            
            # Group by category
            if row9_clean:
                if row9_clean not in categories:
                    categories[row9_clean] = []
                categories[row9_clean].append((j, row10_clean, row11_clean, sample_val))
    
    # Organized view by category
    print(f"\n📋 SUPPLY COLUMNS ORGANIZED BY CATEGORY:")
    print("=" * 100)
    
    for category, columns in categories.items():
        if category:  # Skip empty categories
            print(f"\n🔹 {category.upper()}:")
            for col_idx, subcategory, detail, sample_val in columns:
                sample_str = f"{sample_val:.2f}" if isinstance(sample_val, (int, float)) else "N/A"
                subcategory_display = subcategory if subcategory else detail
                print(f"   Col {col_idx:2d}: {subcategory_display:<30} (Sample: {sample_str})")
    
    # Special analysis for orphaned columns (no clear category)
    orphaned = []
    for j, info in supply_columns.items():
        if not info['category'] and (info['subcategory'] or info['detail']):
            orphaned.append((j, info['subcategory'], info['detail'], info['sample_value']))
    
    if orphaned:
        print(f"\n🔹 OTHER/UNCATEGORIZED:")
        for col_idx, subcategory, detail, sample_val in orphaned:
            sample_str = f"{sample_val:.2f}" if isinstance(sample_val, (int, float)) else "N/A"
            label = subcategory if subcategory else detail
            print(f"   Col {col_idx:2d}: {label:<30} (Sample: {sample_str})")
    
    # Summary statistics
    print(f"\n📊 SUPPLY SECTION SUMMARY:")
    print(f"   Total supply columns: {len(supply_columns)}")
    print(f"   Categories found: {len([c for c in categories.keys() if c])}")
    print(f"   Column range: {supply_start} to {max(supply_columns.keys()) if supply_columns else 'N/A'}")
    
    # Show key totals
    print(f"\n🎯 KEY SUPPLY TOTALS (Sample from 2016-10-04):")
    for col_idx, info in supply_columns.items():
        if 'total' in info['subcategory'].lower() and isinstance(info['sample_value'], (int, float)):
            print(f"   {info['subcategory']}: {info['sample_value']:.2f}")
    
    return supply_columns

if __name__ == "__main__":
    supply_cols = list_all_supply_columns()