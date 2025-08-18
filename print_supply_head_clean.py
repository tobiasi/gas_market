# -*- coding: utf-8 -*-
"""
Print clean supply data head for user confirmation
"""

import pandas as pd

def print_supply_head_clean():
    """Print clean, formatted supply data head"""
    
    print("⛽ SUPPLY DATA HEAD - FOR CONFIRMATION")
    print("=" * 120)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    
    # Read the main sheet
    data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
    
    # Extract supply section structure
    print("📊 SUPPLY SECTION STRUCTURE:")
    print("\nRows 9-11 (Category headers):")
    
    # Show the hierarchical header structure
    supply_start_col = 17  # Where supply data begins
    supply_end_col = min(60, data.shape[1])  # Check up to column 60
    
    print(f"{'Col':<4} {'Row 9':<20} {'Row 10':<20} {'Row 11':<20}")
    print("-" * 70)
    
    supply_columns = {}
    for j in range(supply_start_col, supply_end_col):
        row9 = str(data.iloc[9, j]) if pd.notna(data.iloc[9, j]) else ""
        row10 = str(data.iloc[10, j]) if pd.notna(data.iloc[10, j]) else ""
        row11 = str(data.iloc[11, j]) if pd.notna(data.iloc[11, j]) else ""
        
        # Only show columns with meaningful headers
        if any(x != "nan" and x != "" for x in [row9, row10, row11]):
            print(f"{j:<4} {row9[:18]:<20} {row10[:18]:<20} {row11[:18]:<20}")
            
            # Build supply column mapping
            if row10 != "nan" and row10 != "":
                supply_columns[j] = row10
    
    print(f"\n📋 IDENTIFIED SUPPLY COLUMNS:")
    for col_idx, header in supply_columns.items():
        print(f"   Column {col_idx:2d}: {header}")
    
    # Print actual supply data head
    print(f"\n📊 SUPPLY DATA HEAD (first 10 rows):")
    
    # Create clean column headers
    clean_headers = []
    clean_col_indices = []
    
    for col_idx in sorted(supply_columns.keys())[:15]:  # First 15 supply columns
        header = supply_columns[col_idx]
        clean_headers.append(header[:15])  # Truncate long headers
        clean_col_indices.append(col_idx)
    
    # Print header row
    print(f"{'Date':<12}", end="")
    for header in clean_headers:
        print(f"{header:<16}", end="")
    print()
    print("-" * (12 + 16 * len(clean_headers)))
    
    # Print data rows
    data_start_row = 12
    for i in range(data_start_row, min(data_start_row + 10, data.shape[0])):
        date_val = data.iloc[i, 1]
        date_str = str(date_val)[:10] if pd.notna(date_val) else "N/A"
        print(f"{date_str:<12}", end="")
        
        for col_idx in clean_col_indices:
            val = data.iloc[i, col_idx]
            if pd.notna(val) and isinstance(val, (int, float)):
                print(f"{val:<16.2f}", end="")
            else:
                print(f"{'NaN':<16}", end="")
        print()
    
    # Identify key supply categories
    print(f"\n🔍 KEY SUPPLY CATEGORIES IDENTIFIED:")
    
    categories = {
        'Pipeline Imports': [],
        'Production': [],
        'LNG Imports': [],
        'Storage/Other': [],
        'Total Supply': []
    }
    
    for col_idx, header in supply_columns.items():
        header_lower = header.lower()
        if 'norway' in header_lower or 'russia' in header_lower or 'algeria' in header_lower:
            categories['Pipeline Imports'].append((col_idx, header))
        elif 'production' in header_lower:
            categories['Production'].append((col_idx, header))
        elif 'lng' in header_lower:
            categories['LNG Imports'].append((col_idx, header))
        elif 'total' in header_lower and 'supply' in header_lower:
            categories['Total Supply'].append((col_idx, header))
        else:
            categories['Storage/Other'].append((col_idx, header))
    
    for category, cols in categories.items():
        if cols:
            print(f"\n   {category}:")
            for col_idx, header in cols:
                print(f"     Column {col_idx:2d}: {header}")
    
    # Show sample values for key categories
    print(f"\n📈 SAMPLE VALUES (2016-10-04):")
    target_row = 15  # We know this is our target row from demand analysis
    
    for category, cols in categories.items():
        if cols:
            print(f"\n   {category}:")
            for col_idx, header in cols[:3]:  # First 3 columns per category
                val = data.iloc[target_row, col_idx]
                if pd.notna(val) and isinstance(val, (int, float)):
                    print(f"     {header}: {val:.2f}")
    
    return supply_columns, data

if __name__ == "__main__":
    supply_cols, sheet_data = print_supply_head_clean()