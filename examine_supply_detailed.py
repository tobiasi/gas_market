# -*- coding: utf-8 -*-
"""
Examine supply-related sheets in detail, especially LNG imports
"""

import pandas as pd

def examine_supply_detailed():
    """Examine supply-related sheets in detail"""
    
    print("⛽ DETAILED SUPPLY SHEET EXAMINATION")
    print("=" * 80)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    
    # Check the LNG imports sheet first
    print("🚢 EXAMINING: LNG imports by country")
    print("-" * 50)
    
    try:
        lng_data = pd.read_excel(filename, sheet_name='LNG imports by country', header=None)
        print(f"📊 LNG Sheet shape: {lng_data.shape}")
        
        # Look for header structure
        print(f"\n📋 LNG SHEET STRUCTURE (first 15 rows, first 15 cols):")
        for i in range(min(15, lng_data.shape[0])):
            print(f"\nRow {i:2d}:", end="")
            for j in range(min(15, lng_data.shape[1])):
                val = lng_data.iloc[i, j]
                if pd.notna(val):
                    val_str = str(val)[:12]  # Truncate long values
                    print(f" {val_str:<12}", end="")
                else:
                    print(f" {'NaN':<12}", end="")
            print()
        
    except Exception as e:
        print(f"Error reading LNG sheet: {e}")
    
    # Now check if there's supply data in the main data sheet
    print(f"\n📊 CHECKING 'Daily historic data by category' FOR SUPPLY DATA:")
    print("-" * 60)
    
    try:
        main_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
        print(f"📊 Main sheet shape: {main_data.shape}")
        
        # Look beyond the demand columns we already know about
        print(f"\n🔍 SCANNING FOR SUPPLY KEYWORDS IN MAIN SHEET:")
        supply_keywords = ['supply', 'production', 'import', 'export', 'lng', 'pipeline', 'norway', 'algeria', 'russia', 'nord']
        
        # Check headers and sample data
        for i in range(min(20, main_data.shape[0])):
            for j in range(main_data.shape[1]):
                val = main_data.iloc[i, j]
                if pd.notna(val):
                    val_str = str(val).lower()
                    for keyword in supply_keywords:
                        if keyword in val_str:
                            print(f"   Found '{keyword}' at Row {i}, Col {j}: {str(val)[:50]}")
        
        # Look specifically at columns beyond what we used for demand (beyond column 30)
        if main_data.shape[1] > 30:
            print(f"\n📊 COLUMNS BEYOND DEMAND DATA (30+):")
            print("Header row 10:")
            for j in range(30, min(main_data.shape[1], 50)):
                val = main_data.iloc[10, j]  # Header row we know is 10
                if pd.notna(val):
                    print(f"   Col {j:2d}: {val}")
            
            # Sample data from these columns
            print(f"\nSample data row 12:")
            for j in range(30, min(main_data.shape[1], 50)):
                val = main_data.iloc[12, j]  # Data row we know starts at 12
                if pd.notna(val) and isinstance(val, (int, float)) and abs(val) > 1:
                    header = main_data.iloc[10, j] if pd.notna(main_data.iloc[10, j]) else f'Col_{j}'
                    print(f"   {header}: {val:.2f}")
    
    except Exception as e:
        print(f"Error examining main sheet: {e}")
    
    # Check MultiTicker sheet for supply data
    print(f"\n📊 CHECKING 'MultiTicker' SHEET:")
    print("-" * 40)
    
    try:
        multi_data = pd.read_excel(filename, sheet_name='MultiTicker', header=None)
        print(f"📊 MultiTicker shape: {multi_data.shape}")
        
        # Check if this contains supply data
        print(f"\nFirst few rows and columns:")
        for i in range(min(5, multi_data.shape[0])):
            print(f"Row {i}:", end="")
            for j in range(min(10, multi_data.shape[1])):
                val = multi_data.iloc[i, j]
                if pd.notna(val):
                    val_str = str(val)[:15]
                    print(f" {val_str:<15}", end="")
            print()
    
    except Exception as e:
        print(f"Error examining MultiTicker: {e}")

def print_supply_head_if_found():
    """Print the head of supply data if we can identify it"""
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    
    # Based on our examination, let's check if supply data is in the main sheet
    try:
        main_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
        
        # We know demand data ends around column 28, let's check what's beyond
        print(f"\n⛽ POTENTIAL SUPPLY DATA HEAD:")
        print("=" * 80)
        
        # Print headers from column 30 onwards
        print(f"📋 Headers (from column 30+):")
        supply_headers = {}
        for j in range(30, min(main_data.shape[1], 60)):
            val = main_data.iloc[10, j]  # Header row
            if pd.notna(val) and str(val) != 'nan':
                supply_headers[j] = str(val)
                print(f"   Col {j:2d}: {val}")
        
        # Print sample data
        if supply_headers:
            print(f"\n📊 SUPPLY DATA HEAD (first 10 rows):")
            print(f"{'Date':<12}", end="")
            for col_idx in list(supply_headers.keys())[:10]:  # First 10 supply columns
                header = supply_headers[col_idx]
                print(f"{header[:12]:<15}", end="")
            print()
            print("-" * (12 + 15 * min(10, len(supply_headers))))
            
            # Data rows
            for i in range(12, min(22, main_data.shape[0])):  # 10 data rows
                date_val = main_data.iloc[i, 1]
                date_str = str(date_val)[:10] if pd.notna(date_val) else "N/A"
                print(f"{date_str:<12}", end="")
                
                for col_idx in list(supply_headers.keys())[:10]:
                    val = main_data.iloc[i, col_idx]
                    if pd.notna(val) and isinstance(val, (int, float)):
                        print(f"{val:<15.2f}", end="")
                    else:
                        print(f"{'NaN':<15}", end="")
                print()
        else:
            print("❌ No clear supply data headers found beyond column 30")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    examine_supply_detailed()
    print_supply_head_if_found()