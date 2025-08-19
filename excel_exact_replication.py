#!/usr/bin/env python3
"""
Excel Exact Replication - Fix Supply Calculation
===============================================

Now I have Excel's exact SUMIFS formula structure:
=SUMIFS(MultiTicker!$C111:$ZZ111, MultiTicker!$C$14:$ZZ$14, criteria1, 
        MultiTicker!$C$15:$ZZ$15, criteria2, MultiTicker!$C$16:$ZZ$16, criteria3)

Key insights:
- Excel uses columns C to ZZ (not full range)
- For 2017-01-01, uses row 111 in MultiTicker  
- Criteria in rows 14, 15, 16
- Range excludes columns A and B (date/metadata)
"""

import pandas as pd
import numpy as np
from pathlib import Path

def excel_exact_supply_replication():
    """Replicate Excel's exact SUMIFS logic for supply calculation."""
    
    print("🎯 EXCEL EXACT SUPPLY REPLICATION")
    print("=" * 80)
    
    # Load the Excel file
    excel_file = "2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx"
    
    if not Path(excel_file).exists():
        print(f"❌ Error: {excel_file} not found")
        return
    
    # Load MultiTicker data
    print("Loading MultiTicker data...")
    multiticker_df = pd.read_excel(excel_file, sheet_name='MultiTicker', header=None)
    
    # Load LiveSheet data for criteria
    print("Loading LiveSheet criteria...")
    livesheet_df = pd.read_excel(excel_file, sheet_name='Daily historic data by category', header=None)
    
    # Excel's exact approach:
    # 1. Data row: 111 (for 2017-01-01)
    # 2. Column range: C to ZZ (columns 2 to 701 in 0-indexed, but ZZ is column 701)
    # 3. Criteria rows: 14, 15, 16 (0-indexed: 13, 14, 15)
    
    data_row = 111 - 1  # 0-indexed: 110
    criteria_rows = [13, 14, 15]  # 0-indexed
    
    # Column range: C to ZZ
    # C = column 2, ZZ = column 701
    start_col = 2  # Column C
    # But MultiTicker only has 458 columns, so use actual end
    end_col = multiticker_df.shape[1] - 1
    
    print(f"Data row: {data_row + 1} (Excel row 111)")
    print(f"Column range: C to column {end_col} (MultiTicker actual range)")
    print(f"Criteria rows: 14, 15, 16 (Excel rows)")
    
    # Extract the criteria for supply routes (columns R-AI = 17-34 in 0-indexed)
    supply_criteria = {}
    supply_column_names = [
        'Slovakia_Austria', 'Russia_NordStream', 'Norway_Europe', 'Netherlands_Production',
        'GB_Production', 'LNG_Total', 'Algeria_Italy', 'Libya_Italy', 'Spain_France',
        'Denmark_Germany', 'Czech_Poland_Germany', 'Austria_Hungary_Export',
        'Slovenia_Austria', 'MAB_Austria', 'TAP_Italy', 'Austria_Production',
        'Italy_Production', 'Germany_Production'
    ]
    
    for i, route_name in enumerate(supply_column_names):
        livesheet_col = 17 + i  # Column R = 17
        criteria1 = livesheet_df.iloc[9, livesheet_col]   # Row 10 = index 9
        criteria2 = livesheet_df.iloc[10, livesheet_col]  # Row 11 = index 10  
        criteria3 = livesheet_df.iloc[11, livesheet_col]  # Row 12 = index 11
        
        supply_criteria[route_name] = {
            'criteria1': criteria1,  # Level 1: Import/Production/Export
            'criteria2': criteria2,  # Level 2: Region/Country
            'criteria3': criteria3   # Level 3: Destination/Category
        }
    
    print(f"\\nExtracted criteria for {len(supply_criteria)} supply routes")
    
    # Now replicate Excel's SUMIFS for each route
    route_results = {}
    total_supply = 0.0
    
    for route_name, criteria in supply_criteria.items():
        print(f"\\nProcessing {route_name}:")
        print(f"  Criteria: {criteria['criteria1']} | {criteria['criteria2']} | {criteria['criteria3']}")
        
        # Apply Excel's SUMIFS logic
        route_total = 0.0
        matches_found = 0
        
        for col_idx in range(start_col, end_col + 1):
            # Get criteria values from MultiTicker header rows
            header1 = multiticker_df.iloc[criteria_rows[0], col_idx] if col_idx < multiticker_df.shape[1] else None
            header2 = multiticker_df.iloc[criteria_rows[1], col_idx] if col_idx < multiticker_df.shape[1] else None
            header3 = multiticker_df.iloc[criteria_rows[2], col_idx] if col_idx < multiticker_df.shape[1] else None
            
            # Check if this column matches all three criteria
            match1 = str(header1).strip() == str(criteria['criteria1']).strip() if pd.notna(header1) else False
            match2 = str(header2).strip() == str(criteria['criteria2']).strip() if pd.notna(header2) else False
            
            # Handle wildcard for criteria3 (LNG case)
            if str(criteria['criteria3']).strip() == '*':
                match3 = True
            else:
                match3 = str(header3).strip() == str(criteria['criteria3']).strip() if pd.notna(header3) else False
            
            if match1 and match2 and match3:
                # Get the data value
                data_value = multiticker_df.iloc[data_row, col_idx] if pd.notna(multiticker_df.iloc[data_row, col_idx]) else 0.0
                
                # DON'T apply scaling factor - raw values are already in correct units
                # scaling_factor = multiticker_df.iloc[12, col_idx] if pd.notna(multiticker_df.iloc[12, col_idx]) else 1.0
                
                scaled_value = float(data_value)  # Use raw value directly
                route_total += scaled_value
                matches_found += 1
                
                if matches_found <= 3:  # Show first few matches
                    print(f"    Match col {col_idx}: {header1}|{header2}|{header3} = {data_value} (no scaling)")
        
        route_results[route_name] = route_total
        total_supply += route_total
        print(f"  Route total: {route_total:.2f} ({matches_found} columns)")
    
    # Display results
    print("\\n" + "=" * 80)
    print("EXCEL REPLICATION RESULTS")
    print("=" * 80)
    
    # Get LiveSheet reference for comparison
    livesheet_values = {}
    livesheet_total = 0.0
    
    for i, route_name in enumerate(supply_column_names):
        livesheet_col = 17 + i
        value = livesheet_df.iloc[104, livesheet_col]  # Row 105 = index 104
        livesheet_values[route_name] = float(value) if pd.notna(value) else 0.0
        livesheet_total += livesheet_values[route_name]
    
    print(f"{'Route':<25} {'LiveSheet':<12} {'MyCalc':<12} {'Diff':<10} {'Status':<8}")
    print("-" * 70)
    
    for route_name in supply_column_names:
        live_val = livesheet_values[route_name]
        my_val = route_results[route_name]
        diff = my_val - live_val
        status = "✅" if abs(diff) < 0.1 else "❌"
        
        print(f"{route_name:<25} {live_val:<12.2f} {my_val:<12.2f} {diff:<10.2f} {status:<8}")
    
    print("-" * 70)
    print(f"{'TOTAL':<25} {livesheet_total:<12.2f} {total_supply:<12.2f} {total_supply-livesheet_total:<10.2f}")
    
    # Success metrics
    accuracy = (1 - abs(total_supply - livesheet_total) / livesheet_total) * 100
    print(f"\\n🎯 ACCURACY: {accuracy:.1f}%")
    
    if accuracy > 99.0:
        print("✅ SUCCESS: Excel replication working perfectly!")
    elif accuracy > 95.0:
        print("🟡 GOOD: Close match, minor adjustments needed")
    else:
        print("❌ ISSUES: Major discrepancies need investigation")
    
    return route_results, total_supply

def main():
    """Execute the exact Excel replication."""
    
    print("Starting Excel exact supply replication...")
    
    try:
        results, total = excel_exact_supply_replication()
        
        print("\\n🎯 REPLICATION COMPLETE")
        print(f"Final total: {total:.2f}")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()