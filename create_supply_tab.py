# -*- coding: utf-8 -*-
"""
Create supply tab replicating columns 17-38 from Daily historic data by category
"""

import pandas as pd
import numpy as np

def create_supply_tab():
    """Create supply tab with columns 17-38"""
    
    print("⛽ CREATING SUPPLY TAB")
    print("=" * 80)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    # Read the main sheet
    print("📊 Reading original LiveSheet...")
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    # Extract supply columns 17-38
    supply_start_col = 17
    supply_end_col = 38  # Inclusive
    
    print(f"📋 Extracting supply columns {supply_start_col}-{supply_end_col}...")
    
    # Build supply column mapping
    supply_columns = {}
    supply_headers = []
    
    for j in range(supply_start_col, supply_end_col + 1):
        row9 = str(target_data.iloc[9, j]) if pd.notna(target_data.iloc[9, j]) else ""
        row10 = str(target_data.iloc[10, j]) if pd.notna(target_data.iloc[10, j]) else ""
        row11 = str(target_data.iloc[11, j]) if pd.notna(target_data.iloc[11, j]) else ""
        
        # Clean up
        row9_clean = row9.replace('nan', '').strip()
        row10_clean = row10.replace('nan', '').strip()
        row11_clean = row11.replace('nan', '').strip()
        
        # Create header name - use the most descriptive available
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
        supply_headers.append(header_name)
    
    print(f"✅ Identified {len(supply_columns)} supply columns")
    
    # Show the supply columns we're extracting
    print(f"\n📋 SUPPLY COLUMNS TO EXTRACT:")
    for j, info in supply_columns.items():
        print(f"   Col {j:2d}: {info['header']} ({info['category']})")
    
    # Extract data starting from row 12
    data_start_row = 12
    print(f"\n📊 Extracting data starting from row {data_start_row}...")
    
    # Get dates and data
    dates = []
    supply_data = []
    
    for i in range(data_start_row, target_data.shape[0]):
        date_val = target_data.iloc[i, 1]  # Date in column 1
        
        # Handle date conversion more carefully
        if pd.notna(date_val):
            # Convert to string first to handle various date formats
            if hasattr(date_val, 'date'):  # datetime object
                date_str = date_val.date()
            elif isinstance(date_val, str):
                date_str = date_val
            else:
                date_str = str(date_val)
            
            # Only process if it looks like a valid date
            try:
                date_parsed = pd.to_datetime(date_str)
                dates.append(date_parsed)
                
                # Extract supply values for this row
                row_data = {}
                for j in range(supply_start_col, supply_end_col + 1):
                    header = supply_columns[j]['header']
                    val = target_data.iloc[i, j]
                    
                    if pd.notna(val) and isinstance(val, (int, float)):
                        row_data[header] = float(val)
                    else:
                        row_data[header] = np.nan
                
                supply_data.append(row_data)
                
            except (ValueError, TypeError):
                # Skip invalid dates
                continue
    
    # Create supply DataFrame
    supply_df = pd.DataFrame(supply_data, index=dates)
    
    print(f"✅ Extracted {len(supply_df)} rows of supply data")
    print(f"📅 Date range: {supply_df.index[0]} to {supply_df.index[-1]}")
    
    # Verification - show first few rows
    print(f"\n📊 SUPPLY DATA HEAD (first 5 rows):")
    print(f"{'Date':<12}", end="")
    for header in supply_headers[:8]:  # First 8 columns
        print(f"{header[:12]:<15}", end="")
    print()
    print("-" * (12 + 15 * 8))
    
    for i in range(min(5, len(supply_df))):
        date_str = str(supply_df.index[i])[:10]
        print(f"{date_str:<12}", end="")
        
        row_data = supply_df.iloc[i]
        for header in supply_headers[:8]:
            val = row_data[header]
            if pd.notna(val):
                print(f"{val:<15.2f}", end="")
            else:
                print(f"{'NaN':<15}", end="")
        print()
    
    # Show key supply totals
    print(f"\n🎯 KEY SUPPLY VALUES (2016-10-04):")
    target_date = '2016-10-04'
    if target_date in [str(d)[:10] for d in supply_df.index]:
        target_row = supply_df[supply_df.index.strftime('%Y-%m-%d') == target_date].iloc[0]
        
        key_supplies = ['Norway', 'Russia (Nord Stream)', 'LNG', 'Algeria', 'Netherlands', 'GB', 'Total']
        for key in key_supplies:
            if key in target_row.index and pd.notna(target_row[key]):
                print(f"   {key}: {target_row[key]:.2f}")
    
    # Save supply tab
    output_file = 'European_Gas_Supply_Data.xlsx'
    
    # Create final DataFrame with Date as first column
    final_supply_df = supply_df.copy()
    final_supply_df.reset_index(inplace=True)
    final_supply_df.rename(columns={'index': 'Date'}, inplace=True)
    
    # Reorder columns logically
    column_order = ['Date']
    
    # Group by category
    imports = []
    production = []
    exports = []
    totals = []
    others = []
    
    for j, info in supply_columns.items():
        header = info['header']
        category = info['category'].lower()
        
        if 'import' in category:
            imports.append(header)
        elif 'production' in category:
            production.append(header)
        elif 'export' in category:
            exports.append(header)
        elif 'supply' in category or 'total' in header.lower():
            totals.append(header)
        else:
            others.append(header)
    
    # Organize columns logically: Imports, Production, Exports, Totals, Others
    column_order.extend(imports)
    column_order.extend(production)
    column_order.extend(exports)
    column_order.extend(totals)
    column_order.extend(others)
    
    # Ensure all columns are included
    for col in final_supply_df.columns:
        if col not in column_order:
            column_order.append(col)
    
    final_supply_df = final_supply_df[column_order]
    
    # Save to Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_supply_df.to_excel(writer, sheet_name='Supply', index=False)
    
    print(f"\n✅ SUPPLY TAB CREATED: {output_file}")
    print(f"📊 Shape: {final_supply_df.shape}")
    
    # Also save as CSV
    csv_file = 'European_Gas_Supply_Data.csv'
    final_supply_df.to_csv(csv_file, index=False)
    print(f"✅ Also saved as CSV: {csv_file}")
    
    # Summary
    print(f"\n📈 SUPPLY TAB SUMMARY:")
    print(f"   Total columns: {len(supply_columns)} (columns 17-38)")
    print(f"   Imports: {len(imports)} columns")
    print(f"   Production: {len(production)} columns") 
    print(f"   Exports: {len(exports)} columns")
    print(f"   Totals: {len(totals)} columns")
    print(f"   Other: {len(others)} columns")
    print(f"   Date range: {len(final_supply_df)} rows")
    
    return final_supply_df

if __name__ == "__main__":
    supply_tab = create_supply_tab()