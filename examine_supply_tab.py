# -*- coding: utf-8 -*-
"""
Examine the supply tab in the LiveSheet file to understand its structure
"""

import pandas as pd

def examine_supply_tab():
    """Examine the supply tab structure"""
    
    print("⛽ EXAMINING SUPPLY TAB STRUCTURE")
    print("=" * 80)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    
    # First, let's see all available sheet names
    excel_file = pd.ExcelFile(filename)
    print("📋 ALL AVAILABLE SHEETS:")
    for i, sheet in enumerate(excel_file.sheet_names, 1):
        print(f"  {i:2d}. {sheet}")
    
    # Look for supply-related sheets
    supply_sheets = []
    for sheet in excel_file.sheet_names:
        if any(keyword in sheet.lower() for keyword in ['supply', 'production', 'import', 'export']):
            supply_sheets.append(sheet)
    
    print(f"\n⛽ POTENTIAL SUPPLY-RELATED SHEETS:")
    if supply_sheets:
        for i, sheet in enumerate(supply_sheets, 1):
            print(f"  {i}. {sheet}")
    else:
        print("  No obvious supply sheets found, checking for similar patterns...")
        
        # Look for sheets that might contain supply data
        for sheet in excel_file.sheet_names:
            if any(keyword in sheet.lower() for keyword in ['daily', 'historic', 'balance']):
                print(f"  Potential: {sheet}")
    
    # Let's examine the most likely supply sheet
    # Based on naming convention, it might be similar to demand sheet
    target_sheet = None
    
    # Try common supply sheet names
    possible_names = [
        'Supply', 
        'Daily historic supply by category',
        'Daily historic data by category - Supply',
        'Production',
        'Import/Export'
    ]
    
    for name in possible_names:
        if name in excel_file.sheet_names:
            target_sheet = name
            break
    
    # If not found, look for similar patterns
    if not target_sheet:
        for sheet in excel_file.sheet_names:
            if 'supply' in sheet.lower():
                target_sheet = sheet
                break
    
    # If still not found, let's examine a few sheets manually
    if not target_sheet:
        print(f"\n🔍 EXAMINING FIRST FEW SHEETS FOR SUPPLY DATA:")
        for i, sheet in enumerate(excel_file.sheet_names[:5]):
            print(f"\n📊 Sheet: {sheet}")
            try:
                data = pd.read_excel(filename, sheet_name=sheet, header=None, nrows=15)
                
                # Look for supply-related keywords in the data
                supply_keywords = ['supply', 'production', 'import', 'export', 'lng', 'pipeline']
                found_supply = False
                
                for row_idx in range(min(15, data.shape[0])):
                    for col_idx in range(min(20, data.shape[1])):
                        cell_val = str(data.iloc[row_idx, col_idx]).lower()
                        if any(keyword in cell_val for keyword in supply_keywords):
                            print(f"   Found supply keyword '{keyword}' at row {row_idx}, col {col_idx}: {cell_val}")
                            found_supply = True
                            if not target_sheet:
                                target_sheet = sheet
                
                if not found_supply:
                    print(f"   No obvious supply keywords found")
                    
            except Exception as e:
                print(f"   Error reading sheet: {e}")
    
    # Now examine the target sheet in detail
    if target_sheet:
        print(f"\n⛽ DETAILED EXAMINATION OF: {target_sheet}")
        print("=" * 80)
        
        try:
            data = pd.read_excel(filename, sheet_name=target_sheet, header=None)
            print(f"📊 Sheet shape: {data.shape}")
            
            # Look for header row
            print(f"\n📋 SEARCHING FOR HEADER ROW:")
            supply_keywords = ['supply', 'production', 'import', 'export', 'lng', 'pipeline', 'norway', 'algeria', 'russia']
            
            header_candidates = []
            for i in range(min(20, data.shape[0])):
                row_text = []
                keyword_count = 0
                
                for j in range(min(30, data.shape[1])):
                    val = data.iloc[i, j]
                    if pd.notna(val):
                        val_str = str(val)
                        row_text.append(val_str)
                        if any(keyword in val_str.lower() for keyword in supply_keywords):
                            keyword_count += 1
                
                if keyword_count >= 2:  # At least 2 supply-related keywords
                    header_candidates.append((i, keyword_count, row_text))
                    print(f"   Row {i:2d}: {keyword_count} supply keywords - {row_text[:10]}")
            
            if header_candidates:
                # Use the row with most supply keywords
                best_header = max(header_candidates, key=lambda x: x[1])
                header_row_idx = best_header[0]
                
                print(f"\n✅ BEST HEADER ROW: {header_row_idx}")
                print(f"Headers: {best_header[2][:15]}")
                
                # Extract data starting after header
                data_start_row = header_row_idx + 1
                print(f"\n📊 SUPPLY DATA HEAD (first 10 rows):")
                print(f"Row {header_row_idx} (Headers):")
                
                # Print headers
                headers = []
                for j in range(min(30, data.shape[1])):
                    val = data.iloc[header_row_idx, j]
                    header_str = str(val) if pd.notna(val) else f'Col_{j}'
                    headers.append(header_str)
                    if j < 15:  # Print first 15 headers
                        print(f"  Col {j:2d}: {header_str}")
                
                # Print data rows
                for i in range(data_start_row, min(data_start_row + 10, data.shape[0])):
                    date_val = data.iloc[i, 1] if data.shape[1] > 1 else data.iloc[i, 0]
                    print(f"\nRow {i} ({date_val}):")
                    
                    # Show first 10 columns with non-null values
                    values_shown = 0
                    for j in range(min(30, data.shape[1])):
                        val = data.iloc[i, j]
                        if pd.notna(val) and values_shown < 10:
                            header = headers[j] if j < len(headers) else f'Col_{j}'
                            if isinstance(val, (int, float)):
                                print(f"  {header}: {val:.2f}")
                            else:
                                print(f"  {header}: {val}")
                            values_shown += 1
                
                return target_sheet, header_row_idx, headers, data
                
            else:
                print(f"❌ No clear supply header row found")
                
        except Exception as e:
            print(f"❌ Error examining sheet: {e}")
    
    else:
        print(f"❌ No supply sheet identified")
    
    return None, None, None, None

if __name__ == "__main__":
    sheet_name, header_row, headers, data = examine_supply_tab()