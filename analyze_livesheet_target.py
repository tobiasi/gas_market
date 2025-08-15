# -*- coding: utf-8 -*-
"""
Analyze the LiveSheet target file - specifically the Daily historic data tab
"""

import pandas as pd
import numpy as np

def analyze_livesheet_target():
    """Analyze the target LiveSheet file"""
    
    print(f"\n{'='*80}")
    print("🎯 ANALYZING LIVESHEET TARGET FILE")
    print(f"{'='*80}")
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    
    try:
        # Get all sheet names
        excel_file = pd.ExcelFile(filename)
        print(f"📁 Target file: {filename}")
        print(f"📋 Available sheets:")
        for i, sheet in enumerate(excel_file.sheet_names, 1):
            print(f"  {i:2d}. {sheet}")
        
        # Look for the Daily historic data sheet
        target_sheet = None
        for sheet in excel_file.sheet_names:
            if 'daily' in sheet.lower() and 'historic' in sheet.lower():
                target_sheet = sheet
                break
        
        if not target_sheet:
            # Look for similar names
            for sheet in excel_file.sheet_names:
                if any(keyword in sheet.lower() for keyword in ['daily', 'historic', 'category', 'demand']):
                    target_sheet = sheet
                    print(f"📋 Found similar sheet: {sheet}")
                    break
        
        if not target_sheet:
            print("❌ Could not find 'Daily historic data by category' sheet")
            print("Available sheets for manual selection:")
            for i, sheet in enumerate(excel_file.sheet_names, 1):
                print(f"  {i:2d}. {sheet}")
            return None
        
        print(f"\n📊 Reading target sheet: '{target_sheet}'")
        
        # Read the target data without headers first to understand structure
        target_data = pd.read_excel(filename, sheet_name=target_sheet, header=None)
        print(f"✅ Target data shape: {target_data.shape}")
        
        # Analyze structure - look for headers
        print(f"\n🔍 TARGET STRUCTURE ANALYSIS:")
        print("First 5 rows, first 15 columns:")
        for i in range(min(5, target_data.shape[0])):
            row_data = []
            for j in range(min(15, target_data.shape[1])):
                val = target_data.iloc[i, j]
                if pd.isna(val):
                    row_data.append('NaN')
                elif isinstance(val, (int, float)):
                    row_data.append(f'{val:.2f}')
                else:
                    row_data.append(str(val)[:12])
            print(f"Row {i}: {row_data}")
        
        # Find header row (look for countries)
        header_row = None
        countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany']
        
        for i in range(min(10, target_data.shape[0])):
            row_text = ' '.join([str(target_data.iloc[i, j]) for j in range(min(20, target_data.shape[1]))])
            if any(country in row_text for country in countries):
                header_row = i
                print(f"\n✅ Found header row at: {i}")
                break
        
        if header_row is not None:
            # Extract headers
            headers = []
            for j in range(target_data.shape[1]):
                val = target_data.iloc[header_row, j]
                headers.append(str(val) if pd.notna(val) else '')
            
            print(f"📋 Headers found: {headers[:15]}")
            
            # Find key columns
            key_columns = {}
            for j, header in enumerate(headers):
                if header in ['Italy', 'Total', 'Industrial', 'LDZ', 'Gas-to-Power']:
                    key_columns[header] = j
                    print(f"   {header}: column {j}")
            
            # Read actual data starting after header
            data_start_row = header_row + 1
            if data_start_row < target_data.shape[0]:
                print(f"\n🎯 TARGET VALUES (first 5 data rows):")
                for i in range(data_start_row, min(data_start_row + 5, target_data.shape[0])):
                    row_values = {}
                    for key, col_idx in key_columns.items():
                        if col_idx < target_data.shape[1]:
                            val = target_data.iloc[i, col_idx]
                            row_values[key] = val
                    
                    date_val = target_data.iloc[i, 0] if target_data.shape[1] > 0 else 'No date'
                    print(f"   {date_val}: {row_values}")
                
                # Save target values for comparison
                target_first_row = {}
                if key_columns:
                    for key, col_idx in key_columns.items():
                        if col_idx < target_data.shape[1]:
                            val = target_data.iloc[data_start_row, col_idx]
                            if pd.notna(val) and isinstance(val, (int, float)):
                                target_first_row[key] = float(val)
                
                print(f"\n🎯 TARGET FIRST ROW VALUES:")
                for key, val in target_first_row.items():
                    print(f"   {key}: {val:.2f}")
                
                # Save this for our processing
                with open('target_values.txt', 'w') as f:
                    f.write("TARGET VALUES FOR REPLICATION\n")
                    f.write("=" * 40 + "\n")
                    for key, val in target_first_row.items():
                        f.write(f"{key}: {val:.2f}\n")
                
                print(f"✅ Saved target values to: target_values.txt")
                
                return target_data, header_row, key_columns, target_first_row
        
        print("❌ Could not identify header structure")
        return None
        
    except Exception as e:
        print(f"❌ Error reading LiveSheet file: {e}")
        return None

def find_raw_data_file():
    """Find the raw Bloomberg data file"""
    
    print(f"\n{'='*80}")
    print("🔍 LOOKING FOR RAW BLOOMBERG DATA")
    print(f"{'='*80}")
    
    import os
    
    # Look for files with 'raw' in the name
    raw_files = []
    for file in os.listdir('.'):
        if 'raw' in file.lower() and (file.endswith('.csv') or file.endswith('.xlsx')):
            raw_files.append(file)
    
    if raw_files:
        print(f"📁 Found potential raw data files:")
        for i, file in enumerate(raw_files, 1):
            size = os.path.getsize(file) / 1024 / 1024  # MB
            print(f"  {i}. {file} ({size:.1f} MB)")
        
        # Use the largest one (likely the most complete)
        largest_file = max(raw_files, key=lambda f: os.path.getsize(f))
        print(f"✅ Using largest file: {largest_file}")
        return largest_file
    else:
        print("❌ No raw data files found")
        print("Expected files with 'raw' in the name (.csv or .xlsx)")
        return None

if __name__ == "__main__":
    # Analyze the target
    result = analyze_livesheet_target()
    
    # Find raw data
    raw_file = find_raw_data_file()
    
    if result and raw_file:
        print(f"\n✅ READY FOR PROCESSING:")
        print(f"   Target analyzed: {result[0].shape}")
        print(f"   Raw data file: {raw_file}")
        print(f"   Next: Process raw data to match target")
    else:
        print(f"\n❌ SETUP INCOMPLETE:")
        if not result:
            print("   - Could not analyze target LiveSheet")
        if not raw_file:
            print("   - Could not find raw Bloomberg data file")