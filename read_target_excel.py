# -*- coding: utf-8 -*-
"""
Read the actual "Daily historic data by category" tab from Excel
"""

import pandas as pd

def read_target_excel():
    """Read target Excel structure"""
    
    print(f"\n{'='*80}")
    print("📊 READING TARGET EXCEL STRUCTURE")
    print(f"{'='*80}")
    
    try:
        # Read the target Excel file
        excel_file = pd.ExcelFile('DNB Markets EUROPEAN GAS BALANCE.xlsx')
        
        print("Available sheets:")
        for i, sheet in enumerate(excel_file.sheet_names, 1):
            print(f"  {i}. {sheet}")
        
        # Read "Demand" sheet (this is likely our target)
        if 'Demand' in excel_file.sheet_names:
            print(f"\n📋 Reading 'Demand' sheet...")
            
            # Read without any processing first
            target_sheet = pd.read_excel('DNB Markets EUROPEAN GAS BALANCE.xlsx', 
                                       sheet_name='Demand',
                                       header=None)
            
            print(f"Shape: {target_sheet.shape}")
            
            # Show first few rows and columns to understand structure
            print(f"\n🔍 STRUCTURE ANALYSIS (first 5 rows, 20 columns):")
            for i in range(min(5, target_sheet.shape[0])):
                row_data = [str(val)[:15] if pd.notna(val) else 'NaN' for val in target_sheet.iloc[i, :20]]
                print(f"Row {i}: {row_data}")
            
            # Look for Italy specifically
            print(f"\n🇮🇹 SEARCHING FOR ITALY...")
            italy_positions = []
            for i in range(min(10, target_sheet.shape[0])):
                for j in range(target_sheet.shape[1]):
                    if str(target_sheet.iloc[i, j]).strip() == 'Italy':
                        italy_positions.append((i, j))
                        print(f"Found 'Italy' at row {i}, col {j}")
            
            # Check column structure around Italy
            if italy_positions:
                row, col = italy_positions[0]  # First Italy position
                print(f"\n📊 COLUMNS AROUND ITALY (row {row}):")
                start_col = max(0, col - 5)
                end_col = min(target_sheet.shape[1], col + 10)
                
                headers = [str(target_sheet.iloc[row, c])[:15] if pd.notna(target_sheet.iloc[row, c]) else 'NaN' 
                          for c in range(start_col, end_col)]
                print(f"Columns {start_col}-{end_col-1}: {headers}")
                
                # Check actual data values below Italy header
                if row + 1 < target_sheet.shape[0]:
                    print(f"\nData values below Italy header:")
                    for data_row in range(row + 1, min(row + 6, target_sheet.shape[0])):
                        italy_value = target_sheet.iloc[data_row, col]
                        if pd.notna(italy_value):
                            try:
                                italy_float = float(italy_value)
                                print(f"  Row {data_row}: {italy_float:.6f}")
                            except:
                                print(f"  Row {data_row}: {italy_value} (not numeric)")
            
            # Look for Total, Industrial, LDZ, Gas-to-Power
            key_terms = ['Total', 'Industrial', 'LDZ', 'Gas-to-Power']
            print(f"\n🔍 SEARCHING FOR KEY TERMS...")
            
            for term in key_terms:
                positions = []
                for i in range(min(10, target_sheet.shape[0])):
                    for j in range(target_sheet.shape[1]):
                        cell_val = str(target_sheet.iloc[i, j]).strip()
                        if cell_val == term:
                            positions.append((i, j))
                
                if positions:
                    print(f"{term}: Found at {positions}")
                    # Show some data values
                    row, col = positions[0]
                    if row + 1 < target_sheet.shape[0]:
                        for data_row in range(row + 1, min(row + 4, target_sheet.shape[0])):
                            value = target_sheet.iloc[data_row, col]
                            if pd.notna(value):
                                try:
                                    val_float = float(value)
                                    print(f"  {term} Row {data_row}: {val_float:.6f}")
                                except:
                                    print(f"  {term} Row {data_row}: {value}")
                else:
                    print(f"{term}: Not found")
        
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")

if __name__ == "__main__":
    read_target_excel()