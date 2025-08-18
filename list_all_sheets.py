# -*- coding: utf-8 -*-
"""
List all sheets in the LiveSheet file with details
"""

import pandas as pd

def list_all_sheets():
    """List all sheets in the LiveSheet file"""
    
    print("📋 ALL SHEETS IN LIVESHEET FILE")
    print("=" * 80)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    
    try:
        excel_file = pd.ExcelFile(filename)
        
        print(f"📁 File: {filename}")
        print(f"📊 Total sheets: {len(excel_file.sheet_names)}")
        print()
        
        for i, sheet_name in enumerate(excel_file.sheet_names, 1):
            print(f"{i:2d}. {sheet_name}")
            
            # Try to get basic info about each sheet
            try:
                sheet_data = pd.read_excel(filename, sheet_name=sheet_name, header=None, nrows=5)
                shape = sheet_data.shape
                print(f"    Shape: {shape[0]}+ rows × {shape[1]} columns")
                
                # Look for meaningful content in first few cells
                content_sample = []
                for row in range(min(3, shape[0])):
                    for col in range(min(5, shape[1])):
                        val = sheet_data.iloc[row, col]
                        if pd.notna(val) and str(val) != 'nan' and len(str(val)) > 1:
                            content_sample.append(str(val)[:20])
                            if len(content_sample) >= 3:
                                break
                    if len(content_sample) >= 3:
                        break
                
                if content_sample:
                    print(f"    Sample content: {', '.join(content_sample)}")
                else:
                    print(f"    Content: [Empty or minimal data]")
                    
            except Exception as e:
                print(f"    Error reading sheet: {str(e)[:50]}")
            
            print()
    
    except Exception as e:
        print(f"❌ Error opening file: {e}")

if __name__ == "__main__":
    list_all_sheets()