# -*- coding: utf-8 -*-
"""
Read and analyze Excel file sheets
"""

import pandas as pd

def read_excel_sheets():
    """Read and analyze Excel sheets"""
    
    print(f"\n{'='*80}")
    print("📊 READING EXCEL FILE SHEETS")
    print(f"{'='*80}")
    
    # Read use4.xlsx
    try:
        print(f"📁 Reading use4.xlsx...")
        
        # Get all sheet names
        excel_file = pd.ExcelFile('use4.xlsx')
        sheet_names = excel_file.sheet_names
        
        print(f"✅ Found {len(sheet_names)} sheets in use4.xlsx:")
        for i, sheet in enumerate(sheet_names, 1):
            print(f"   {i}. {sheet}")
        
        # Read TickerList sheet specifically
        if 'TickerList' in sheet_names:
            print(f"\n📋 READING TICKERLIST SHEET:")
            ticker_sheet = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
            print(f"   Shape: {ticker_sheet.shape}")
            print(f"   Columns: {list(ticker_sheet.columns)}")
            
            # Show sample data
            print(f"\n   First 3 rows:")
            for idx, row in ticker_sheet.head(3).iterrows():
                ticker = row.get('Ticker', 'N/A')
                desc = str(row.get('Description', 'N/A'))[:50]
                category = row.get('Category', 'N/A')
                norm_factor = row.get('Normalization factor', 'N/A')
                print(f"      {ticker}: {desc}... ({category}, norm: {norm_factor})")
        
    except Exception as e:
        print(f"❌ Error reading use4.xlsx: {e}")
    
    # Read template file if it exists
    try:
        print(f"\n📁 Reading template file...")
        template_file = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'
        
        template_excel = pd.ExcelFile(template_file)
        template_sheets = template_excel.sheet_names
        
        print(f"✅ Found {len(template_sheets)} sheets in template file:")
        for i, sheet in enumerate(template_sheets, 1):
            print(f"   {i}. {sheet}")
            
    except Exception as e:
        print(f"⚠️  Could not read template file: {e}")
    
    # Check output file if it exists
    try:
        print(f"\n📁 Checking output file...")
        output_file = 'DNB Markets EUROPEAN GAS BALANCE_temp.xlsx'
        
        output_excel = pd.ExcelFile(output_file)
        output_sheets = output_excel.sheet_names
        
        print(f"✅ Found {len(output_sheets)} sheets in output file:")
        for i, sheet in enumerate(output_sheets, 1):
            print(f"   {i}. {sheet}")
            
        # Read a sample from Countries sheet if it exists
        if 'Countries' in output_sheets:
            print(f"\n📊 SAMPLE FROM COUNTRIES SHEET:")
            countries_data = pd.read_excel(output_file, sheet_name='Countries', index_col=0)
            print(f"   Shape: {countries_data.shape}")
            print(f"   Columns: {list(countries_data.columns)}")
            
            # Show sample values
            if countries_data.shape[0] > 0:
                print(f"\n   First row values:")
                first_row = countries_data.iloc[0]
                for col in countries_data.columns[:10]:  # Show first 10 columns
                    value = first_row[col] if not pd.isna(first_row[col]) else 'NaN'
                    print(f"      {col}: {value}")
                    
    except Exception as e:
        print(f"⚠️  Could not read output file: {e}")

if __name__ == "__main__":
    read_excel_sheets()