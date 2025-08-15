import pandas as pd
import os

print("Analyzing use4.xlsx structure...")
print("="*50)

try:
    # Read the Excel file to see all sheet names
    xl_file = pd.ExcelFile('use4.xlsx')
    print(f"Sheet names: {xl_file.sheet_names}")
    
    # Try to read the TickerList sheet (similar to original)
    print("\nTrying to read TickerList sheet...")
    try:
        # Try the same parameters as original
        df = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8, usecols=range(1,9))
        print(f"Successfully read with original parameters")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print(f"\nFirst few rows:")
        print(df.head())
        
    except Exception as e:
        print(f"Failed with original parameters: {e}")
        
        # Try without skiprows and see the raw structure
        try:
            df_raw = pd.read_excel('use4.xlsx', sheet_name='TickerList')
            print(f"\nRaw sheet structure (first 15 rows):")
            print(df_raw.head(15))
            print(f"Columns: {list(df_raw.columns)}")
            
        except Exception as e2:
            print(f"Failed to read raw sheet: {e2}")
            
            # Try different sheet names if TickerList doesn't exist
            for sheet in xl_file.sheet_names:
                try:
                    df_sheet = pd.read_excel('use4.xlsx', sheet_name=sheet)
                    print(f"\nSheet '{sheet}':")
                    print(f"Shape: {df_sheet.shape}")
                    print(f"Columns: {list(df_sheet.columns)}")
                    if 'ticker' in str(df_sheet.columns).lower() or any('ticker' in str(col).lower() for col in df_sheet.columns):
                        print("*** This sheet might contain ticker data ***")
                        print(df_sheet.head())
                        break
                except:
                    continue
                    
except Exception as main_error:
    print(f"Error reading file: {main_error}")