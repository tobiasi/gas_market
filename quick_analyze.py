#!/usr/bin/env python3

import sys
sys.path.append("C:/development/commodities")

try:
    import pandas as pd
    print("Pandas imported successfully")
    
    # Try to read use4.xlsx
    print("Reading use4.xlsx...")
    
    # First, try to see all sheets
    try:
        xl = pd.ExcelFile('use4.xlsx')
        print(f"Available sheets: {xl.sheet_names}")
    except Exception as e:
        print(f"Could not read Excel file: {e}")
        sys.exit(1)
    
    # Try to read TickerList sheet as in original
    for sheet_name in ['TickerList', 'Ticker', 'ticker', 'tickers']:
        try:
            print(f"\nTrying sheet: {sheet_name}")
            
            # Try original approach first
            df = pd.read_excel('use4.xlsx', sheet_name=sheet_name, skiprows=8, usecols=range(1,9))
            print("SUCCESS with original parameters!")
            print(f"Shape: {df.shape}")
            print(f"Columns: {df.columns.tolist()}")
            print("\nFirst 3 rows:")
            print(df.head(3))
            
            # Check if 'Replace blanks with #N/A' column exists
            has_na_col = any('replace' in str(col).lower() and 'blank' in str(col).lower() for col in df.columns)
            print(f"\nHas 'Replace blanks' column: {has_na_col}")
            
            break
            
        except Exception as e:
            print(f"Failed: {e}")
            
            # Try without skiprows to see raw structure
            try:
                df_raw = pd.read_excel('use4.xlsx', sheet_name=sheet_name)
                print(f"Raw structure (first 10 rows):")
                for i in range(min(10, len(df_raw))):
                    print(f"Row {i}: {df_raw.iloc[i].tolist()}")
                break
            except:
                continue
    
except ImportError as e:
    print(f"Import error: {e}")
except Exception as e:
    print(f"Other error: {e}")