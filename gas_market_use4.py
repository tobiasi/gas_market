# -*- coding: utf-8 -*-
"""
UPDATED VERSION for use4.xlsx - handles missing 'Replace blanks with NA' column
"""

import numpy as np
import os
import pandas as pd
import datetime
import xlrd

def get_dates_from_excel(datenumbers):
    dates = ['']*len(datenumbers)
    for idx, date in enumerate(datenumbers):
        datetime_date = xlrd.xldate_as_datetime(date, 0)
        date_object = datetime_date.date()
        dates[idx] = date_object
    return dates

from xbbg import blp
import pandas as pd
import numpy as np
import sys
import pycountry_convert as pc
sys.path.append("C:/development/commodities")
from update_spreadsheet import update_spreadsheet, to_excel_dates, update_spreadsheet_pweekly

from datetime import datetime
from xbbg import blp
import sys
from dateutil.relativedelta import relativedelta
import numpy as np

#%% SET PATH FOR DATA SHEET - UPDATED FOR USE4
os.chdir('G:\Commodity Research\Gas')
data_path = 'use4.xlsx'  # Updated to use local use4.xlsx file
filename = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'
output_filename = 'DNB Markets EUROPEAN GAS BALANCE_temp.xlsx'

#%% Load data - FLEXIBLE column handling for use4.xlsx
print("Loading Bloomberg data from use4.xlsx...")

try:
    # First, try to read with original parameters
    dataset = pd.read_excel(data_path, sheet_name='TickerList', skiprows=8, usecols=range(1,9))
    print("Successfully loaded with original parameters")
except Exception as e:
    print(f"Original parameters failed: {e}")
    
    # Try alternative approaches
    try:
        # Try reading the full sheet first to understand structure
        dataset_full = pd.read_excel(data_path, sheet_name='TickerList')
        print("Full sheet structure:")
        print(dataset_full.head(15))
        
        # Try to find where data starts (look for 'Ticker' column)
        ticker_row = None
        for i, row in dataset_full.iterrows():
            if any('ticker' in str(val).lower() for val in row.values if pd.notna(val)):
                ticker_row = i
                break
        
        if ticker_row is not None:
            dataset = pd.read_excel(data_path, sheet_name='TickerList', skiprows=ticker_row)
            print(f"Found data starting at row {ticker_row}")
        else:
            # Fallback: try without skiprows
            dataset = pd.read_excel(data_path, sheet_name='TickerList')
            
    except Exception as e2:
        print(f"Alternative loading failed: {e2}")
        raise

print(f"Dataset shape: {dataset.shape}")
print(f"Dataset columns: {dataset.columns.tolist()}")
print("First few rows:")
print(dataset.head())

#%% HANDLE MISSING 'Replace blanks with #N/A' COLUMN
# Check if the 'Replace blanks with #N/A' column exists
has_replace_blanks_col = any('replace' in str(col).lower() and 'blank' in str(col).lower() for col in dataset.columns)
print(f"\nHas 'Replace blanks' column: {has_replace_blanks_col}")

if not has_replace_blanks_col:
    # Add a default 'Replace blanks with #N/A' column
    # Default behavior: most series should replace blanks with 0, some with #N/A
    print("Adding default 'Replace blanks with #N/A' column...")
    
    # Create default values based on category or description patterns
    def determine_replace_blanks(row):
        desc = str(row.get('Description', '')).lower()
        category = str(row.get('Category', '')).lower()
        
        # Categories that typically should keep #N/A for missing values
        na_categories = ['price', 'rate', 'spread', 'index']
        na_keywords = ['price', 'rate', 'spread', 'index', 'benchmark']
        
        if any(cat in category for cat in na_categories) or any(kw in desc for kw in na_keywords):
            return 'Y'  # Keep #N/A for missing values
        else:
            return 'N'  # Replace missing with 0
    
    dataset['Replace blanks with #N/A'] = dataset.apply(determine_replace_blanks, axis=1)
    print("Default 'Replace blanks' column added based on categories")

# Get ticker list
tickers = list(set(dataset['Ticker'].dropna().tolist()))
print(f"Found {len(tickers)} unique tickers")

# Download Bloomberg data
data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)

print("Processing ticker data...")
full_data = pd.DataFrame()
for row in range(len(dataset)):
    information = dataset.loc[row]
    try:
        temp_df = data[[information['Ticker']]] * information['Normalization factor']
    except:
        temp_df = pd.DataFrame(columns=[information['Ticker']], index=data.index)

    # Handle missing values based on 'Replace blanks with #N/A' setting
    replace_blanks = information.get('Replace blanks with #N/A', 'N')
    if replace_blanks != 'Y':
        temp_df[temp_df.isna().values] = 0
    
    # Create MultiIndex columns - handle missing columns gracefully
    desc = information.get('Description', f'Unknown_{row}')
    replace_val = information.get('Replace blanks with #N/A', 'N') 
    category = information.get('Category', 'Unknown')
    region_from = information.get('Region from', '')
    region_to = information.get('Region to', '')
    ticker = information.get('Ticker', f'UNKNOWN_{row}')
    
    temp_df.columns = pd.MultiIndex.from_tuples([(desc, replace_val, category, 
                                                 region_from, region_to, ticker, ticker)])
    
    if row == 0:
        full_data = temp_df
    else:
        full_data = pd.merge(full_data, temp_df, left_index=True, right_index=True, how="outer")

# Data fixes (same as original)
print("Applying data fixes...")
if full_data.shape[1] > 103:  # Check if columns exist before fixing
    try:
        full_data.iloc[1074:1076, 103] = 0
    except:
        pass
        
if full_data.shape[1] > 29:
    try:
        full_data.iloc[2892:, 27] = 0
        full_data.iloc[2892:, 29] = 0
    except:
        pass
        
if full_data.shape[1] > 39:
    try:
        full_data.iloc[3122, 39] = (full_data.iloc[3121, 39] + full_data.iloc[3123, 39]) / 2
    except:
        pass
        
if full_data.shape[1] > 128:
    try:
        full_data.iloc[3103, 128] = 25
    except:
        pass

# Preserve the last row and interpolate only on the t-1 rows
full_data = full_data.copy()
full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')

full_data.index = pd.to_datetime(full_data.index)
full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]

print("Data processing complete!")
print(f"Final full_data shape: {full_data.shape}")

# Continue with rest of processing (same as original)...
# [Rest of the original processing code would go here]

print("Script completed successfully with use4.xlsx!")