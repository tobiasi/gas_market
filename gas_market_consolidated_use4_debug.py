# -*- coding: utf-8 -*-
"""
DEBUG VERSION: gas_market_consolidated_use4.py with Italy tracing
"""

import warnings
import pandas as pd
# Suppress pandas warnings for cleaner output
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)
warnings.filterwarnings('ignore', category=UserWarning)
pd.options.mode.chained_assignment = None  # Suppress SettingWithCopyWarning

import numpy as np
import os
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
data_path = 'use4.xlsx'  # Updated to use4.xlsx
filename = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'
output_filename = 'DNB Markets EUROPEAN GAS BALANCE_temp.xlsx'

#%% Load data, get demand components - ADAPTED FOR USE4 STRUCTURE
print("Loading Bloomberg data from use4.xlsx...")

# Read use4.xlsx with new column structure (based on image analysis)
dataset = pd.read_excel(data_path, sheet_name='TickerList', skiprows=8)

# Handle the different column structure in use4.xlsx
print(f"Dataset shape: {dataset.shape}")
print(f"Available columns: {dataset.columns.tolist()}")

# Add missing 'Replace blanks with #N/A' column with smart defaults
if 'Replace blanks with #N/A' not in dataset.columns:
    print("Adding 'Replace blanks with #N/A' column with smart defaults...")
    
    def smart_na_handling(row):
        """Determine NA handling based on category and description"""
        desc = str(row.get('Description', '')).lower()
        category = str(row.get('Category', '')).lower()
        
        # Keep #N/A for price/rate/index data
        na_keywords = ['price', 'rate', 'spread', 'index', 'benchmark', 'curve']
        # Use 0 for flow/volume data  
        zero_keywords = ['flow', 'consumption', 'production', 'import', 'export', 'demand', 'supply']
        
        if any(kw in desc or kw in category for kw in na_keywords):
            return 'Y'  # Keep #N/A
        elif any(kw in desc or kw in category for kw in zero_keywords):
            return 'N'  # Replace with 0
        else:
            return 'N'  # Default to 0
    
    dataset['Replace blanks with #N/A'] = dataset.apply(smart_na_handling, axis=1)
    print("✅ Smart 'Replace blanks with #N/A' column added")

tickers = list(set(dataset.Ticker.dropna().to_list()))
print(f"Found {len(tickers)} unique tickers")

# Download Bloomberg data with error tracking
try:
    data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)
    available_tickers = set(data.columns)
    requested_tickers = set(tickers)
    missing_tickers = requested_tickers - available_tickers
    
    print(f"✅ Bloomberg data downloaded successfully")
    print(f"📊 Available tickers: {len(available_tickers)}")
    
    if missing_tickers:
        print(f"❌ Missing tickers from Bloomberg: {len(missing_tickers)}")
        print("Missing tickers list:")
        for i, ticker in enumerate(sorted(missing_tickers), 1):
            print(f"  {i:2d}. {ticker}")
        print()
    
except Exception as e:
    print(f"❌ Error downloading Bloomberg data: {e}")
    print("Continuing with available data...")

#%% Create the MultiIndex structure for full_data
print("Processing ticker data...")

# Initialize variables
successful_tickers = []
unavailable_in_processing = []

# Create multiindex lists
date_index = data.index
level_0 = []  # Description
level_1 = []  # Empty
level_2 = []  # Category  
level_3 = []  # Region from
level_4 = []  # Region to

# Process each ticker in dataset
for index, row in dataset.iterrows():
    ticker = row.get('Ticker')
    
    if pd.isna(ticker) or ticker not in data.columns:
        if not pd.isna(ticker):
            # Track unavailable tickers with their details
            unavailable_in_processing.append({
                'ticker': ticker,
                'description': row.get('Description', 'Unknown'),
                'category': row.get('Category', 'Unknown')
            })
        print(f"Skipping row {index}: missing ticker")
        continue
    
    # Add to successful list
    successful_tickers.append(ticker)
    
    # Build MultiIndex levels
    level_0.append(row.get('Description', ''))
    level_1.append('')  # Empty level as in original structure
    level_2.append(row.get('Category', ''))
    level_3.append(row.get('Region from', ''))
    level_4.append(row.get('Region to', ''))

# Create the full_data DataFrame with MultiIndex columns
if successful_tickers:
    # Get data for successful tickers only
    ticker_data = data[successful_tickers]
    
    # Create MultiIndex
    multi_index = pd.MultiIndex.from_arrays([level_0, level_1, level_2, level_3, level_4])
    
    # Create full_data
    full_data = pd.DataFrame(ticker_data.values, index=ticker_data.index, columns=multi_index)
    
    print(f"✅ Created full_data with shape: {full_data.shape}")
else:
    print("❌ No successful tickers to process")
    exit()

print("Applying data fixes...")

# Apply specific data fixes (from original script)
try:
    full_data.iloc[33, 39] = 28
except:
    print("Data fix 1 skipped (position not available)")

try:
    full_data.iloc[1072, 39] = 27
except:
    print("Data fix 2 skipped (position not available)")

try:
    full_data.iloc[1825, 39] = 40
except:
    print("Data fix 3 skipped (position not available)")

try:
    if full_data.shape[1] > 39:
        full_data.iloc[2578, 39] = 44
except:
    print("Data fix 4 skipped (column 39 not available)")

try:
    if full_data.shape[1] > 128:
        full_data.iloc[3103, 128] = 25
except:
    print("Data fix 5 skipped (column 128 not available)")

# Preserve the last row and interpolate only on the t-1 rows
full_data = full_data.copy()
full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')

full_data.index = pd.to_datetime(full_data.index)
full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]

# 1️⃣ DEBUG: Italy in full_data MultiTicker
print('\n=== DEBUG: Italy in full_data MultiTicker ===')
italy_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
             (full_data.columns.get_level_values(3) == 'Italy')
italy_series = full_data.iloc[:, italy_mask]
print(f'  Italy DEMAND series count: {italy_series.shape[1]}')
if italy_series.shape[1] > 0:
    italy_total = italy_series.sum(axis=1, skipna=False)
    print(f'  Italy total from MultiTicker: {italy_total.iloc[0]:.2f}')
    print(f'  Italy series names:')
    for i, col in enumerate(italy_series.columns):
        val = italy_series.iloc[0, i]
        print(f'    {i+1}. {col[0][:50]}... = {val:.2f}')
else:
    print(f'  No Italy DEMAND series in MultiTicker')

# DETAILED TICKER AVAILABILITY REPORT
print(f"\n{'='*60}")
print("📊 TICKER AVAILABILITY REPORT")
print(f"{'='*60}")
print(f"✅ Total tickers requested: {len(tickers)}")
print(f"✅ Successfully processed: {len(successful_tickers)}")
print(f"❌ Unavailable tickers: {len(unavailable_in_processing)}")

if unavailable_in_processing:
    print(f"\n{'='*60}")
    print("❌ UNAVAILABLE TICKERS DETAILS:")
    print(f"{'='*60}")
    print(f"{'#':<3} {'Ticker':<15} {'Description':<50} {'Category':<20}")
    print("-" * 90)
    
    for i, info in enumerate(unavailable_in_processing, 1):
        ticker = info['ticker'][:14]  # Truncate if too long
        desc = info['description'][:49] if len(str(info['description'])) > 49 else str(info['description'])
        cat = info['category'][:19] if len(str(info['category'])) > 19 else str(info['category'])
        print(f"{i:<3} {ticker:<15} {desc:<50} {cat:<20}")
    
    print(f"\n💡 TIP: You may want to:")
    print("   1. Check if these tickers are correctly spelled")
    print("   2. Verify if they're active Bloomberg tickers")
    print("   3. Check if they need different Bloomberg fields")
    print("   4. Consider if they should be removed from the ticker list")
    
print(f"\n{'='*60}")
print("🔄 CONTINUING WITH AVAILABLE DATA...")
print(f"{'='*60}\n")

#%% Industry
print("Processing industry data...")
demand_1 = ['Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand','Demand', 'Intermediate Calculation','Demand','Demand', 'Intermediate Calculation', 'Demand']
demand_2 = ['France','Belgium','Italy','GB','Netherlands','Netherlands','Germany','#Germany','Germany','#Netherlands', '#Netherlands', 'Netherlands']
demand_3 = ['Industrial','Industrial','Industrial','Industrial','Industrial and Power','Zebra','Industrial and Power','Gas-to-Power','Industrial (calculated)','Industrial', 'Gas-to-Power', 'Industrial (calculated to 30/6/22 then actual)']
index = full_data.index

industry = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

for a, b, c in zip(demand_1, demand_2, demand_3):
    l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
    industry.iloc[:, (industry.columns.get_level_values(0)==a) & (industry.columns.get_level_values(2)==c) & (industry.columns.get_level_values(1)==b)] = l

ind_cols = (industry.columns.get_level_values(0)=='Demand') & (industry.columns.get_level_values(1)=='Netherlands') & (industry.columns.get_level_values(2)=='Industrial (calculated to 30/6/22 then actual)')
ind_cols_1 = (industry.columns.get_level_values(0)=='Demand') & (industry.columns.get_level_values(1)=='Netherlands') & (industry.columns.get_level_values(2)=='Industrial and Power')
ind_cols_2 = (industry.columns.get_level_values(0)=='Intermediate Calculation') & (industry.columns.get_level_values(1)=='#Netherlands') & (industry.columns.get_level_values(2)=='Gas-to-Power')
ind_cols_3 = (industry.columns.get_level_values(0)=='Demand') & (industry.columns.get_level_values(1)=='#Netherlands') & (industry.columns.get_level_values(2)=='Industrial')

for ii in range(len(index)):
    if industry.iloc[ii, ind_cols_3][0] == 0:
        if industry.iloc[ii, ind_cols_2][0] > 0:
            industry.iloc[ii, ind_cols] = industry.iloc[ii, ind_cols_1][0] - industry.iloc[ii, ind_cols_2][0]
        else:
            industry.iloc[ii, ind_cols] = np.nan
    else:
        industry.iloc[ii, ind_cols] = industry.iloc[ii, ind_cols_3]

industry[pd.MultiIndex.from_tuples([('Demand','Germany','Industrial (calculated)')])] = industry[pd.MultiIndex.from_tuples([('Demand','Germany','Industrial and Power')])].values - industry[pd.MultiIndex.from_tuples([('Intermediate Calculation','#Germany','Gas-to-Power')])].values
industry[pd.MultiIndex.from_tuples([('Demand','-','Total')])] = pd.DataFrame(industry.iloc[:,[0,1,2,3,8,11]].sum(axis=1, skipna=False))

# 2️⃣ DEBUG: Italy in industry DataFrame
print('\n=== DEBUG: Italy in industry DataFrame ===')
italy_industry_cols = [col for col in industry.columns if 'Italy' in str(col)]
print(f'  Italy industry columns: {italy_industry_cols}')
for col in italy_industry_cols:
    sample_val = industry[col].iloc[0]
    print(f'    {col}: {sample_val:.2f}')

#%% GTP
print("Processing Gas-to-Power data...")
demand_1 = ['Demand', 'Demand', 'Demand' ,'Demand' ,'Demand', 'Intermediate Calculation' ,'Demand' , 'Demand', 'Demand', 'Intermediate Calculation', 'Demand']
demand_2 = ['France','Belgium','Italy','GB','Germany', '#Germany','Germany','Netherlands','#Netherlands','#Netherlands','Netherlands']
demand_3 = ['Gas-to-Power', 'Gas-to-Power' ,'Gas-to-Power' ,'Gas-to-Power', 'Industrial and Power', 'Gas-to-Power', 'Gas-to-Power (calculated)' , 'Industrial and Power', 'Gas-to-Power' ,'Gas-to-Power' ,'Gas-to-Power (calculated to 30/6/22 then actual)']

gtp = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

for a, b, c in zip(demand_1, demand_2, demand_3):
    l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
    gtp.iloc[:, (gtp.columns.get_level_values(0)==a) & (gtp.columns.get_level_values(2)==c) & (gtp.columns.get_level_values(1)==b)] = l

ind_cols = (gtp.columns.get_level_values(0)=='Demand') & (gtp.columns.get_level_values(1)=='#Netherlands') & (gtp.columns.get_level_values(2)=='Gas-to-Power')
ind_cols_1 = (gtp.columns.get_level_values(0)=='Intermediate Calculation') & (gtp.columns.get_level_values(1)=='#Netherlands') & (gtp.columns.get_level_values(2)=='Gas-to-Power')
ind_cols_2 = (gtp.columns.get_level_values(0)=='Demand') & (gtp.columns.get_level_values(1)=='Netherlands') & (gtp.columns.get_level_values(2)=='Gas-to-Power (calculated to 30/6/22 then actual)')

for ii in range(len(index)):
    if gtp.iloc[ii, ind_cols][0] == 0:
        if gtp.iloc[ii, ind_cols_1][0] > 0:
            gtp.iloc[ii, ind_cols_2] = gtp.iloc[ii, ind_cols_1][0]
        else:
            gtp.iloc[ii, ind_cols_2] = np.nan
    else:
        gtp.iloc[ii, ind_cols_2] = gtp.iloc[ii, ind_cols][0]

gtp[pd.MultiIndex.from_tuples([('Demand','Germany','Gas-to-Power (calculated)')])] = gtp[pd.MultiIndex.from_tuples([('Demand','Germany','Industrial and Power')])].values - gtp[pd.MultiIndex.from_tuples([('Intermediate Calculation','#Germany','Gas-to-Power')])].values
gtp[pd.MultiIndex.from_tuples([('Demand','','Total')])] = pd.DataFrame(gtp.iloc[:,[0,1,2,3,6,10]].sum(axis=1, skipna=False))

# 3️⃣ DEBUG: Italy in gtp DataFrame
print('\n=== DEBUG: Italy in gtp DataFrame ===')
italy_gtp_cols = [col for col in gtp.columns if 'Italy' in str(col)]
print(f'  Italy gtp columns: {italy_gtp_cols}')
for col in italy_gtp_cols:
    sample_val = gtp[col].iloc[0]
    print(f'    {col}: {sample_val:.2f}')

#%% LDZ
print("Processing LDZ data...")
demand_1 = ['Demand' ,'Demand' ,'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand (Net)', 'Demand' ,'Delta', 'Delta']
demand_2 = ['France', 'Belgium' ,'Italy' ,'Italy' ,'Netherlands', 'GB', 'Austria' ,'Germany' ,'Switzerland', 'Luxembourg', 'Island of Ireland' , '#Germany' ,'Germany', 'Germany']
demand_3 = ['LDZ', 'LDZ', 'LDZ', 'Other', 'LDZ', 'LDZ', 'Austria', 'LDZ', 'Switzerland', 'Luxembourg', 'Island of Ireland', 'Implied LDZ','', 'Cumulative']

ldz = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

for a, b, c in zip(demand_1, demand_2, demand_3):
    l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
    ldz.iloc[:, (ldz.columns.get_level_values(0)==a) & (ldz.columns.get_level_values(2)==c) & (ldz.columns.get_level_values(1)==b)] = l

ldz[pd.MultiIndex.from_tuples([('Demand','','Total')])] = pd.DataFrame(ldz.iloc[:,[0,1,2,3,4,5,6,7,8,9,10]].sum(axis=1, skipna=False))

# 4️⃣ DEBUG: Italy in ldz DataFrame
print('\n=== DEBUG: Italy in ldz DataFrame ===')
italy_ldz_cols = [col for col in ldz.columns if 'Italy' in str(col)]
print(f'  Italy ldz columns: {italy_ldz_cols}')
for col in italy_ldz_cols:
    sample_val = ldz[col].iloc[0]
    print(f'    {col}: {sample_val:.2f}')

#%% Main panel: Daily historic data by category - IMPROVED CONSISTENT AGGREGATION
print("Processing country data...")

# Define country list for consistent processing
country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']

# Create countries DataFrame with proper structure
demand_1 = ['Demand'] * len(country_list) + ['Demand (Net)']
demand_2 = [''] * (len(country_list) + 1) 
demand_3 = country_list + ['Island of Ireland']

countries = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

# Populate individual countries using consistent logic
for country in country_list:
    # Get all demand data for this country (regardless of subcategory)
    country_mask = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(3) == country)
    country_total = full_data.iloc[:, country_mask].sum(axis=1, skipna=False)
    countries[('Demand', '', country)] = country_total

# 5️⃣ DEBUG: Italy country calculation
print('\n=== DEBUG: Italy country calculation ===')
italy_country_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
                     (full_data.columns.get_level_values(3) == 'Italy')
italy_country_total = full_data.iloc[:, italy_country_mask].sum(axis=1, skipna=False)
print(f'  Italy country mask finds {italy_country_mask.sum()} series')
print(f'  Italy country total: {italy_country_total.iloc[0]:.2f}')

# Show which specific series are being summed for Italy
italy_country_series = full_data.iloc[:, italy_country_mask]
print(f'  Italy country series details:')
for i, col in enumerate(italy_country_series.columns):
    val = italy_country_series.iloc[0, i]
    print(f'    {i+1}. {col[0][:60]}... | Category: {col[2]} | From: {col[3]} | To: {col[4]} = {val:.2f}')

# Handle Island of Ireland separately (Net demand)
ireland_mask = (full_data.columns.get_level_values(2) == 'Demand (Net)') & (full_data.columns.get_level_values(3) == 'Island of Ireland')
ireland_total = full_data.iloc[:, ireland_mask].sum(axis=1, skipna=False) 
countries[('Demand (Net)', '', 'Island of Ireland')] = ireland_total

print("🔧 REVERT TO ORIGINAL PRECISE AGGREGATION METHOD:")
print("   Using existing DataFrames that match Excel SUMIFS logic...")

# REVERT: Use the precise DataFrames that were already built correctly
# These match the Excel SUMIFS logic exactly (country + category filtering)
print("\n🏭 Using Industrial total from industry DataFrame...")
industrial_total = industry[('Demand','-','Total')]

print("\n🏘️ Using LDZ total from ldz DataFrame...")  
ldz_total = ldz[('Demand','','Total')]

print("\n⚡ Using Gas-to-Power total from gtp DataFrame...")
gtp_total = gtp[('Demand','','Total')]

# Calculate country total from individual countries (consistent method)
print("\n🌍 Building Country total from individual countries...")

# Get all country columns properly
country_columns = []
for col in countries.columns:
    if len(col) == 3:
        category, middle, country = col
        if country in country_list and category == 'Demand':
            country_columns.append(col)

if country_columns:
    country_total = countries[country_columns].sum(axis=1, skipna=False)
    print(f"   Found {len(country_columns)} country columns to sum")
else:
    country_total = pd.Series(0.0, index=index)
    print("   Warning: No country columns found!")

# Add Ireland separately if it exists
ireland_col = ('Demand (Net)', '', 'Island of Ireland')
if ireland_col in countries.columns:
    country_total += countries[ireland_col]
    print("   Added Ireland (Net demand)")

# Add all totals to countries DataFrame with correct 3-level structure
countries[('', '', 'Total')] = country_total
countries[('', '', 'Industrial')] = industrial_total
countries[('', '', 'LDZ')] = ldz_total
countries[('', '', 'Gas-to-Power')] = gtp_total

# ENHANCED DEBUG: Check if sums add up with ORIGINAL precise method
print("\n🔍 ENHANCED DEBUGGING WITH ORIGINAL PRECISE METHOD...")
total_col = countries[('','','Total')]
industrial_col = countries[('','','Industrial')]
ldz_col = countries[('','','LDZ')]
gtp_col = countries[('','','Gas-to-Power')]

manual_sum = industrial_col + ldz_col + gtp_col
difference = total_col - manual_sum
max_diff = abs(difference).max()

print(f"📊 ORIGINAL PRECISE AGGREGATION RESULTS:")
print(f"   Country Total vs (Industrial + LDZ + GTP)")
print(f"   Maximum difference: {max_diff:.6f}")
print(f"   Mean difference: {difference.mean():.6f}")
print(f"   Standard deviation: {difference.std():.6f}")

# Check what percentage of total this represents
avg_total = total_col.mean()
percentage_error = (max_diff / avg_total) * 100 if avg_total > 0 else 0

print(f"   Average total demand: {avg_total:.2f}")
print(f"   Max error as % of total: {percentage_error:.4f}%")

if max_diff > 0.1:  # Lower tolerance for better validation
    print(f"⚠️  DIFFERENCES DETECTED (>{0.1} threshold):")
    
    # Show a sample of the problematic data
    sample_idx = abs(difference).idxmax()
    print(f"\n   Worst case on {sample_idx}:")
    print(f"   Total: {total_col[sample_idx]:.6f}")
    print(f"   Industrial: {industrial_col[sample_idx]:.6f}")
    print(f"   LDZ: {ldz_col[sample_idx]:.6f}")  
    print(f"   Gas-to-Power: {gtp_col[sample_idx]:.6f}")
    print(f"   Sum: {manual_sum[sample_idx]:.6f}")
    print(f"   Difference: {difference[sample_idx]:.6f}")
    
    # Additional diagnostic information
    print(f"\n🔍 DIAGNOSTIC INFO:")
    all_demand_total = full_data.iloc[:, full_data.columns.get_level_values(2)=='Demand'].sum(axis=1, skipna=False)
    print(f"   Raw total demand (all 'Demand' series): {all_demand_total[sample_idx]:.6f}")
    print(f"   Our country total: {total_col[sample_idx]:.6f}")
    print(f"   Difference from raw: {all_demand_total[sample_idx] - total_col[sample_idx]:.6f}")
    
else:
    print(f"✅ EXCELLENT! Differences are within acceptable tolerance")
    print(f"   The consistent aggregation method has resolved the discrepancies!")

print(f"\n=== ITALY DEBUG SUMMARY ===")
print(f"Italy value at each stage:")
print(f"  1. MultiTicker: {italy_total.iloc[0] if italy_series.shape[1] > 0 else 0:.2f}")
if italy_industry_cols:
    italy_ind_total = sum(industry[col].iloc[0] for col in italy_industry_cols)
    print(f"  2. Industry: {italy_ind_total:.2f}")
else:
    print(f"  2. Industry: 0.00")
if italy_gtp_cols:
    italy_gtp_total = sum(gtp[col].iloc[0] for col in italy_gtp_cols)
    print(f"  3. GTP: {italy_gtp_total:.2f}")
else:
    print(f"  3. GTP: 0.00")
if italy_ldz_cols:
    italy_ldz_total = sum(ldz[col].iloc[0] for col in italy_ldz_cols) 
    print(f"  4. LDZ: {italy_ldz_total:.2f}")
else:
    print(f"  4. LDZ: 0.00")
print(f"  5. Final Country: {italy_country_total.iloc[0]:.2f}")

# Continue with rest of original script...
print("Processing LNG imports by country...")
print("Processing calendar year analysis...")
print("Writing consolidated file with all 17 sheets...")
print(f"\n{'='*50}")
print("DEBUG VERSION - EXECUTION STOPPED HERE")
print("Add the rest of your script below this line...")
print(f"{'='*50}")