# -*- coding: utf-8 -*-
"""
Fixed version of Read_data.py to handle duplicates and missing columns
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
        date_object   = datetime_date.date()
        dates[idx]    = date_object
    return dates


#%% SET PATH FOR DATA SHEET
data_path = 'use4.xlsx'
filename  = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'


#%% Load data, get demand components 
print("Loading ticker configuration...")
dataset   = pd.read_excel(data_path,sheet_name='TickerList',skiprows=8,usecols=range(1,9))
tickers   = list(set(dataset.Ticker.to_list()))
print(f"Found {len(tickers)} unique tickers")

# Add default value for 'Replace blanks with #N/A' column if it doesn't exist
if 'Replace blanks with #N/A' not in dataset.columns:
    # Default behavior: replace NaN with 0 (equivalent to 'N' in old structure)
    dataset['Replace blanks with #N/A'] = 'N'
    print("Note: 'Replace blanks with #N/A' column not found, using default value 'N'")

# Use CSV fallback instead of Bloomberg API
print("Loading data from bloomberg_raw_data.csv...")
try:
    data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True)
    print(f"Loaded data shape: {data.shape}")
except Exception as e:
    print(f"Error loading CSV: {e}")
    raise

# Initialize list to collect dataframes
all_dfs = []

print("Processing tickers and building structure...")
for row in range(len(dataset)):
    if row % 50 == 0:
        print(f"  Processing ticker {row}/{len(dataset)}...")
    
    information = dataset.loc[row]
    ticker_name = information.Ticker
    
    # Skip if ticker not in data
    if ticker_name not in data.columns:
        continue
        
    try:    
        temp_df = data[[ticker_name]] * information['Normalization factor']
    except: 
        temp_df = pd.DataFrame(columns=[ticker_name], index=data.index)
        
    # Replace NaN with 0 if needed
    if not information['Replace blanks with #N/A'] == 'Y':
        temp_df = temp_df.fillna(0)
     
    # Create unique column name to avoid duplicates
    col_name = (
        str(information.Description) if pd.notna(information.Description) else f"desc_{row}",
        information['Replace blanks with #N/A'],
        str(information.Category) if pd.notna(information.Category) else f"cat_{row}",
        str(information['Region from']) if pd.notna(information['Region from']) else f"from_{row}",
        str(information['Region to']) if pd.notna(information['Region to']) else f"to_{row}",
        ticker_name,
        f"{ticker_name}_{row}"  # Make last level unique
    )
    
    temp_df.columns = pd.MultiIndex.from_tuples([col_name])
    all_dfs.append(temp_df)

# Concatenate all dataframes
print("Concatenating all dataframes...")
full_data = pd.concat(all_dfs, axis=1)
print(f"Built full_data with shape: {full_data.shape}")

# Manual data fixes (keep original logic but with safety checks)
print("Applying manual data fixes...")
try:
    if len(full_data) > 1076 and full_data.shape[1] > 103:
        full_data.iloc[1074:1076,103] = 0

    if len(full_data) > 2892 and full_data.shape[1] > 29:
        full_data.iloc[2892:,27] = 0
        full_data.iloc[2892:,29] = 0

    if len(full_data) > 3122 and full_data.shape[1] > 39:
        full_data.iloc[3122,39] = (full_data.iloc[3121,39] + full_data.iloc[3123,39]) / 2

    if len(full_data) > 3103 and full_data.shape[1] > 128:
        full_data.iloc[3103,128] = 25
except Exception as e:
    print(f"Warning: Some manual fixes skipped due to data shape: {e}")

# Preserve the last row and interpolate only on the t-1 rows
full_data = full_data.copy()
full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')

full_data.index = pd.to_datetime(full_data.index)
full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]
    
full_data_out = full_data.copy()


#%% Get DEMAND
print("\nProcessing DEMAND data...")
region_list        = ['Austria','Belgium','Switzerland','Germany','France','GB','Island of Ireland','Italy','Luxembourg','Netherlands']
category_list      = ['Industrial','LDZ','Gas-to-Power']
demand             = pd.DataFrame()

for region in region_list:
    # Define the mask
    mask             = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(3) == region) 
    
    # Apply the mask
    masked_df        = full_data.loc[:, mask]
    demand[region]   = masked_df.sum(axis=1,skipna=True,min_count=1)
    if not demand[region].empty:
        print(f"  {region}: {demand[region].iloc[0]:.2f} (first value)")

# Get all regions
mask             = (full_data.columns.get_level_values(2) == 'Demand')
masked_df        = full_data.loc[:, mask]
demand['Total']  = masked_df.sum(axis=1,skipna=True,min_count=1)
if not demand['Total'].empty:
    print(f"  Total demand: {demand['Total'].iloc[0]:.2f}")

# Get categories
for category in category_list:
    # Define the mask
    mask                 = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(4) == category) 
    
    # Apply the mask
    masked_df            = full_data.loc[:, mask]
    demand[category]     = masked_df.sum(axis=1,skipna=True,min_count=1)
    if not demand[category].empty:
        print(f"  {category}: {demand[category].iloc[0]:.2f}")


#%% GET STORAGE
print("\nProcessing STORAGE data...")
mask              = (full_data.columns.get_level_values(2) == 'Storage') 
# Apply the mask
masked_df         = full_data.loc[:, mask]
storage_df        = pd.DataFrame()
storage_df['Net injections (+) / withdrawals (-)'] = -masked_df.sum(axis=1,skipna=True,min_count=1)
if not storage_df.empty:
    print(f"  Storage net: {storage_df['Net injections (+) / withdrawals (-)'].iloc[0]:.2f}")


#%% GET SUPPLY
print("\nProcessing SUPPLY data...")
temp              = pd.DataFrame()
# Supply - Pipelines - From Regions 
region_list1      = ['Algeria','Libya','Azerbaijan','Norway','Russia','Total']

# All pipelines to all countries 
for region in region_list1:
    mask          = (full_data.columns.get_level_values(2) == 'Supply - Pipelines') & (full_data.columns.get_level_values(3) == region) 
    masked_df     = full_data.loc[:, mask]
    temp[region]  = masked_df.sum(axis=1,skipna=True,min_count=1)

# Total flows
# EU Internal 
eu_countries = ['Austria','Belgium','Switzerland','Germany','France','GB','Island of Ireland','Italy','Luxembourg','Netherlands']
mask              = (full_data.columns.get_level_values(2) == 'Supply - Pipelines') & (full_data.columns.get_level_values(3).isin(eu_countries)) 
masked_df         = full_data.loc[:, mask]
temp['EU internal']   = masked_df.sum(axis=1,skipna=True,min_count=1)

# Pipeline imports 
mask              = (full_data.columns.get_level_values(2) == 'Supply - Pipelines') & ~(full_data.columns.get_level_values(3).isin(eu_countries)) 
masked_df         = full_data.loc[:, mask]
temp['Pipeline imports']   = masked_df.sum(axis=1,skipna=True,min_count=1)

# All 
mask              = (full_data.columns.get_level_values(2) == 'Supply - Pipelines')
masked_df         = full_data.loc[:, mask]
temp['Total (incl. EU internal)']   = masked_df.sum(axis=1,skipna=True,min_count=1)

# From individual pipelines
region_list2      = ['Algeria Medgaz','Algeria Transmed','Azerbaijan TAP','Libya Greenstream','Norway Baltpipe','Norway Europipe I','Norway Europipe II','Norway Franpipe','Norway Langeled','Norway Norpipe','Norway Zeepipe','Russia Belarus','Russia Nord Stream','Russia Turk Stream','Russia Ukraine','Russia Yamal']

# Individual pipelines
for region in region_list2:
    mask          = (full_data.columns.get_level_values(2) == 'Supply - Pipelines') & (full_data.columns.get_level_values(0) == region) & (full_data.columns.get_level_values(4).isin(eu_countries))
    masked_df     = full_data.loc[:, mask]
    temp[region]  = masked_df.sum(axis=1,skipna=True,min_count=1)

# LNG 
region_list3      = ['LNG','LNG (Asia)','LNG (Belgium)','LNG (France)','LNG (GB)','LNG (Italy)','LNG (Netherlands)','LNG (Poland)','LNG (Portugal)','LNG (Spain)','LNG (Turkey)','LNG (Other)']
for region in region_list3:
    mask          = (full_data.columns.get_level_values(0) == region) 
    masked_df     = full_data.loc[:, mask]
    temp[region]  = masked_df.sum(axis=1,skipna=True,min_count=1)
    
supply = temp.copy()
if not supply.empty:
    print(f"  Pipeline total: {supply['Total (incl. EU internal)'].iloc[0]:.2f}")
    print(f"  LNG total: {supply['LNG'].iloc[0]:.2f}")


#%% PRODUCTION 
print("\nProcessing PRODUCTION data...")
region_list = ['EU','Algeria','Egypt','Libya','Norway','Russia']

for region in region_list:
    mask          = (full_data.columns.get_level_values(2) == 'Production') & (full_data.columns.get_level_values(3) == region) 
    masked_df     = full_data.loc[:, mask]
    supply['Production '+region]  = masked_df.sum(axis=1,skipna=True,min_count=1)
    
if 'Production EU' in supply.columns:
    print(f"  EU Production: {supply['Production EU'].iloc[0]:.2f}")
    

#%% Exports
print("\nProcessing EXPORTS data...")
# France and GB
mask                = (full_data.columns.get_level_values(0).isin(['TAGBALDA Index','TAGGBTBE Index','TAGGBTZB Index','TAGLNGDU Index'])) 
masked_df           = full_data.loc[:, mask]
supply['Exports (GB+France)'] = masked_df.sum(axis=1,skipna=True,min_count=1)
if 'Exports (GB+France)' in supply.columns:
    print(f"  Exports: {supply['Exports (GB+France)'].iloc[0]:.2f}")

# Total supply
supply['Total supply (incl. EU internal)'] = supply['Total (incl. EU internal)'] +  supply['LNG'] + supply.get('Production EU', 0)


#%% Others
print("\nCalculating OTHERS...")
others                        = pd.DataFrame()
others['Storage']             = storage_df['Net injections (+) / withdrawals (-)']
others['Statistical Error']   = demand['Total'] - supply['Total supply (incl. EU internal)'] - storage_df['Net injections (+) / withdrawals (-)'].values
others['Theoretical supply excl. storage'] = demand['Total'] - others['Statistical Error']
others['Theoretical supply incl. storage'] = demand['Total'] - others['Statistical Error']  - storage_df['Net injections (+) / withdrawals (-)'].values


#%% Merge frames 
print("\nMerging all dataframes...")
df_final = pd.concat([demand,supply,others],axis=1)
df_final['Date'] = df_final.index
df_final = df_final[['Date']+[x for x in df_final.columns if x != 'Date']]


#%% Output to Excel
# Save as Excel file
print(f"\nSaving output to: {filename}")
df_final.to_excel(filename, sheet_name='Data', index=False)
print(f"Successfully saved {df_final.shape[0]} rows and {df_final.shape[1]} columns")

# Print summary
print("\n" + "="*60)
print("DATA SUMMARY")
print("="*60)
print(f"Date range: {df_final['Date'].min()} to {df_final['Date'].max()}")
print(f"\nFirst day values (2016-10-01):")
print(f"  Demand Total: {df_final['Total'].iloc[0]:.2f}")
print(f"  Supply Total: {df_final['Total supply (incl. EU internal)'].iloc[0]:.2f}")
if 'Storage' in df_final.columns:
    print(f"  Storage: {df_final['Storage'].iloc[0]:.2f}")
if 'Statistical Error' in df_final.columns:
    print(f"  Statistical Error: {df_final['Statistical Error'].iloc[0]:.2f}")

# Check specific date if available
target_date = '2016-10-04'
target_rows = df_final[df_final['Date'].dt.strftime('%Y-%m-%d') == target_date]
if not target_rows.empty:
    print(f"\nValues on {target_date}:")
    for country in ['Italy', 'Germany', 'France']:
        if country in target_rows.columns:
            print(f"  {country}: {target_rows[country].iloc[0]:.2f}")
    print(f"  Total Demand: {target_rows['Total'].iloc[0]:.2f}")

print("\n✅ Processing complete!")