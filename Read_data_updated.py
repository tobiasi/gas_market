# -*- coding: utf-8 -*-
"""
Updated version of Read_data.py to work with new TickerList structure
Minimal changes to handle missing 'Replace blanks with #N/A' column
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


from xbbg                import blp
import pandas            as pd
import numpy             as np
import sys
import pycountry_convert as pc

# Comment out if update_spreadsheet not available
# sys.path.append("C:/development/commodities")
# from update_spreadsheet import update_spreadsheet

from datetime import datetime
from xbbg                import blp
import sys
from dateutil.relativedelta import relativedelta
import numpy as np

# Comment out if update_spreadsheet not available
# sys.path.append("C:/development/commodities")
# from update_spreadsheet import update_spreadsheet, to_excel_dates,update_spreadsheet_pweekly


#%% SET PATH FOR DATA SHEET
# Update paths as needed
# os.chdir('G:\Commodity Research\Gas')
# data_path = 'G://Commodity Research//Gas//data//use3.xlsm'

# Use current directory and use4.xlsx
data_path = 'use4.xlsx'
filename  = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'


#%% Load data, get demand components 
dataset   = pd.read_excel(data_path,sheet_name='TickerList',skiprows=8,usecols=range(1,9))
tickers   = list(set(dataset.Ticker.to_list()))

# Add default value for 'Replace blanks with #N/A' column if it doesn't exist
if 'Replace blanks with #N/A' not in dataset.columns:
    # Default behavior: replace NaN with 0 (equivalent to 'N' in old structure)
    dataset['Replace blanks with #N/A'] = 'N'
    print("Note: 'Replace blanks with #N/A' column not found, using default value 'N'")

data      = blp.bdh(tickers,start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)
full_data = pd.DataFrame()

for row in range(len(dataset)):
    information = dataset.loc[row]
    try:    
        temp_df     = data[[information.Ticker]]*information['Normalization factor']
    except: 
        temp_df     =  pd.DataFrame(columns= [information.Ticker],index=data.index )
        
    if not information['Replace blanks with #N/A'] == 'Y':
        temp_df[temp_df.isna().values] = 0
     
    
    temp_df.columns = pd.MultiIndex.from_tuples([(information.Description,information['Replace blanks with #N/A'],information.Category,information['Region from'],information['Region to'],information.Ticker,information.Ticker)])

    if row == 0:
        full_data = temp_df
    else:
        full_data = pd.merge(full_data,temp_df,left_index=True, right_index=True,how="outer")
   
    

# Manual data fixes (keep original logic)
#full_data.iloc[2318,40] = 16.44
full_data.iloc[1074:1076,103] = 0

#full_data.iloc[1074:1076,68] = 0

full_data.iloc[2892:,27] = 0

full_data.iloc[2892:,29] = 0


full_data.iloc[3122,39] = (full_data.iloc[3121,39] +full_data.iloc[3123,39] )/2

full_data.iloc[3103,128] = 25

# Preserve the last row and interpolate only on the t-1 rows
full_data = full_data.copy()
full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')



full_data.index = pd.to_datetime(full_data.index)
full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]
    
full_data_out = full_data.copy()


#%% Get DEMAND
region_list        = ['Austria','Belgium','Switzerland','Germany','France','GB','Island of Ireland','Italy','Luxembourg','Netherlands']
category_list      = ['Industrial','LDZ','Gas-to-Power']
demand             = pd.DataFrame()
for region in region_list:
    
    # Define the mask
    mask             = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(3) == region) 
    
    # Apply the mask
    masked_df        = full_data.loc[:, mask]
    demand[region]   = masked_df.sum(axis=1,skipna=True,min_count=1)

# Get all regions
mask             = (full_data.columns.get_level_values(2) == 'Demand')

masked_df        = full_data.loc[:, mask]
demand['Total']  = masked_df.sum(axis=1,skipna=True,min_count=1)

# Get categories
for category in category_list:
    
    # Define the mask
    mask                 = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(4) == category) 
    
    # Apply the mask
    masked_df            = full_data.loc[:, mask]
    demand[category]     = masked_df.sum(axis=1,skipna=True,min_count=1)



#%% GET STORAGE
mask              = (full_data.columns.get_level_values(2) == 'Storage') 
# Apply the mask
masked_df         = full_data.loc[:, mask]
storage_df        = pd.DataFrame()
storage_df['Net injections (+) / withdrawals (-)'] = -masked_df.sum(axis=1,skipna=True,min_count=1)




#%% GET SUPPLY
temp              = pd.DataFrame()
# Supply - Pipelines - From Regions 
region_list1      = ['Algeria','Libya','Azerbaijan','Norway','Russia','Total']
region_tot_list   = ['EU internal','Pipeline imports']

# All pipelines to all countries 
for region in region_list1:
    mask          = (full_data.columns.get_level_values(2) == 'Supply - Pipelines') & (full_data.columns.get_level_values(3) == region) 
    masked_df     = full_data.loc[:, mask]
    temp[region]  = masked_df.sum(axis=1,skipna=True,min_count=1)

# Total flows
# EU Internal 
mask              = (full_data.columns.get_level_values(2) == 'Supply - Pipelines') & (full_data.columns.get_level_values(3).isin(['Austria','Belgium','Switzerland','Germany','France','GB','Island of Ireland','Italy','Luxembourg','Netherlands'])) 
masked_df         = full_data.loc[:, mask]
temp['EU internal']   = masked_df.sum(axis=1,skipna=True,min_count=1)

# Pipeline imports 
mask              = (full_data.columns.get_level_values(2) == 'Supply - Pipelines') & ~(full_data.columns.get_level_values(3).isin(['Austria','Belgium','Switzerland','Germany','France','GB','Island of Ireland','Italy','Luxembourg','Netherlands'])) 
masked_df         = full_data.loc[:, mask]
temp['Pipeline imports']   = masked_df.sum(axis=1,skipna=True,min_count=1)

# All 
mask              = (full_data.columns.get_level_values(2) == 'Supply - Pipelines')
masked_df         = full_data.loc[:, mask]
temp['Total (incl. EU internal)']   = masked_df.sum(axis=1,skipna=True,min_count=1)



# From individual pipelines, only sum those to EU
region_list2      = ['Algeria Medgaz','Algeria Transmed','Azerbaijan TAP','Libya Greenstream','Norway Baltpipe','Norway Europipe I','Norway Europipe II','Norway Franpipe','Norway Langeled','Norway Norpipe','Norway Zeepipe','Russia Belarus','Russia Nord Stream','Russia Turk Stream','Russia Ukraine','Russia Yamal']
region_list2b     = ['Algeria Medgaz','Algeria Transmed','Azerbaijan TAP','Libya Greenstream','Norway Baltpipe','Norway Europipe I','Norway Europipe II','Norway Franpipe','Norway Langeled','Norway Norpipe','Norway Zeepipe','Russia Belarus','Russia Nord Stream','Russia Turk Stream','Russia Ukraine','Russia Yamal']

# EDIT: The individual pipeline (region_list2), to country in EU 
# (full_data.columns.get_level_values(4).isin(['Austria','Belgium','Switzerland','Germany','France','GB','Island of Ireland','Italy','Luxembourg','Netherlands']))


# Individual pipelines
for count, region in enumerate(region_list2):
    mask          = (full_data.columns.get_level_values(2) == 'Supply - Pipelines') & (full_data.columns.get_level_values(0) == region) & (full_data.columns.get_level_values(4).isin(['Austria','Belgium','Switzerland','Germany','France','GB','Island of Ireland','Italy','Luxembourg','Netherlands']))
    masked_df     = full_data.loc[:, mask]
    temp[region_list2b[count]]  = masked_df.sum(axis=1,skipna=True,min_count=1)





# LNG 
region_list3      = ['LNG','LNG (Asia)','LNG (Belgium)','LNG (France)','LNG (GB)','LNG (Italy)','LNG (Netherlands)','LNG (Poland)','LNG (Portugal)','LNG (Spain)','LNG (Turkey)','LNG (Other)']
for region in region_list3:
    mask          = (full_data.columns.get_level_values(0) == region) 
    masked_df     = full_data.loc[:, mask]
    temp[region]  = masked_df.sum(axis=1,skipna=True,min_count=1)
    

supply = temp.copy()




#%% PRODUCTION 
region_list = ['EU','Algeria','Egypt','Libya','Norway','Russia']

for region in region_list:
    mask          = (full_data.columns.get_level_values(2) == 'Production') & (full_data.columns.get_level_values(3) == region) 
    masked_df     = full_data.loc[:, mask]
    supply['Production '+region]  = masked_df.sum(axis=1,skipna=True,min_count=1)
    
    
    

#%% Exports
# France and GB
mask                = (full_data.columns.get_level_values(0).isin(['TAGBALDA Index','TAGGBTBE Index','TAGGBTZB Index','TAGLNGDU Index'])) 
masked_df           = full_data.loc[:, mask]
supply['Exports (GB+France)'] = masked_df.sum(axis=1,skipna=True,min_count=1)



# Total supply
supply['Total supply (incl. EU internal)'] = supply['Total (incl. EU internal)'] +  supply['LNG'] + supply['Production EU']


#%% Others
others                        = pd.DataFrame()
others['Storage']             = storage_df['Net injections (+) / withdrawals (-)']
others['Statistical Error']   = demand['Total'] - supply['Total supply (incl. EU internal)'] - storage_df['Net injections (+) / withdrawals (-)'].values
others['Theoretical supply excl. storage'] = demand['Total'] - others['Statistical Error']
others['Theoretical supply incl. storage'] = demand['Total'] - others['Statistical Error']  - storage_df['Net injections (+) / withdrawals (-)'].values


#%% Merge frames 
df_final = pd.concat([demand,supply,others],axis=1)
df_final['Date'] = df_final.index
df_final = df_final[['Date']+[x for x in df_final.columns if x != 'Date']]


#%% Output to Excel
# Save as Excel file
print(f"Saving output to: {filename}")
df_final.to_excel(filename, sheet_name='Data', index=False)
print(f"Successfully saved {df_final.shape[0]} rows and {df_final.shape[1]} columns")

# Print summary
print("\nData Summary:")
print(f"Date range: {df_final['Date'].min()} to {df_final['Date'].max()}")
print(f"Demand Total (first day): {df_final['Total'].iloc[0]:.2f}")
print(f"Supply Total (first day): {df_final['Total supply (incl. EU internal)'].iloc[0]:.2f}")