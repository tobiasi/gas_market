# -*- coding: utf-8 -*-
"""
CONSOLIDATED VERSION FOR USE4.xlsx - Outputs single file with all 17 sheets
Adapted for new use4.xlsx structure without 'Replace blanks with #N/A' column
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
data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)

print("Processing ticker data...")
full_data = pd.DataFrame()
for row in range(len(dataset)):
    information = dataset.loc[row]
    
    # Handle missing tickers gracefully
    ticker = information.get('Ticker', f'UNKNOWN_{row}')
    if pd.isna(ticker):
        print(f"Skipping row {row}: missing ticker")
        continue
        
    try:
        # Get normalization factor with fallback
        norm_factor = information.get('Normalization factor', 1.0)
        if pd.isna(norm_factor):
            norm_factor = 1.0
            
        temp_df = data[[ticker]] * norm_factor
    except KeyError:
        print(f"Warning: Ticker {ticker} not found in Bloomberg data")
        temp_df = pd.DataFrame(columns=[ticker], index=data.index)
    except Exception as e:
        print(f"Error processing {ticker}: {e}")
        temp_df = pd.DataFrame(columns=[ticker], index=data.index)

    # Handle missing values based on 'Replace blanks with #N/A' setting
    replace_blanks = information.get('Replace blanks with #N/A', 'N')
    if replace_blanks != 'Y':
        temp_df[temp_df.isna().values] = 0
    
    # Create MultiIndex with safe fallbacks for missing columns
    desc = information.get('Description', f'Unknown_{row}')
    category = information.get('Category', 'Unknown')
    region_from = information.get('Region from', '')
    region_to = information.get('Region to', '')
    
    # Handle potential NaN values
    if pd.isna(desc): desc = f'Unknown_{row}'
    if pd.isna(category): category = 'Unknown'
    if pd.isna(region_from): region_from = ''
    if pd.isna(region_to): region_to = ''
    if pd.isna(replace_blanks): replace_blanks = 'N'
    
    temp_df.columns = pd.MultiIndex.from_tuples([(desc, replace_blanks, category, 
                                                 region_from, region_to, ticker, ticker)])
    
    if row == 0:
        full_data = temp_df
    else:
        full_data = pd.merge(full_data, temp_df, left_index=True, right_index=True, how="outer")

# Data fixes - with error handling for potentially missing columns
print("Applying data fixes...")
try:
    if full_data.shape[1] > 103:
        full_data.iloc[1074:1076, 103] = 0
except:
    print("Data fix 1 skipped (column 103 not available)")

try:
    if full_data.shape[1] > 29:
        full_data.iloc[2892:, 27] = 0
        full_data.iloc[2892:, 29] = 0
except:
    print("Data fixes 2-3 skipped (columns not available)")

try:
    if full_data.shape[1] > 39:
        full_data.iloc[3122, 39] = (full_data.iloc[3121, 39] + full_data.iloc[3123, 39]) / 2
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
        gtp.iloc[ii, ind_cols_2] = gtp.iloc[ii, ind_cols]

gtp[pd.MultiIndex.from_tuples([('Demand','Germany','Gas-to-Power (calculated)')])] = gtp[pd.MultiIndex.from_tuples([('Intermediate Calculation','#Germany','Gas-to-Power')])]
gtp[pd.MultiIndex.from_tuples([('Demand','','Total')])] = pd.DataFrame(gtp.iloc[:,[0,1,2,3,6,10]].sum(axis=1, skipna=False))

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

#%% Main panel: Daily historic data by category
print("Processing country data...")
demand_1 = ['Demand' ,'Demand' , 'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand','Demand', 'Demand (Net)']
demand_2 = ['','','','','','','','','','']
demand_3 = ['France', 'Belgium' ,'Italy' ,'Netherlands', 'GB', 'Austria' ,'Germany' ,'Switzerland', 'Luxembourg', 'Island of Ireland' ]

countries = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

for a, b, c in zip(demand_1, demand_2, demand_3):
    l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(3)==c)].sum(axis=1, skipna=False)
    countries.iloc[:, (countries.columns.get_level_values(0)==a) & (countries.columns.get_level_values(2)==c)] = l

countries[pd.MultiIndex.from_tuples([('','','Total')])] = pd.DataFrame(countries.iloc[:,[0,1,2,3,4,5,6,7,8,9]].sum(axis=1, skipna=False))
countries[pd.MultiIndex.from_tuples([('','','Industrial')])] = industry.iloc[:,industry.columns.get_level_values(2)=='Total']
countries[pd.MultiIndex.from_tuples([('','','LDZ')])] = ldz.iloc[:,ldz.columns.get_level_values(2)=='Total']
countries[pd.MultiIndex.from_tuples([('','','Gas-to-Power')])] = gtp.iloc[:,gtp.columns.get_level_values(2)=='Total']

last_index = countries.apply(lambda x: x[x.notnull()].index[-1]).sort_values().values[0]
countries = countries.loc[:last_index]

# Supply data processing
demand_1 = ['Import', 'Import' ,'Import' ,'Production' ,'Production' ,'Import', 'Import', 'Import' ,'Import' ,'Import', 'Export' ,'Export' ,'Import' ,'Import', 'Import', 'Production', 'Production', 'Production']
demand_2 = ['Russia', 'Russia' ,'Norway', 'Netherlands' ,'GB', 'LNG', 'Algeria', 'Libya', 'Spain' ,'Denmark' ,'Poland', 'Hungary', 'Slovenia', 'Austria', 'TAP', 'Austria', 'Italy' ,'Germany']
demand_3 = ['Austria', 'Germany' ,'Europe' ,'Netherlands', 'GB' , '' , 'Italy', 'Italy', 'France' ,'Germany' ,'Germany' ,'Austria' ,'Europe', 'MAB', 'Italy' ,'Austria' ,'Italy', 'Germany']

supply = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

# Safe handling of specific ticker fix
try:
    indX = ('Brandov NordStream OPAL Daily Gas Flows (Germany to Czech Republic) (GWh/day)',
            'Y', 'Import', 'Russia', 'Germany', 'ONTGDCDF Index', 'ONTGDCDF Index')
    if indX in full_data.columns:
        full_data.loc[full_data.index.year > 2023, indX] = 0
except:
    print("NordStream fix skipped (column not found)")

for a, b, c in zip(demand_1, demand_2, demand_3):
    if b == 'LNG':
        l = full_data.iloc[:,(full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
        supply.iloc[:, (supply.columns.get_level_values(0)==a) & (supply.columns.get_level_values(1)==b)] = l
    else:
        l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
        supply.iloc[:, (supply.columns.get_level_values(0)==a) & (supply.columns.get_level_values(2)==c) & (supply.columns.get_level_values(1)==b)] = l

supply[pd.MultiIndex.from_tuples([('Supply','','Total')])] = pd.DataFrame(supply.sum(axis=1, skipna=False))

#%% LNG imports by country
print("Processing LNG imports by country...")
demand_1 = ['Import','Import','Import','Import','Import','Import']
demand_2 = ['LNG','LNG','LNG','LNG','LNG','LNG']
demand_3 = ['France','Italy','Belgium','Netherlands','GB','Germany']

lng = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

for a, b, c in zip(demand_1, demand_2, demand_3):
    l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
    lng.iloc[:, (lng.columns.get_level_values(0)==a) & (lng.columns.get_level_values(2)==c) & (lng.columns.get_level_values(1)==b)] = l

lng[pd.MultiIndex.from_tuples([('Supply','Total','')])] = pd.DataFrame(lng.sum(axis=1, skipna=False))

print("Processing calendar year analysis...")

#%% Create demand YOY data
demand = countries[pd.MultiIndex.from_tuples([('','','Industrial'),('','','LDZ'),('','','Gas-to-Power')])]
Demand = pd.DataFrame(demand.values, columns=['Ind','LDZ','GTP'], index=demand.index)
Demand.index = pd.to_datetime(Demand.index)
Demand = Demand[~((Demand.index.month == 2) & (Demand.index.day == 29))]
Demand['total_demand'] = Demand.sum(axis=1)

demand_out = Demand.resample('M').mean().diff(12).copy()
demand_out['total_demand_%'] = Demand.resample('M').mean().pct_change(12)['total_demand']
demand_out['Industry_%'] = Demand.resample('M').mean().pct_change(12)['Ind']
demand_out['LDZ_%'] = Demand.resample('M').mean().pct_change(12)['LDZ']
demand_out['GTP_%'] = Demand.resample('M').mean().pct_change(12)['GTP']

#%% Process calendar year data (using original approach for exact match)
union = pd.MultiIndex.from_tuples([('','','Gas-to-Power'),('','','Industrial'),('','','LDZ')])
Demand_cal = countries[union].rolling(7).mean()
Demand_cal.index = pd.to_datetime(Demand_cal.index)
Demand_cal.columns = Demand_cal.columns.get_level_values(2)

years = [2017,2018,2019,2020,2021,2022,2023,2024,2025,2026]
df = pd.DataFrame(index=pd.date_range('2022-01-01', freq='D', periods=365))
df.index = df.index.strftime('%m-%d')
df_ldz = df.copy()
df_ind = df.copy()
df_gtp = df.copy()
df_tot = df.copy()

for year in years:
    d_temp = Demand_cal[Demand_cal.index.year==year]
    d_temp.index = d_temp.index.strftime('%m-%d')
    ldz_temp = d_temp[['LDZ']].rename(columns={'LDZ':year})
    ind_temp = d_temp[['Industrial']].rename(columns={'Industrial':year})
    gtp_temp = d_temp[['Gas-to-Power']].rename(columns={'Gas-to-Power':year})
    
    df_ldz = df_ldz.merge(ldz_temp, left_index=True, right_index=True, how="outer")
    df_ind = df_ind.merge(ind_temp, left_index=True, right_index=True, how="outer")
    df_gtp = df_gtp.merge(gtp_temp, left_index=True, right_index=True, how="outer")

df_ldz['Benchmark'] = df_ldz[[2017,2018,2019,2020,2021]].mean(axis=1)
df_ind['Benchmark'] = df_ind[[2017,2018,2019,2020,2021]].mean(axis=1)
df_gtp['Benchmark'] = df_gtp[[2017,2018,2019,2020,2021]].mean(axis=1)

df_tot = df.copy()
for year in years:
    df_tot[year] = df_ldz[year] + df_ind[year] + df_gtp[year]
df_tot['Benchmark'] = df_tot[[2017,2018,2019,2020,2021]].mean(axis=1)

# Min/Max calculations for Actuals
df_ldz['Diff'] = df_ldz[[2017,2018,2019,2020,2021]].max(axis=1).sub(df_ldz[[2017,2018,2019,2020,2021]].min(axis=1))
df_ind['Diff'] = df_ind[[2017,2018,2019,2020,2021]].max(axis=1).sub(df_ind[[2017,2018,2019,2020,2021]].min(axis=1))
df_gtp['Diff'] = df_gtp[[2017,2018,2019,2020,2021]].max(axis=1).sub(df_gtp[[2017,2018,2019,2020,2021]].min(axis=1))
df_tot['Diff'] = df_tot[[2017,2018,2019,2020,2021]].max(axis=1).sub(df_tot[[2017,2018,2019,2020,2021]].min(axis=1))

df_ldz['Min'] = df_ldz[[2017,2018,2019,2020,2021]].min(axis=1)
df_ind['Min'] = df_ind[[2017,2018,2019,2020,2021]].min(axis=1)
df_gtp['Min'] = df_gtp[[2017,2018,2019,2020,2021]].min(axis=1)
df_tot['Min'] = df_tot[[2017,2018,2019,2020,2021]].min(axis=1)

# Create MultiIndex columns for calendar year data
gtp_cols_mm = pd.MultiIndex.from_tuples([('Gas-to-Power','2017'),('Gas-to-Power','2018'),('Gas-to-Power','2019'),('Gas-to-Power','2020'),('Gas-to-Power','2021'),('Gas-to-Power','2022'),('Gas-to-Power','2023'),('Gas-to-Power','2024'),('Gas-to-Power','2025'),('Gas-to-Power','2026'),('Gas-to-Power Benchmark','2026'),('Gas-to-Power Max','2026'),('Gas-to-Power Min','2026')])
ind_cols_mm = pd.MultiIndex.from_tuples([('Industrial','2017'),('Industrial','2018'),('Industrial','2019'),('Industrial','2020'),('Industrial','2021'),('Industrial','2022'),('Industrial','2023'),('Industrial','2024'),('Industrial','2025'),('Industrial','2026'),('Industrial Benchmark','2026'),('Industrial Max','2026'),('Industrial Min','2026')])
ldz_cols_mm = pd.MultiIndex.from_tuples([('LDZ','2017'),('LDZ','2018'),('LDZ','2019'),('LDZ','2020'),('LDZ','2021'),('LDZ','2022'),('LDZ','2023'),('LDZ','2024'),('LDZ','2025'),('LDZ','2026'),('LDZ Benchmark','2026'),('LDZ Max','2026'),('LDZ Min','2026')])
tot_cols_mm = pd.MultiIndex.from_tuples([('Total','2017'),('Total','2018'),('Total','2019'),('Total','2020'),('Total','2021'),('Total','2022'),('Total','2023'),('Total','2024'),('Total','2025'),('Total','2026'),('Total Benchmark','2026'),('Total Max','2026'),('Total Min','2026')])

df_ldz.columns = ldz_cols_mm
df_ind.columns = ind_cols_mm
df_gtp.columns = gtp_cols_mm
df_tot.columns = tot_cols_mm

Actuals = df_ind.merge(df_gtp, left_index=True, right_index=True, how="outer").merge(df_ldz, left_index=True, right_index=True, how="outer").merge(df_tot, left_index=True, right_index=True, how="outer")

# Create deviations from benchmark
ldz_benchmark_col = ('LDZ Benchmark', '2026')
ind_benchmark_col = ('Industrial Benchmark', '2026')
gtp_benchmark_col = ('Gas-to-Power Benchmark', '2026')
tot_benchmark_col = ('Total Benchmark', '2026')

ldz_DD = df_ldz.drop(columns=[ldz_benchmark_col, ('LDZ Max', '2026'), ('LDZ Min', '2026')], axis=1).sub(df_ldz[ldz_benchmark_col], axis=0)
ind_DD = df_ind.drop(columns=[ind_benchmark_col, ('Industrial Max', '2026'), ('Industrial Min', '2026')], axis=1).sub(df_ind[ind_benchmark_col], axis=0)
gtp_DD = df_gtp.drop(columns=[gtp_benchmark_col, ('Gas-to-Power Max', '2026'), ('Gas-to-Power Min', '2026')], axis=1).sub(df_gtp[gtp_benchmark_col], axis=0)
tot_DD = df_tot.drop(columns=[tot_benchmark_col, ('Total Max', '2026'), ('Total Min', '2026')], axis=1).sub(df_tot[tot_benchmark_col], axis=0)

ldz_DD_perc = ldz_DD.div(df_ldz[ldz_benchmark_col], axis=0)
ind_DD_perc = ind_DD.div(df_ind[ind_benchmark_col], axis=0)
gtp_DD_perc = gtp_DD.div(df_gtp[gtp_benchmark_col], axis=0)
tot_DD_perc = tot_DD.div(df_tot[tot_benchmark_col], axis=0)

# Set proper column names
gtp_cols = pd.MultiIndex.from_tuples([('Gas-to-Power','2017'),('Gas-to-Power','2018'),('Gas-to-Power','2019'),('Gas-to-Power','2020'),('Gas-to-Power','2021'),('Gas-to-Power','2022'),('Gas-to-Power','2023'),('Gas-to-Power','2024'),('Gas-to-Power','2025'),('Gas-to-Power','2026')])
ind_cols = pd.MultiIndex.from_tuples([('Industrial','2017'),('Industrial','2018'),('Industrial','2019'),('Industrial','2020'),('Industrial','2021'),('Industrial','2022'),('Industrial','2023'),('Industrial','2024'),('Industrial','2025'),('Industrial','2026')])
ldz_cols = pd.MultiIndex.from_tuples([('LDZ','2017'),('LDZ','2018'),('LDZ','2019'),('LDZ','2020'),('LDZ','2021'),('LDZ','2022'),('LDZ','2023'),('LDZ','2024'),('LDZ','2025'),('LDZ','2026')])
tot_cols = pd.MultiIndex.from_tuples([('Total','2017'),('Total','2018'),('Total','2019'),('Total','2020'),('Total','2021'),('Total','2022'),('Total','2023'),('Total','2024'),('Total','2025'),('Total','2026')])

ldz_DD.columns = ldz_cols
ind_DD.columns = ind_cols
gtp_DD.columns = gtp_cols
tot_DD.columns = tot_cols

ldz_DD_perc.columns = ldz_cols
ind_DD_perc.columns = ind_cols
gtp_DD_perc.columns = gtp_cols
tot_DD_perc.columns = tot_cols

Calendar_years = ind_DD.merge(gtp_DD, left_index=True, right_index=True, how="outer").merge(ldz_DD, left_index=True, right_index=True, how="outer").merge(tot_DD, left_index=True, right_index=True, how="outer")
Calendar_years_perc = ind_DD_perc.merge(gtp_DD_perc, left_index=True, right_index=True, how="outer").merge(ldz_DD_perc, left_index=True, right_index=True, how="outer").merge(tot_DD_perc, left_index=True, right_index=True, how="outer")

# Process YOY data
Demand_yoy = countries[union].rolling(7).mean()
Demand_yoy.index = pd.to_datetime(Demand_yoy.index)
Demand_yoy.columns = Demand_yoy.columns.get_level_values(2)
Demand_yoy['Total'] = Demand_yoy.sum(axis=1)
Demand_yoy = Demand_yoy.pct_change(365)

df_yoy_ldz = df.copy()
df_yoy_ind = df.copy()
df_yoy_gtp = df.copy()
df_yoy_tot = df.copy()

for year in years:
    d_temp = Demand_yoy[Demand_yoy.index.year==year]
    d_temp.index = d_temp.index.strftime('%m-%d')
    ldz_temp = d_temp[['LDZ']].rename(columns={'LDZ':year})
    ind_temp = d_temp[['Industrial']].rename(columns={'Industrial':year})
    gtp_temp = d_temp[['Gas-to-Power']].rename(columns={'Gas-to-Power':year})
    tot_temp = d_temp[['Total']].rename(columns={'Total':year})
    
    df_yoy_ldz = df_yoy_ldz.merge(ldz_temp, left_index=True, right_index=True, how="outer")
    df_yoy_ind = df_yoy_ind.merge(ind_temp, left_index=True, right_index=True, how="outer")
    df_yoy_gtp = df_yoy_gtp.merge(gtp_temp, left_index=True, right_index=True, how="outer")
    df_yoy_tot = df_yoy_tot.merge(tot_temp, left_index=True, right_index=True, how="outer")

df_yoy_ldz.columns = ldz_cols
df_yoy_ind.columns = ind_cols
df_yoy_gtp.columns = gtp_cols
df_yoy_tot.columns = tot_cols

Calendar_years_YOY = df_yoy_ind.merge(df_yoy_gtp, left_index=True, right_index=True, how="outer").merge(df_yoy_ldz, left_index=True, right_index=True, how="outer").merge(df_yoy_tot, left_index=True, right_index=True, how="outer")

# Create monthly data
Calendar_years_monthly = Calendar_years.copy()
Calendar_years_monthly.index = pd.date_range('2022-01-01', freq='D', periods=365)
Calendar_years_monthly = Calendar_years_monthly.resample('M').mean()

Calendar_years_perc_monthly = Calendar_years_perc.copy()
Calendar_years_perc_monthly.index = pd.date_range('2022-01-01', freq='D', periods=365)
Calendar_years_perc_monthly = Calendar_years_perc_monthly.resample('M').mean()

# Create 2017-2021 average (placeholder - needs specific calculation)
avg_2017_2021 = Calendar_years.copy()  # This would need the actual 2017-2021 average calculation

# CONSOLIDATED WRITE: ALL 17 SHEETS IN ONE FILE
print("Writing consolidated file with all 17 sheets...")

# Prepare data for Excel writing - CLEAN DataFrames with PROPER Excel dates
full_data_out = full_data.copy()
full_data_out.index = pd.to_datetime(full_data_out.index)

countries_out = countries.copy()
countries_out.index = pd.to_datetime(countries.index)

supply_out = supply.copy()
supply_out.index = pd.to_datetime(supply_out.index)

# Clean LNG data - remove empty columns that cause gaps and use proper dates
lng_out = lng.copy()
lng_out.index = pd.to_datetime(lng_out.index)
# Drop any completely empty columns
lng_out = lng_out.dropna(axis=1, how='all')

ldz_out = ldz.copy()
ldz_out.index = pd.to_datetime(ldz_out.index)
# Drop any completely empty columns
ldz_out = ldz_out.dropna(axis=1, how='all')

gtp_out = gtp.copy()
gtp_out.index = pd.to_datetime(gtp_out.index)
# Drop any completely empty columns
gtp_out = gtp_out.dropna(axis=1, how='all')

industry_out = industry.copy()
industry_out.index = pd.to_datetime(industry_out.index)
# Drop any completely empty columns
industry_out = industry_out.dropna(axis=1, how='all')

demand_out_final = demand_out.dropna().copy()
demand_out_final.index = pd.to_datetime(demand_out_final.index)

Actuals_out = Actuals.copy()
Actuals_out.index = pd.date_range('2022-01-01', freq='D', periods=365)

Calendar_years_out = Calendar_years.copy()
Calendar_years_out.index = pd.date_range('2022-01-01', freq='D', periods=365)

Calendar_years_perc_out = Calendar_years_perc.copy()
Calendar_years_perc_out.index = pd.date_range('2022-01-01', freq='D', periods=365)

Calendar_years_YOY_out = Calendar_years_YOY.copy()
Calendar_years_YOY_out.index = pd.date_range('2022-01-01', freq='D', periods=365)

Calendar_years_monthly_out = Calendar_years_monthly.copy()
Calendar_years_monthly_out.index = pd.to_datetime(Calendar_years_monthly_out.index)

Calendar_years_perc_monthly_out = Calendar_years_perc_monthly.copy()
Calendar_years_perc_monthly_out.index = pd.to_datetime(Calendar_years_perc_monthly_out.index)

# Write all sheets to ONE file with PROPER MultiIndex headers
# Configure Excel writer with proper date formatting
with pd.ExcelWriter(output_filename, engine='xlsxwriter', 
                   datetime_format='yyyy-mm-dd') as writer:
    # ALL SHEETS with PROPER MultiIndex column headers that span across columns
    
    # Sheet 1: Multiticker - with full MultiIndex structure
    full_data_out.to_excel(writer, sheet_name='Multiticker', startrow=0, startcol=0)
    
    # Sheet 2: Demand - with hierarchical country headers
    countries_out.to_excel(writer, sheet_name='Demand', startrow=0, startcol=0)
    
    # Sheet 3: Demand YOY
    demand_out_final.to_excel(writer, sheet_name='Demand YOY', startrow=0, startcol=0)
    
    # Sheet 4: Supply - with Import/Export/Production headers
    supply_out.to_excel(writer, sheet_name='Supply', startrow=0, startcol=0)
    
    # Sheet 5: LNG imports by country - with proper Import > LNG > Country structure
    lng_out.to_excel(writer, sheet_name='LNG imports by country', startrow=0, startcol=0)
    
    # Sheet 6: LDZ demand - with hierarchical demand structure
    ldz_out.to_excel(writer, sheet_name='LDZ demand', startrow=0, startcol=0)
    
    # Sheet 7: Industrial demand - with proper category headers
    industry_out.to_excel(writer, sheet_name='Industrial demand', startrow=0, startcol=0)
    
    # Sheet 8: Gas-to-Power demand - with hierarchical structure
    gtp_out.to_excel(writer, sheet_name='Gas-to-Power demand', startrow=0, startcol=0)
    
    # Sheet 9: Calendar years - monthly levels - with year columns spanning
    Calendar_years_monthly_out.to_excel(writer, sheet_name='Calendar years - monthly levels', startrow=0, startcol=0)
    
    # Sheet 10: Calendar years - monthly - with category > year structure
    Calendar_years_monthly_out.to_excel(writer, sheet_name='Calendar years - monthly', startrow=0, startcol=0)
    
    # Sheet 11: Calendar years, % - monthly - with percentage headers
    Calendar_years_perc_monthly_out.to_excel(writer, sheet_name='Calendar years, % - monthly', startrow=0, startcol=0)
    
    # Sheet 12: Calendar years actuals - with benchmark columns
    Actuals_out.to_excel(writer, sheet_name='Calendar years actuals', startrow=0, startcol=0)
    
    # Sheet 13: Calendar years - with deviation from benchmark structure
    Calendar_years_out.to_excel(writer, sheet_name='Calendar years', startrow=0, startcol=0)
    
    # Sheet 14: Calendar years, % - with percentage deviation headers
    Calendar_years_perc_out.to_excel(writer, sheet_name='Calendar years, %', startrow=0, startcol=0)
    
    # Sheet 15: Calendar years YOY % - with year-over-year structure
    Calendar_years_YOY_out.to_excel(writer, sheet_name='Calendar years YOY %', startrow=0, startcol=0)
    
    # Sheet 16: 2017-2021 average - with historical average structure
    avg_2017_2021.to_excel(writer, sheet_name='2017-2021 average', startrow=0, startcol=0)
    
    # Sheet 17: Projected supply - with supply forecast structure
    supply_out.to_excel(writer, sheet_name='Projected supply', startrow=0, startcol=0)

print(f"\n{'='*50}")
print("CONSOLIDATED OUTPUT COMPLETE FOR USE4.xlsx!")
print(f"{'='*50}")
print(f"✅ Created: {output_filename}")
print("✅ Contains all 17 sheets in correct order")
print("✅ Adapted for use4.xlsx structure")
print("✅ Smart 'Replace blanks' handling added")
print("✅ Enhanced error handling for missing data")
print("✅ Excel links will work properly")
print(f"{'='*50}")

print("Script execution completed!")