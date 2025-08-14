# -*- coding: utf-8 -*-
"""
Created on Fri Feb 24 14:08:02 2023
@author: AD08394
OPTIMIZED VERSION - Speed improvements for Excel operations
"""

import numpy as np
import os
import pandas as pd
import datetime
import xlrd
from contextlib import ExitStack

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

# OPTIMIZED: Batch Excel writing function
def batch_write_excel(filename, data_dict, mode='w'):
    """
    Write multiple DataFrames to Excel sheets in a single operation
    data_dict: {sheet_name: (dataframe, start_row, start_col)}
    """
    temp_filename = filename.replace('.xlsx', '_temp.xlsx').replace('.xlsm', '_temp.xlsm')
    
    with pd.ExcelWriter(temp_filename, engine='xlsxwriter') as writer:
        for sheet_name, (df, start_row, start_col) in data_dict.items():
            # Convert datetime index to Excel dates if needed
            if hasattr(df.index, 'dtype') and 'datetime' in str(df.index.dtype):
                df_copy = df.copy()
                df_copy.index = to_excel_dates(pd.to_datetime(df_copy.index))
            else:
                df_copy = df
            
            df_copy.to_excel(writer, sheet_name=sheet_name, 
                           startrow=start_row-1, startcol=start_col-1)
    
    print(f"Batch written to: {temp_filename}")

#%% SET PATH FOR DATA SHEET
os.chdir('G:\Commodity Research\Gas')
data_path = 'G://Commodity Research//Gas//data//use3.xlsm'
filename = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'

#%% Load data, get demand components
print("Loading Bloomberg data...")
dataset = pd.read_excel(data_path, sheet_name='TickerList', skiprows=8, usecols=range(1,9))
tickers = list(set(dataset.Ticker.to_list()))
data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)

print("Processing ticker data...")
full_data = pd.DataFrame()
for row in range(len(dataset)):
    information = dataset.loc[row]
    try:
        temp_df = data[[information.Ticker]] * information['Normalization factor']
    except:
        temp_df = pd.DataFrame(columns=[information.Ticker], index=data.index)

    if not information['Replace blanks with #N/A'] == 'Y':
        temp_df[temp_df.isna().values] = 0
    
    temp_df.columns = pd.MultiIndex.from_tuples([(information.Description, information['Replace blanks with #N/A'], 
                                                 information.Category, information['Region from'], 
                                                 information['Region to'], information.Ticker, information.Ticker)])
    
    if row == 0:
        full_data = temp_df
    else:
        full_data = pd.merge(full_data, temp_df, left_index=True, right_index=True, how="outer")

# Data fixes
full_data.iloc[1074:1076, 103] = 0
full_data.iloc[2892:, 27] = 0
full_data.iloc[2892:, 29] = 0
full_data.iloc[3122, 39] = (full_data.iloc[3121, 39] + full_data.iloc[3123, 39]) / 2
full_data.iloc[3103, 128] = 25

# Preserve the last row and interpolate only on the t-1 rows
full_data = full_data.copy()
full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')

full_data.index = pd.to_datetime(full_data.index)
full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]

# OPTIMIZED: Prepare first batch of data for Excel writing
print("Preparing main data for Excel...")
full_data_out = full_data.copy()
full_data_out.index = to_excel_dates(pd.to_datetime(full_data_out.index))

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

# OPTIMIZED: Prepare demand component data for batch writing
ldz_out = ldz.copy()
gtp_out = gtp.copy()
industry_out = industry.copy()

ldz_out.index = to_excel_dates(pd.to_datetime(ldz_out.index))
gtp_out.index = to_excel_dates(pd.to_datetime(gtp_out.index))
industry_out.index = to_excel_dates(pd.to_datetime(industry_out.index))

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

indX = ('Brandov NordStream OPAL Daily Gas Flows (Germany to Czech Republic) (GWh/day)',
        'Y', 'Import', 'Russia', 'Germany', 'ONTGDCDF Index', 'ONTGDCDF Index')
full_data.loc[full_data.index.year > 2023, indX] = 0

for a, b, c in zip(demand_1, demand_2, demand_3):
    if b == 'LNG':
        l = full_data.iloc[:,(full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
        supply.iloc[:, (supply.columns.get_level_values(0)==a) & (supply.columns.get_level_values(1)==b)] = l
    else:
        l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
        supply.iloc[:, (supply.columns.get_level_values(0)==a) & (supply.columns.get_level_values(2)==c) & (supply.columns.get_level_values(1)==b)] = l

supply[pd.MultiIndex.from_tuples([('Supply','','Total')])] = pd.DataFrame(supply.sum(axis=1, skipna=False))

#%% Create a new dataframe with days-years
print("Processing calendar year data...")
union = pd.MultiIndex.from_tuples([('','','Gas-to-Power'),('','','Industrial'),('','','LDZ')])
Demand = countries[union]
Demand.index = pd.to_datetime(Demand.index)
Demand.columns = Demand.columns.get_level_values(2)

years = [2017,2018,2019,2020,2021,2022,2023,2024,2025,2026]
df = pd.DataFrame(index=pd.date_range('2022-01-01', freq='D', periods=365))
df.index = df.index.strftime('%m-%d')
df_ldz = df.copy()
df_ind = df.copy()
df_gtp = df.copy()
df_tot = df.copy()

for year in years:
    d_temp = Demand[Demand.index.year==year]
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

ldz_DD = df_ldz.drop(columns=['Benchmark'], axis=1).sub(df_ldz['Benchmark'], axis=0)
ind_DD = df_ind.drop(columns=['Benchmark'], axis=1).sub(df_ind['Benchmark'], axis=0)
gtp_DD = df_gtp.drop(columns=['Benchmark'], axis=1).sub(df_gtp['Benchmark'], axis=0)
tot_DD = df_tot.drop(columns=['Benchmark'], axis=1).sub(df_tot['Benchmark'], axis=0)

ldz_DD_perc = ldz_DD.div(df_ldz['Benchmark'], axis=0)
ind_DD_perc = ind_DD.div(df_ind['Benchmark'], axis=0)
gtp_DD_perc = gtp_DD.div(df_gtp['Benchmark'], axis=0)
tot_DD_perc = tot_DD.div(df_tot['Benchmark'], axis=0)

# Create MultiIndex columns
gtp_cols = pd.MultiIndex.from_tuples([('Gas-to-Power','2017'),('Gas-to-Power','2018'),('Gas-to-Power','2019'),('Gas-to-Power','2020'),('Gas-to-Power','2021'),('Gas-to-Power','2022'),('Gas-to-Power','2023'),('Gas-to-Power','2024'),('Gas-to-Power','2025'),('Gas-to-Power','2026')])
ind_cols = pd.MultiIndex.from_tuples([('Industrial','2017'),('Industrial','2018'),('Industrial','2019'),('Industrial','2020'),('Industrial','2021'),('Industrial','2022'),('Industrial','2023'),('Industrial','2024'),('Industrial','2025'),('Industrial','2026')])
ldz_cols = pd.MultiIndex.from_tuples([('LDZ','2017'),('LDZ','2018'),('LDZ','2019'),('LDZ','2020'),('LDZ','2021'),('LDZ','2022'),('LDZ','2023'),('LDZ','2024'),('LDZ','2025'),('LDZ','2026')])
tot_cols = pd.MultiIndex.from_tuples([('Total','2017'),('Total','2018'),('Total','2019'),('Total','2020'),('Total','2021'),('Total','2022'),('Total','2023'),('Total','2024'),('Total','2025'),('Total','2026')])

gtp_cols_long = pd.MultiIndex.from_tuples([('Gas-to-Power','2022'),('Gas-to-Power','2023'),('Gas-to-Power','2024'),('Gas-to-Power','2025'),('Gas-to-Power','2026')])
ind_cols_long = pd.MultiIndex.from_tuples([('Industrial','2022'),('Industrial','2023'),('Industrial','2024'),('Industrial','2025'),('Industrial','2026')])
ldz_cols_long = pd.MultiIndex.from_tuples([('LDZ','2022'),('LDZ','2023'),('LDZ','2024'),('LDZ','2025'),('LDZ','2026')])
tot_cols_long = pd.MultiIndex.from_tuples([('Total','2022'),('Total','2023'),('Total','2024'),('Total','2025'),('Total','2026')])

ldz_DD.columns = ldz_cols
ind_DD.columns = ind_cols
gtp_DD.columns = gtp_cols
tot_DD.columns = tot_cols

ldz_DD_perc.columns = ldz_cols
ind_DD_perc.columns = ind_cols
gtp_DD_perc.columns = gtp_cols
tot_DD_perc.columns = tot_cols

DD = ind_DD.merge(gtp_DD, left_index=True, right_index=True, how="outer").merge(ldz_DD, left_index=True, right_index=True, how="outer").merge(tot_DD, left_index=True, right_index=True, how="outer")
DD_perc = ind_DD_perc.merge(gtp_DD_perc, left_index=True, right_index=True, how="outer").merge(ldz_DD_perc, left_index=True, right_index=True, how="outer").merge(tot_DD_perc, left_index=True, right_index=True, how="outer")

# Fix: Create proper date index instead of converting "%m-%d" strings
DD.index = pd.date_range('2022-01-01', freq='D', periods=365)
DD_perc.index = pd.date_range('2022-01-01', freq='D', periods=365)

# Long format data
index_long = pd.date_range('2022-01-01', freq='D', periods=1826)
index_long = index_long[~((index_long.month == 2) & (index_long.day == 29))]

ldz_DD_long = pd.DataFrame({'LDZ DD': ldz_DD[ldz_cols_long].T.values.ravel()}, index=index_long)
ind_DD_long = pd.DataFrame({'Ind DD': ind_DD[ind_cols_long].T.values.ravel()}, index=index_long)
gtp_DD_long = pd.DataFrame({'GtP DD': gtp_DD[gtp_cols_long].T.values.ravel()}, index=index_long)
tot_DD_long = pd.DataFrame({'Total DD': tot_DD[tot_cols_long].T.values.ravel()}, index=index_long)

ldz_DD_perc_long = pd.DataFrame({'LDZ DD': ldz_DD_perc[ldz_cols_long].T.values.ravel()}, index=index_long)
ind_DD_perc_long = pd.DataFrame({'Ind DD': ind_DD_perc[ind_cols_long].T.values.ravel()}, index=index_long)
gtp_DD_perc_long = pd.DataFrame({'GtP DD': gtp_DD_perc[gtp_cols_long].T.values.ravel()}, index=index_long)
tot_DD_perc_long = pd.DataFrame({'Total DD': tot_DD_perc[tot_cols_long].T.values.ravel()}, index=index_long)

DD_long = ind_DD_long.merge(gtp_DD_long, left_index=True, right_index=True, how="outer").merge(ldz_DD_long, left_index=True, right_index=True, how="outer").merge(tot_DD_long, left_index=True, right_index=True, how="outer")
DD_perc_long = ind_DD_perc_long.merge(gtp_DD_perc_long, left_index=True, right_index=True, how="outer").merge(ldz_DD_perc_long, left_index=True, right_index=True, how="outer").merge(tot_DD_perc_long, left_index=True, right_index=True, how="outer")

DD_long.index = pd.to_datetime(DD_long.index)
DD_perc_long.index = pd.to_datetime(DD_perc_long.index)

DD_m = DD.resample('M').mean()
DD_perc_m = DD_perc.resample('M').mean()

DD_cum = DD_long.copy()
DD_cum_m = DD_cum.resample('M').sum()
DD_cum_m = DD_cum_m.cumsum()

DD_long_m = DD_long.resample('M').mean()
DD_perc_long_m = DD_perc_long.resample('M').mean()

# OPTIMIZED: Batch write calendar year data
print("Writing calendar year data to Excel...")
calendar_data = {
    'Calendar years - monthly': (DD_m, 2, 3),
    'Calendar years, % - monthly': (DD_perc_m, 2, 3),
}

batch_write_excel(filename, calendar_data)

# OPTIMIZED: Write additional monthly data to separate temp files to avoid merged cell conflicts
print("Writing additional monthly data...")
try:
    monthly_temp_filename = filename.replace('.xlsx', '_monthly_temp.xlsx').replace('.xlsm', '_monthly_temp.xlsm')
    
    with pd.ExcelWriter(monthly_temp_filename, engine='xlsxwriter') as writer:
        # Additional monthly data
        DD_long_m_copy = DD_long_m.copy()
        DD_long_m_copy.index = to_excel_dates(pd.to_datetime(DD_long_m_copy.index))
        DD_long_m_copy.to_excel(writer, sheet_name='Calendar years - monthly', startrow=1, startcol=19)
        
        DD_perc_long_m_copy = DD_perc_long_m.copy()
        DD_perc_long_m_copy.index = to_excel_dates(pd.to_datetime(DD_perc_long_m_copy.index))
        DD_perc_long_m_copy.to_excel(writer, sheet_name='Calendar years, % - monthly', startrow=1, startcol=19)
        
        # Cumulative data
        DD_cum_m_copy = DD_cum_m.copy()
        DD_cum_m_copy.index = to_excel_dates(pd.to_datetime(DD_cum_m_copy.index))
        DD_cum_m_copy.to_excel(writer, sheet_name='Calendar years - cumulative', startrow=1, startcol=1)
    
    print(f"Additional monthly data saved to: {monthly_temp_filename}")
except Exception as e:
    print(f"Warning: Could not write additional monthly data: {e}")

#%% Rolling average data processing
print("Processing rolling average data...")
union = pd.MultiIndex.from_tuples([('','','LDZ')]).union(pd.MultiIndex.from_tuples([('','','Gas-to-Power')])).union(pd.MultiIndex.from_tuples([('','','Industrial')]))

Demand = countries[union].rolling(7).mean()
Demand.index = pd.to_datetime(Demand.index)
Demand.columns = Demand.columns.get_level_values(2)

# Repeat the same processing for rolling average...
# (Similar code structure as above but with rolling data)

years = [2017,2018,2019,2020,2021,2022,2023,2024,2025,2026]
df = pd.DataFrame(index=pd.date_range('2022-01-01', freq='D', periods=365))
df.index = df.index.strftime('%m-%d')
df_ldz = df.copy()
df_ind = df.copy()
df_gtp = df.copy()
df_tot = df.copy()

for year in years:
    d_temp = Demand[Demand.index.year==year]
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

# Min/Max calculations
df_ldz['Diff'] = df_ldz[[2017,2018,2019,2020,2021]].max(axis=1).sub(df_ldz[[2017,2018,2019,2020,2021]].min(axis=1))
df_ind['Diff'] = df_ind[[2017,2018,2019,2020,2021]].max(axis=1).sub(df_ind[[2017,2018,2019,2020,2021]].min(axis=1))
df_gtp['Diff'] = df_gtp[[2017,2018,2019,2020,2021]].max(axis=1).sub(df_gtp[[2017,2018,2019,2020,2021]].min(axis=1))
df_tot['Diff'] = df_tot[[2017,2018,2019,2020,2021]].max(axis=1).sub(df_tot[[2017,2018,2019,2020,2021]].min(axis=1))

df_ldz['Min'] = df_ldz[[2017,2018,2019,2020,2021]].min(axis=1)
df_ind['Min'] = df_ind[[2017,2018,2019,2020,2021]].min(axis=1)
df_gtp['Min'] = df_gtp[[2017,2018,2019,2020,2021]].min(axis=1)
df_tot['Min'] = df_tot[[2017,2018,2019,2020,2021]].min(axis=1)

gtp_cols_mm = pd.MultiIndex.from_tuples([('Gas-to-Power','2017'),('Gas-to-Power','2018'),('Gas-to-Power','2019'),('Gas-to-Power','2020'),('Gas-to-Power','2021'),('Gas-to-Power','2022'),('Gas-to-Power','2023'),('Gas-to-Power','2024'),('Gas-to-Power','2025'),('Gas-to-Power','2026'),('Gas-to-Power Benchmark','2026'),('Gas-to-Power Max','2026'),('Gas-to-Power Min','2026')])
ind_cols_mm = pd.MultiIndex.from_tuples([('Industrial','2017'),('Industrial','2018'),('Industrial','2019'),('Industrial','2020'),('Industrial','2021'),('Industrial','2022'),('Industrial','2023'),('Industrial','2024'),('Industrial','2025'),('Industrial','2026'),('Industrial Benchmark','2026'),('Industrial Max','2026'),('Industrial Min','2026')])
ldz_cols_mm = pd.MultiIndex.from_tuples([('LDZ','2017'),('LDZ','2018'),('LDZ','2019'),('LDZ','2020'),('LDZ','2021'),('LDZ','2022'),('LDZ','2023'),('LDZ','2024'),('LDZ','2025'),('LDZ','2026'),('LDZ Benchmark','2026'),('LDZ Max','2026'),('LDZ Min','2026')])
tot_cols_mm = pd.MultiIndex.from_tuples([('Total','2017'),('Total','2018'),('Total','2019'),('Total','2020'),('Total','2021'),('Total','2022'),('Total','2023'),('Total','2024'),('Total','2025'),('Total','2026'),('Total Benchmark','2026'),('Total Max','2026'),('Total Min','2026')])

df_ldz.columns = ldz_cols_mm
df_ind.columns = ind_cols_mm
df_gtp.columns = gtp_cols_mm
df_tot.columns = tot_cols_mm

Actuals = df_ind.merge(df_gtp, left_index=True, right_index=True, how="outer").merge(df_ldz, left_index=True, right_index=True, how="outer").merge(df_tot, left_index=True, right_index=True, how="outer")

ldz_DD = df_ldz.drop(columns=['Benchmark'], axis=1).sub(df_ldz['Benchmark'], axis=0)
ind_DD = df_ind.drop(columns=['Benchmark'], axis=1).sub(df_ind['Benchmark'], axis=0)
gtp_DD = df_gtp.drop(columns=['Benchmark'], axis=1).sub(df_gtp['Benchmark'], axis=0)
tot_DD = df_tot.drop(columns=['Benchmark'], axis=1).sub(df_tot['Benchmark'], axis=0)

ldz_DD_perc = ldz_DD.div(df_ldz['Benchmark'], axis=0)
ind_DD_perc = ind_DD.div(df_ind['Benchmark'], axis=0)
gtp_DD_perc = gtp_DD.div(df_gtp['Benchmark'], axis=0)
tot_DD_perc = tot_DD.div(df_tot['Benchmark'], axis=0)

ldz_DD.columns = ldz_cols
ind_DD.columns = ind_cols
gtp_DD.columns = gtp_cols
tot_DD.columns = tot_cols

ldz_DD_perc.columns = ldz_cols
ind_DD_perc.columns = ind_cols
gtp_DD_perc.columns = gtp_cols
tot_DD_perc.columns = tot_cols

DD = ind_DD.merge(gtp_DD, left_index=True, right_index=True, how="outer").merge(ldz_DD, left_index=True, right_index=True, how="outer").merge(tot_DD, left_index=True, right_index=True, how="outer")
DD_perc = ind_DD_perc.merge(gtp_DD_perc, left_index=True, right_index=True, how="outer").merge(ldz_DD_perc, left_index=True, right_index=True, how="outer").merge(tot_DD_perc, left_index=True, right_index=True, how="outer")

# OPTIMIZED: Write rolling data to separate temp file to avoid conflicts
print("Writing rolling data to Excel...")
try:
    rolling_temp_filename = filename.replace('.xlsx', '_rolling_temp.xlsx').replace('.xlsm', '_rolling_temp.xlsm')
    
    rolling_data = {
        'Calendar years actuals': (Actuals, 2, 3),
        'Calendar years': (DD, 2, 3),
        'Calendar years, %': (DD_perc, 2, 3)
    }
    
    with pd.ExcelWriter(rolling_temp_filename, engine='xlsxwriter') as writer:
        for sheet, (df, row, col) in rolling_data.items():
            df_copy = df.copy()
            df_copy.index = to_excel_dates(pd.date_range('2022-01-01', freq='D', periods=365))
            df_copy.to_excel(writer, sheet_name=sheet, startrow=row-1, startcol=col-1)
    
    print(f"Rolling data saved to: {rolling_temp_filename}")
except Exception as e:
    print(f"Warning: Could not write rolling data: {e}")

#%% YOY percentage data
print("Processing YOY percentage data...")
union = pd.MultiIndex.from_tuples([('','','LDZ')]).union(pd.MultiIndex.from_tuples([('','','Gas-to-Power')])).union(pd.MultiIndex.from_tuples([('','','Industrial')]))

Demand = countries[union].rolling(7).mean()
Demand.index = pd.to_datetime(Demand.index)
Demand.columns = Demand.columns.get_level_values(2)
Demand['Total'] = Demand.sum(axis=1)
Demand = Demand.pct_change(365)

years = [2017,2018,2019,2020,2021,2022,2023,2024,2025,2026]
df = pd.DataFrame(index=pd.date_range('2022-01-01', freq='D', periods=365))
df.index = df.index.strftime('%m-%d')
df_ldz = df.copy()
df_ind = df.copy()
df_gtp = df.copy()
df_tot = df.copy()

for year in years:
    d_temp = Demand[Demand.index.year==year]
    d_temp.index = d_temp.index.strftime('%m-%d')
    ldz_temp = d_temp[['LDZ']].rename(columns={'LDZ':year})
    ind_temp = d_temp[['Industrial']].rename(columns={'Industrial':year})
    gtp_temp = d_temp[['Gas-to-Power']].rename(columns={'Gas-to-Power':year})
    tot_temp = d_temp[['Total']].rename(columns={'Total':year})
    
    df_ldz = df_ldz.merge(ldz_temp, left_index=True, right_index=True, how="outer")
    df_ind = df_ind.merge(ind_temp, left_index=True, right_index=True, how="outer")
    df_gtp = df_gtp.merge(gtp_temp, left_index=True, right_index=True, how="outer")
    df_tot = df_tot.merge(tot_temp, left_index=True, right_index=True, how="outer")

df_ldz.columns = ldz_cols
df_ind.columns = ind_cols
df_gtp.columns = gtp_cols
df_tot.columns = tot_cols

Actuals_t = df_ind.merge(df_gtp, left_index=True, right_index=True, how="outer").merge(df_ldz, left_index=True, right_index=True, how="outer").merge(df_tot, left_index=True, right_index=True, how="outer")

# OPTIMIZED: Write YOY data to separate temp file to avoid conflicts
print("Writing YOY data to Excel...")
try:
    yoy_temp_filename = filename.replace('.xlsx', '_yoy_temp.xlsx').replace('.xlsm', '_yoy_temp.xlsm')
    
    with pd.ExcelWriter(yoy_temp_filename, engine='xlsxwriter') as writer:
        df_copy = Actuals_t.copy()
        df_copy.index = to_excel_dates(pd.date_range('2022-01-01', freq='D', periods=365))
        df_copy.to_excel(writer, sheet_name='Calendar years YOY %', startrow=1, startcol=2)
    
    print(f"YOY data saved to: {yoy_temp_filename}")
except Exception as e:
    print(f"Warning: Could not write YOY data: {e}")

#%% Simulation section
print("Processing simulation data...")
import pandas as pd

# Initialize DataFrames for supply simulation
average_supply = pd.DataFrame(columns=supply.columns, index=pd.date_range('2022-01-01', freq='D', periods=365))
new_supply = pd.DataFrame(columns=supply.columns, index=pd.date_range('2022-01-01', freq='D', periods=365 * 6))

columns = average_supply.columns
new_supply = new_supply[~((new_supply.index.month == 2) & (new_supply.index.day == 29))]
new_supply = new_supply[new_supply.index.year < 2027]

deviation_supply = supply[supply.index.year == 2022]
years = [2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026]
df = pd.DataFrame(index=pd.date_range('2022-01-01', freq='D', periods=365))

# Calculate average supply and deviations
for col in average_supply.columns:
    column = supply[col]
    df_t = df.copy()
    
    for year in range(2017, 2022):
        temp = column[column.index.year == year]
        temp.index = pd.date_range('2022-01-01', freq='D', periods=365)
        df_t = df_t.merge(temp.rename(year), left_index=True, right_index=True, how="outer")
    
    average_supply[col] = df_t[[2017, 2018, 2019, 2020, 2021]].mean(axis=1)
    deviation_supply.loc[:, col] = deviation_supply[col].sub(average_supply[col], axis=0)

# Supply forecasting
new_supply.iloc[:, :2] = 0.25  # Russian supply assumption

# Norwegian supply adjustment
try:
    nor_main = pd.read_excel('G:\\Commodity Research\\Gas\\data\\Norwegianmaintenance.xlsx', index_col=0)[['decline']]
    nor_main = nor_main[~((nor_main.index.month == 2) & (nor_main.index.day == 29))]
    
    multi_index = pd.MultiIndex.from_tuples([('Import', 'Norway', 'Europe')], names=['Type', 'Country', 'Region'])
    nor = pd.DataFrame(pd.concat([average_supply.iloc[:, 2]] * 5), columns=multi_index)
    nor.index = new_supply.index
    nor_main = nor_main.reindex(nor.index).fillna(0)
    nor_main[('Import', 'Norway', 'Europe')] = nor[('Import', 'Norway', 'Europe')]
    new_supply[('Import', 'Norway', 'Europe')] = nor_main.sum(axis=1) + 10.26 - 20
except:
    print("Warning: Norwegian maintenance file not found, using default values")
    new_supply[('Import', 'Norway', 'Europe')] = average_supply.iloc[:, 2]

# Dutch production forecast
nl_decline_factors = [0.6, 0.55, 0.5, 0.45, 0.4]
nl = pd.concat([
    pd.Series(supply.iloc[:, 3].loc[supply.index.year == 2022].values * factor) for factor in nl_decline_factors
], axis=0)
nl.index = new_supply.index
new_supply.iloc[:, 3] = nl.values.flatten()

# UK production forecast
uk_decline_factors = [1, 0.98, 0.96, 0.94, 0.92]
uk = pd.concat([
    pd.Series(average_supply.iloc[:, 4].values * factor) for factor in uk_decline_factors
], axis=0)
uk.index = new_supply.index
new_supply.iloc[:, 4] = uk.values.flatten()

# Denmark production set to zero
new_supply.iloc[:, 9] = 0

# LNG forecast
lng_supply = pd.DataFrame(pd.concat([average_supply.iloc[:, 5]] * 5))
lng_supply += deviation_supply.iloc[:, 5][deviation_supply.index.month > 6].mean() - 20
lng_supply.index = new_supply.index

# Future LNG imports
lng_supply.loc[(lng_supply.index.year > 2023) & (lng_supply.index.year < 2025)] += 0
lng_supply.loc[(lng_supply.index.year > 2024) & (lng_supply.index.year < 2026)] += 0
lng_supply.loc[(lng_supply.index.year > 2025) & (lng_supply.index.year < 2027)] += 30
lng_supply.loc[(lng_supply.index.year > 2026) & (lng_supply.index.year < 2028)] += 30
new_supply.iloc[:, 5] = lng_supply.values.flatten()

# Other supply sources
for col in columns[2:]:
    if col[1] in ['Netherlands', 'GB', 'LNG', 'Norway', 'Denmark']:
        continue
    expanded_supply = pd.concat([average_supply[col]] * 5).values.flatten()
    expanded_supply = pd.Series(expanded_supply, index=new_supply.index)
    mean_deviation = deviation_supply[col][deviation_supply.index.month > 6].mean()
    new_supply[col] = expanded_supply + mean_deviation

# Calculate total supply
new_supply.iloc[:, -1] = new_supply.iloc[:, :-1].sum(axis=1)

# Demand forecast
years_demand = [2022, 2023, 2024, 2025, 2026]
ind_decline = [1, 0.95, 0.90, 0.85, 0.80]
ldz_decline = [1, 0.98, 0.95, 0.92, 0.90]
gtp_decline = [1, 0.93, 0.90, 0.87, 0.85]

demand = countries['', '','Total']
average_demand = pd.DataFrame(index=new_supply.index, columns=range(1, 6))

d_ind_total = pd.Series(dtype=float)
d_ldz_total = pd.Series(dtype=float)
d_gtp_total = pd.Series(dtype=float)

# Process each year
for idx, year in enumerate(years_demand):
    dates = pd.date_range(f'{year}-01-01', f'{year}-12-31', freq='D')
    dates = dates[~((dates.month == 2) & (dates.day == 29))]
    
    # Industrial demand
    d_ind = pd.Series(Actuals[('Industrial Benchmark', '2026')].values * ind_decline[idx], index=dates)
    if d_ind_total.empty:
        d_ind_total = d_ind
    else:
        d_ind_total = pd.concat([d_ind_total, d_ind])
    
    # LDZ demand
    d_ldz = pd.Series(Actuals[('LDZ Benchmark', '2026')].values * ldz_decline[idx], index=dates)
    if d_ldz_total.empty:
        d_ldz_total = d_ldz
    else:
        d_ldz_total = pd.concat([d_ldz_total, d_ldz])
    
    # Gas-to-power demand
    d_gtp = pd.Series(Actuals[('Gas-to-Power Benchmark', '2026')].values * gtp_decline[idx], index=dates)
    if d_gtp_total.empty:
        d_gtp_total = d_gtp
    else:
        d_gtp_total = pd.concat([d_gtp_total, d_gtp])

# Combine all categories
aggregate_demand = d_ind_total + d_ldz_total + d_gtp_total
aggregate_demand = aggregate_demand.reindex(new_supply.index)

# Calculate net supply
net_supply = pd.DataFrame(index=new_supply.index)
net_supply['Net Supply'] = new_supply.iloc[:, -1] - aggregate_demand
net_supply.index = net_supply.index.strftime('%Y-%m-%d')

# OPTIMIZED: Write simulation data to separate temp file
print("Writing simulation data...")
try:
    forecasting_filename = 'G:\\Commodity Research\\Gas\\data\\Forecasting_storage_temp.xlsx'
    
    # Create day-year formatted DataFrames
    def convert_to_day_year_with_date(series, column_name, start_year=2022):
        df = series.reset_index()
        df.columns = ['date', column_name]
        df = df[~((df['date'].dt.month == 2) & (df['date'].dt.day == 29))]
        df['year'] = df['date'].dt.year
        df['day'] = df.groupby('year').cumcount() + 1
        pivot_df = df.pivot(index='day', columns='year', values=column_name)
        start_date = f"{start_year}-01-01"
        pivot_df['date'] = pd.date_range(start=start_date, periods=365)
        pivot_df.set_index('date', inplace=True)
        return pivot_df
    
    d_ind_pivot = convert_to_day_year_with_date(d_ind_total, 'Industrial Demand', start_year=2022)
    d_ldz_pivot = convert_to_day_year_with_date(d_ldz_total, 'LDZ Demand', start_year=2022)
    d_gtp_pivot = convert_to_day_year_with_date(d_gtp_total, 'Gas-to-Power Demand', start_year=2022)
    
    new_supply_filtered = new_supply[new_supply.index >= '2025-01-01']
    
    with pd.ExcelWriter(forecasting_filename, engine='xlsxwriter') as writer:
        net_supply.to_excel(writer, sheet_name='Net supply', startrow=1, startcol=1)
        d_ind_pivot.to_excel(writer, sheet_name='Industrial Demand')
        d_ldz_pivot.to_excel(writer, sheet_name='LDZ Demand')
        d_gtp_pivot.to_excel(writer, sheet_name='Gas-to-Power Demand')
        new_supply_filtered.to_excel(writer, sheet_name='New Supply')
    
    print(f"Simulation data saved to: {forecasting_filename}")
except Exception as e:
    print(f"Warning: Could not write simulation data: {e}")

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

countries_out = countries.copy()
countries_out.index = to_excel_dates(pd.to_datetime(countries.index))

supply_out = supply.copy()
supply_out.index = to_excel_dates(pd.to_datetime(supply_out.index))

lng_out = lng.copy()
lng_out.index = to_excel_dates(pd.to_datetime(lng_out.index))

# OPTIMIZED: Final batch write for main data
print("Writing main data sheets...")
main_data = {
    'Multiticker': (full_data_out, 2, 8),
    'LDZ demand': (ldz_out, 2, 4),
    'Gas-to-Power demand': (gtp_out, 2, 4),
    'Industrial demand': (industry_out, 2, 4),
    'Demand': (countries_out, 2, 4),
    'Supply': (supply_out, 2, 4),
    'LNG imports by country': (lng_out, 2, 4)
}

batch_write_excel(filename, main_data)

#%% Create balance dataframes and final processing
print("Creating balance dataframes...")

# Create demand dataframe
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

# Write demand YOY data to separate temp file to avoid conflicts
print("Writing demand YOY data...")
try:
    demand_yoy_filename = filename.replace('.xlsx', '_demand_yoy_temp.xlsx').replace('.xlsm', '_demand_yoy_temp.xlsm')
    
    with pd.ExcelWriter(demand_yoy_filename, engine='xlsxwriter') as writer:
        demand_out_copy = demand_out.dropna().copy()
        demand_out_copy.index = to_excel_dates(demand_out_copy.index)
        demand_out_copy.to_excel(writer, sheet_name='Demand YOY', startrow=1, startcol=1)
    
    print(f"Demand YOY data saved to: {demand_yoy_filename}")
except Exception as e:
    print(f"Warning: Could not write demand YOY data: {e}")

# Create supply dataframes
supp = supply.copy()
supp = supp[pd.MultiIndex.from_tuples([('Supply','','Total')])]
Supply = pd.DataFrame(supp.values, columns=['total_supply'], index=supp.index)
Supply.index = pd.to_datetime(Supply.index)
Supply = Supply[~((Supply.index.month == 2) & (Supply.index.day == 29))]

# Russia supply
supply_russia = supply[pd.MultiIndex.from_tuples([('Import','Russia','Austria'),('Import','Russia','Germany')])].sum(axis=1)
Supply_r = pd.DataFrame(supply_russia.values, columns=['total_supply_russia'], index=supply_russia.index)
Supply_r.index = pd.to_datetime(Supply_r.index)
Supply_r = Supply_r[~((Supply_r.index.month == 2) & (Supply_r.index.day == 29))]

# Add to supply dataframe
Supply = pd.concat([Supply, Supply_r], axis=1, sort=False).dropna()
Supply['total_supply_ex_russia'] = Supply['total_supply'].values - Supply['total_supply_russia'].abs().values

# LNG supply
supply_lng = lng[pd.MultiIndex.from_tuples([('Supply','Total','')])]
Supply_lng = pd.DataFrame(supply_lng.values, columns=['total_lng_supply'], index=supply_lng.index)
Supply_lng.index = pd.to_datetime(Supply_lng.index)
Supply_lng = Supply_lng[~((Supply_lng.index.month == 2) & (Supply_lng.index.day == 29))]

Supply = pd.concat([Supply, Supply_lng], axis=1, sort=False)

# Norwegian supply
supply_NO = supply[pd.MultiIndex.from_tuples([('Import','Norway','Europe')])]
Supply_NO = pd.DataFrame(supply_NO.values, columns=['total_supply_no'], index=supply_NO.index)
Supply_NO.index = pd.to_datetime(Supply_NO.index)
Supply_NO = Supply_NO[~((Supply_NO.index.month == 2) & (Supply_NO.index.day == 29))]

Supply = pd.concat([Supply, Supply_NO], axis=1, sort=False).dropna()
Supply['total_supply_ex_russia_no'] = Supply['total_supply_ex_russia'].values - Supply['total_supply_no'].values

# Create complete supply dataframe
columns = ['Russia_Austria', 'Russia_Germany', 'Norway_Europe', 'Netherlands_Netherlands', 'GB_GB', 'LNG', 'Algeria_Italy', 'Libya_Italy', 'Spain_France', 'Denmark_Germany', 'Poland_Germany', 'Hungary_Austria',
           'Slovenia_Europe', 'Austria_MAB', 'TAP_Italy', 'Austria_Austria', 'Italy_Italy', 'Germany_Germany']

total_supply = supply.iloc[:,:-1]
Supply_tot = pd.DataFrame(total_supply.values, columns=columns, index=total_supply.index)
res_idx = Supply_tot.fillna(method='bfill').dropna().index
Supply_tot = Supply_tot.loc[res_idx]
Supply_tot.index = pd.to_datetime(Supply_tot.index)
Supply_tot = Supply_tot[~((Supply_tot.index.month == 2) & (Supply_tot.index.day == 29))]
Supply_complete = Supply_tot

# Create net withdrawals dataframe
demand_1 = ['Net Withdrawals','Net Withdrawals','Net Withdrawals','Net Withdrawals','Net Withdrawals','Net Withdrawals','Net Withdrawals']
demand_2 = ['France','Belgium','Italy','Netherlands','GB','Austria','Germany']
demand_3 = ['France','Belgium','Italy','Netherlands','GB','Austria','Germany']

net_with = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

for a, b, c in zip(demand_1, demand_2, demand_3):
    l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
    net_with.iloc[:, (net_with.columns.get_level_values(0)==a) & (net_with.columns.get_level_values(2)==c) & (net_with.columns.get_level_values(1)==b)] = l

net_with[pd.MultiIndex.from_tuples([('Net Withdrawals','Total','')])] = pd.DataFrame(net_with.sum(axis=1, skipna=False))

withdrawal = net_with[('Net Withdrawals','Total','')]
Withdrawal = pd.DataFrame(withdrawal.values, columns=['net_withdrawals'], index=withdrawal.index)
Withdrawal.index = pd.to_datetime(Withdrawal.index)
Withdrawal = Withdrawal[~((Withdrawal.index.month == 2) & (Withdrawal.index.day == 29))]

# Create inventory dataframe
demand_1 = ['Inventory','Inventory','Inventory','Inventory','Inventory','Inventory','Inventory']
demand_2 = ['France','Belgium','Italy','Netherlands','GB','Austria','Germany']
demand_3 = ['France','Belgium','Italy','Netherlands','GB','Austria','Germany']

invent = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))

for a, b, c in zip(demand_1, demand_2, demand_3):
    l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=False)
    invent.iloc[:, (invent.columns.get_level_values(0)==a) & (invent.columns.get_level_values(2)==c) & (invent.columns.get_level_values(1)==b)] = l

invent[pd.MultiIndex.from_tuples([('Inventory','Total','')])] = pd.DataFrame(invent.sum(axis=1, skipna=False))

inventory = invent[pd.MultiIndex.from_tuples([('Inventory','Total','')])]
Inventory = pd.DataFrame(inventory.values, columns=['inventory_bnef'], index=inventory.index)
Inventory.index = pd.to_datetime(Inventory.index)
Inventory = Inventory[~((Inventory.index.month == 2) & (Inventory.index.day == 29))]
Inventory['delta_inventory_bnef'] = Inventory['inventory_bnef'].diff()

# Create country-specific dataframes
id_country = [0,1,2,3,4,5,6,7,8,9]
country = countries.iloc[:,id_country]
Countries = pd.DataFrame(country.values, columns=['France','Belgium','Italy','Netherlands','UK','Austria','Germany','Switzerland','Luxembourg','Ireland'], index=countries.index)
Countries.index = pd.to_datetime(Countries.index)
Countries = Countries[~((Countries.index.month == 2) & (Countries.index.day == 29))]

id_country = [0,1,2,3,4,5,6,7,8,9,10]
countries_ldz = ldz.iloc[:,id_country]
Countries_LDZ = pd.DataFrame(countries_ldz.values, columns=['France','Belgium','Italy1','Italy2','Netherlands','UK','Austria','Germany','Switzerland','Luxembourg','Ireland'], index=countries_ldz.index)
Countries_LDZ['Italy'] = Countries_LDZ[['Italy1','Italy2']].sum(axis=1).values
Countries_LDZ.drop(columns=['Italy1','Italy2'], inplace=True)
Countries_LDZ.index = pd.to_datetime(Countries_LDZ.index)

Countries_GTP = gtp.copy()
Countries_GTP.index = pd.to_datetime(Countries_GTP.index)

id_country = [0,1,2,3,8,11]
countries_ind = industry.iloc[:,id_country]
Countries_IND = pd.DataFrame(countries_ind.values, columns=['France','Belgium','Italy','UK','Germany','Netherlands'], index=countries_ind.index)
Countries_IND.dropna(inplace=True)
Countries_IND.index = pd.to_datetime(Countries_IND.index)

# OPTIMIZED: Final balance data write
print("Writing balance data to Excel...")
try:
    os.chdir('G:\Commodity Research\Gas\data')
    balance_filename = 'Balance_temp.xlsx'
    
    # Combine complete data
    Complete = pd.concat([Demand[['total_demand']], Supply, Withdrawal, Inventory], axis=1, sort=False)
    
    with pd.ExcelWriter(balance_filename, engine='xlsxwriter') as writer:
        Demand.to_excel(writer, sheet_name='Demand')
        Supply.to_excel(writer, sheet_name='Supply')
        Inventory.to_excel(writer, sheet_name='Inventory BNEF')
        Withdrawal.to_excel(writer, sheet_name='Net Withdrawals')
        Complete.to_excel(writer, sheet_name='Total')
        Countries.to_excel(writer, sheet_name='Demand by country')
        Countries_LDZ.to_excel(writer, sheet_name='LDZ demand by country')
        Countries_GTP.to_excel(writer, sheet_name='GTP demand by country')
        Countries_IND.to_excel(writer, sheet_name='IND demand by country')
        Supply_complete.to_excel(writer, sheet_name='Complete_historic_supply')
    
    print(f"Balance data saved to: {balance_filename}")
    print("\n" + "="*50)
    print("OPTIMIZATION COMPLETE!")
    print("="*50)
    print("All files have been saved with '_temp' suffix for safe testing.")
    print("Key optimizations implemented:")
    print("1. Batch Excel writing operations")
    print("2. Reduced file open/close cycles")
    print("3. Used xlsxwriter engine for better performance")
    print("4. Pre-computed all data before writing")
    print("5. Added error handling for missing files")
    print("="*50)
    
except Exception as e:
    print(f"Warning: Could not write balance data: {e}")

print("Script execution completed!")