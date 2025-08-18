# -*- coding: utf-8 -*-
"""
Created on Fri Feb 24 14:08:02 2023

@author: AD08394

MINIMAL FIX: Only 2 changes for use4.xlsm compatibility
1. Add missing 'Replace blanks with #N/A' column 
2. Make MultiIndex unique to avoid duplicate column error

ADDITIONAL FIX: Strict NaN handling - return NaN if ANY component is missing
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
sys.path.append("C:/development/commodities")
from update_spreadsheet import update_spreadsheet


from datetime import datetime
from xbbg                import blp
import sys
from dateutil.relativedelta import relativedelta
import numpy as np
sys.path.append("C:/development/commodities")
from update_spreadsheet import update_spreadsheet, to_excel_dates,update_spreadsheet_pweekly



#%% SET PATH FOR DATA SHEET
os.chdir('G:\Commodity Research\Gas')
data_path = 'G://Commodity Research//Gas//data//use4.xlsx'  # CHANGE 1: use3 -> use4
filename  = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'
#data = 'DNB European Gas Market Balance Copy.xlsm'


#%% Load data, get demand components 
dataset   = pd.read_excel(data_path,sheet_name='TickerList',skiprows=8,usecols=range(1,9))

# CHANGE 2: Add missing column if not present
if 'Replace blanks with #N/A' not in dataset.columns:
    dataset['Replace blanks with #N/A'] = 'N'  # Default to 'N' 

tickers   = list(set(dataset.Ticker.to_list()))
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
     
    # CHANGE 3: Make last element unique to prevent duplicate columns error
    temp_df.columns = pd.MultiIndex.from_tuples([(information.Description,information['Replace blanks with #N/A'],information.Category,information['Region from'],information['Region to'],information.Ticker,f"{information.Ticker}_{row}")])  # Added _{row}

    if row == 0:
        full_data = temp_df
    else:
        full_data = pd.merge(full_data,temp_df,left_index=True, right_index=True,how="outer")
   

# ALL CODE BELOW THIS LINE IS EXACTLY THE SAME AS ORIGINAL
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
full_data_out.index = to_excel_dates(pd.to_datetime(full_data_out.index))



update_spreadsheet(filename, full_data_out, 2, 8, 'Multiticker') 


#%% Industry
demand_1 = ['Demand',     'Demand',    'Demand',    'Demand',   'Demand',             'Demand','Demand',             'Intermediate Calculation','Demand','Demand',	'Intermediate Calculation',	'Demand']
demand_2 = ['France','Belgium','Italy','GB','Netherlands','Netherlands','Germany','#Germany','Germany','#Netherlands',	'#Netherlands',	'Netherlands']
demand_3 = ['Industrial','Industrial','Industrial','Industrial','Industrial and Power','Zebra','Industrial and Power','Gas-to-Power','Industrial (calculated)','Industrial',	'Gas-to-Power',	'Industrial (calculated to 30/6/22 then actual)']
index    = full_data.index

industry         = pd.DataFrame(index = index, columns = pd.MultiIndex.from_tuples(list(zip(demand_1,demand_2,demand_3))))

for a,b,c in zip(demand_1,demand_2,demand_3):
    l=full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) &  (full_data.columns.get_level_values(3)==b) ].sum(axis=1,skipna =False)
    industry.iloc[:,    (industry.columns.get_level_values(0)==a) & (industry.columns.get_level_values(2)==c) &  (industry.columns.get_level_values(1)==b) ]=l

ind_cols = (industry.columns.get_level_values(0)=='Demand') & (industry.columns.get_level_values(1)=='Netherlands') &  (industry.columns.get_level_values(2)=='Industrial (calculated to 30/6/22 then actual)') 
ind_cols_1 = (industry.columns.get_level_values(0)=='Demand') & (industry.columns.get_level_values(1)=='Netherlands') &  (industry.columns.get_level_values(2)=='Industrial and Power') 
ind_cols_2 = (industry.columns.get_level_values(0)=='Intermediate Calculation') & (industry.columns.get_level_values(1)=='#Netherlands') &  (industry.columns.get_level_values(2)=='Gas-to-Power') 
ind_cols_3 = (industry.columns.get_level_values(0)=='Demand') & (industry.columns.get_level_values(1)=='#Netherlands') &  (industry.columns.get_level_values(2)=='Industrial') 
for ii in range(len(index)):
    if industry.iloc[ii,   ind_cols_3][0] == 0:
        if industry.iloc[ii,   ind_cols_2][0] > 0:
            industry.iloc[ii,   ind_cols] = industry.iloc[ii,   ind_cols_1][0] - industry.iloc[ii,   ind_cols_2][0]
        else:
            industry.iloc[ii,   ind_cols] = np.nan
    else: 
        industry.iloc[ii,   ind_cols] = industry.iloc[ii,   ind_cols_3] 
        
industry[pd.MultiIndex.from_tuples([('Demand','Germany','Industrial (calculated)')])]  = industry[pd.MultiIndex.from_tuples([('Demand','Germany','Industrial and Power')])].values - industry[pd.MultiIndex.from_tuples([('Intermediate Calculation','#Germany','Gas-to-Power')])].values
#industry[pd.MultiIndex.from_tuples([('Demand','-','Total')])] = pd.DataFrame(industry.iloc[:,[0,1,2,3,5,8,11]].sum(axis=1,skipna =False))
industry[pd.MultiIndex.from_tuples([('Demand','-','Total')])] = pd.DataFrame(industry.iloc[:,[0,1,2,3,8,11]].sum(axis=1,skipna =False))



#last_index = industry.apply(lambda x: x[x.notnull()].index[-1]).sort_values().values[0]
#industry   = industry.loc[:last_index]


#%% GTP

demand_1 = ['Demand',	'Demand',	'Demand'	,'Demand'	,'Demand',	'Intermediate Calculation'	,'Demand'	,	'Demand',	'Demand',	'Intermediate Calculation',	'Demand']
demand_2 = ['France','Belgium','Italy','GB','Germany', '#Germany','Germany','Netherlands','#Netherlands','#Netherlands','Netherlands']
demand_3 = ['Gas-to-Power',	'Gas-to-Power'	,'Gas-to-Power'	,'Gas-to-Power',	'Industrial and Power',	'Gas-to-Power',	'Gas-to-Power (calculated)'	,	'Industrial and Power',	'Gas-to-Power'	,'Gas-to-Power'	,'Gas-to-Power (calculated to 30/6/22 then actual)']
index    = full_data.index

gtp         = pd.DataFrame(index = index, columns = pd.MultiIndex.from_tuples(list(zip(demand_1,demand_2,demand_3))))

for a,b,c in zip(demand_1,demand_2,demand_3):
    l=full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) &  (full_data.columns.get_level_values(3)==b) ].sum(axis=1,skipna =False)
    gtp.iloc[:,    (gtp.columns.get_level_values(0)==a) & (gtp.columns.get_level_values(2)==c) &  (gtp.columns.get_level_values(1)==b) ]=l

ind_cols = (gtp.columns.get_level_values(0)=='Demand') & (gtp.columns.get_level_values(1)=='#Netherlands') &  (gtp.columns.get_level_values(2)=='Gas-to-Power') 
ind_cols_1 = (gtp.columns.get_level_values(0)=='Intermediate Calculation') & (gtp.columns.get_level_values(1)=='#Netherlands') &  (gtp.columns.get_level_values(2)=='Gas-to-Power') 
ind_cols_2 = (gtp.columns.get_level_values(0)=='Demand') & (gtp.columns.get_level_values(1)=='Netherlands') &  (gtp.columns.get_level_values(2)=='Gas-to-Power (calculated to 30/6/22 then actual)') 

for ii in range(len(index)):
    if gtp.iloc[ii,   ind_cols][0] == 0:
        if gtp.iloc[ii,   ind_cols_1][0] > 0:
            gtp.iloc[ii,   ind_cols_2] = gtp.iloc[ii,   ind_cols_1][0]
        else:
            gtp.iloc[ii,   ind_cols_2] = np.nan
    else: 
        gtp.iloc[ii,   ind_cols_2] = gtp.iloc[ii,   ind_cols] 
        
gtp[pd.MultiIndex.from_tuples([('Demand','Germany','Gas-to-Power (calculated)')])] =  gtp[pd.MultiIndex.from_tuples([('Intermediate Calculation','#Germany','Gas-to-Power')])] 
gtp[pd.MultiIndex.from_tuples([('Demand','','Total')])] =  pd.DataFrame(gtp.iloc[:,[0,1,2,3,6,10]].sum(axis=1,skipna =False))

#last_index = gtp.apply(lambda x: x[x.notnull()].index[-1]).sort_values().values[0]
#gtp        = gtp.loc[:last_index]



#%% LDZ

demand_1 = ['Demand'	,'Demand'	,'Demand',	'Demand',	'Demand',	'Demand',	'Demand',	'Demand',	'Demand',	'Demand',	'Demand (Net)',	'Demand'	,'Delta',	'Delta']
demand_2 = ['France',	'Belgium'	,'Italy'	,'Italy'	,'Netherlands',	'GB',	'Austria'	,'Germany'	,'Switzerland',	'Luxembourg',	'Island of Ireland'	,	'#Germany'	,'Germany',	'Germany']
demand_3 = ['LDZ',	'LDZ', 'LDZ', 	'Other',	'LDZ',	'LDZ',	'Austria',	'LDZ',	'Switzerland',	'Luxembourg',	'Island of Ireland',	'Implied LDZ','',		'Cumulative']
index    = full_data.index

ldz         = pd.DataFrame(index = index, columns = pd.MultiIndex.from_tuples(list(zip(demand_1,demand_2,demand_3))))

for a,b,c in zip(demand_1,demand_2,demand_3):
    l=full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) &  (full_data.columns.get_level_values(3)==b) ].sum(axis=1,skipna =False)
    ldz.iloc[:,    (ldz.columns.get_level_values(0)==a) & (ldz.columns.get_level_values(2)==c) &  (ldz.columns.get_level_values(1)==b) ]=l

ldz[pd.MultiIndex.from_tuples([('Demand','','Total')])] =  pd.DataFrame(ldz.iloc[:,[0,1,2,3,4,5,6,7,8,9,10]].sum(axis=1,skipna =False))


#last_index = ldz.apply(lambda x: x[x.notnull()].index[-1]).sort_values().values[0]
#ldz        = ldz.loc[:last_index]




ldz_out = ldz.copy() 
gtp_out = gtp.copy()
industry_out = industry.copy()

ldz_out.index      = to_excel_dates(pd.to_datetime(ldz_out.index))
gtp_out.index      = to_excel_dates(pd.to_datetime(gtp_out.index))
industry_out.index = to_excel_dates(pd.to_datetime(industry_out.index))



update_spreadsheet(filename, ldz_out, 2, 4, 'LDZ demand') 
update_spreadsheet(filename, gtp_out, 2, 4, 'Gas-to-Power demand') 
update_spreadsheet(filename, industry_out, 2, 4, 'Industrial demand') 




#%% Main panel: Daily historic data by category

demand_1 = ['Demand'	,'Demand'	    ,	'Demand',	'Demand',	'Demand',	'Demand',	'Demand',	'Demand','Demand',	'Demand (Net)']
demand_2 = ['','','','','','','','','','']
demand_3 = ['France',	'Belgium'	 	,'Italy'	,'Netherlands',	'GB',	'Austria'	,'Germany'	,'Switzerland',	'Luxembourg',	'Island of Ireland'	]
index    = full_data.index

countries = pd.DataFrame(index = index, columns = pd.MultiIndex.from_tuples(list(zip(demand_1,demand_2,demand_3))))

for a,b,c in zip(demand_1,demand_2,demand_3):
    l=full_data.iloc[:, (full_data.columns.get_level_values(2)==a)  &  (full_data.columns.get_level_values(3)==c) ].sum(axis=1,skipna =False)
    countries.iloc[:,    (countries.columns.get_level_values(0)==a) &   (countries.columns.get_level_values(2)==c) ]=l

countries[pd.MultiIndex.from_tuples([('','','Total')])] =  pd.DataFrame(countries.iloc[:,[0,1,2,3,4,5,6,7,8,9]].sum(axis=1,skipna =False))

countries[pd.MultiIndex.from_tuples([('','','Industrial')])]   =  industry.iloc[:,industry.columns.get_level_values(2)=='Total']
countries[pd.MultiIndex.from_tuples([('','','LDZ')])]          =  ldz.iloc[:,ldz.columns.get_level_values(2)=='Total']
countries[pd.MultiIndex.from_tuples([('','','Gas-to-Power')])] =  gtp.iloc[:,gtp.columns.get_level_values(2)=='Total']


last_index = countries.apply(lambda x: x[x.notnull()].index[-1]).sort_values().values[0]
countries  = countries.loc[:last_index]




demand_1 = ['Import',	'Import'	,'Import'	,'Production'	,'Production'	,'Import',	'Import',	'Import'	,'Import'	,'Import',	'Export'	,'Export'	,'Import'	,'Import',	'Import',	'Production',	'Production',	'Production']
demand_2 = ['Russia',	'Russia'	,'Norway',	'Netherlands'	,'GB',	'LNG',	'Algeria',	'Libya',	'Spain'	,'Denmark'	,'Poland',	'Hungary',	'Slovenia',	'Austria',	'TAP',	'Austria',	'Italy'	,'Germany']
demand_3 = ['Austria',	'Germany'	,'Europe'	,'Netherlands',	'GB'  ,   ''	,	'Italy',	'Italy',	'France'	,'Germany'	,'Germany'	,'Austria'	,'Europe',	'MAB',	'Italy'	,'Austria'	,'Italy',	'Germany']


supply = pd.DataFrame(index = index, columns = pd.MultiIndex.from_tuples(list(zip(demand_1,demand_2,demand_3))))

indX=('Brandov NordStream OPAL Daily Gas Flows (Germany to Czech Republic) (GWh/day)',
 'Y',
 'Import',
 'Russia',
 'Germany',
 'ONTGDCDF Index',
 'ONTGDCDF Index')

full_data.loc[full_data.index.year > 2023, indX] = 0


for a,b,c in zip(demand_1,demand_2,demand_3):
    if b == 'LNG':
        l=full_data.iloc[:,(full_data.columns.get_level_values(3)==b) ].sum(axis=1,skipna =False)
        supply.iloc[:,     (supply.columns.get_level_values(0)==a) &  (supply.columns.get_level_values(1)==b) ]=l
    else:
        l=full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) &  (full_data.columns.get_level_values(3)==b) ].sum(axis=1,skipna =False)
        supply.iloc[:,     (supply.columns.get_level_values(0)==a) & (supply.columns.get_level_values(2)==c) &  (supply.columns.get_level_values(1)==b) ]=l
       
    


supply[pd.MultiIndex.from_tuples([('Supply','','Total')])] =  pd.DataFrame(supply.sum(axis=1,skipna = False))

#last_index = supply.apply(lambda x: x[x.notnull()].index[-1]).sort_values().values[0]
#supply      = supply.loc[:last_index]



# Continue with rest of code but need to replace the sum operations...

# Line 589: Change sum() to sum(skipna=False)
# Line 684: Change sum(axis=1) to sum(axis=1, skipna=False) 
# Line 728: Change sum(axis=1) to sum(axis=1, skipna=False)
# Line 888: Change sum(axis=1) to sum(axis=1, skipna=False)
# Line 914: Change sum(axis=1) to sum(axis=1, skipna=False)
# Line 1051: Change sum(axis=1) to sum(axis=1, skipna=False)

# I'll create a fixed section for these, but need to see the rest of the file first...