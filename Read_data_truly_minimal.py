# -*- coding: utf-8 -*-
"""
Created on Fri Feb 24 14:08:02 2023

@author: AD08394
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
data_path = 'G://Commodity Research//Gas//data//use4.xlsm'  # CHANGED FROM use3 to use4
filename  = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'
#data = 'DNB European Gas Market Balance Copy.xlsm'


#%% Load data, get demand components 
dataset   = pd.read_excel(data_path,sheet_name='TickerList',skiprows=8,usecols=range(1,9))

# ONLY CHANGE: Add the missing column if it doesn't exist
if 'Replace blanks with #N/A' not in dataset.columns:
    dataset['Replace blanks with #N/A'] = 'N'

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
     
    
    temp_df.columns = pd.MultiIndex.from_tuples([(information.Description,information['Replace blanks with #N/A'],information.Category,information['Region from'],information['Region to'],information.Ticker,information.Ticker)])

    if row == 0:
        full_data = temp_df
    else:
        full_data = pd.merge(full_data,temp_df,left_index=True, right_index=True,how="outer")
   
    

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