# -*- coding: utf-8 -*-
"""
Created on Fri Feb 24 14:08:02 2023

@author: AD08394

SIMPLE FIX: Only change line 53 (add missing column) and line 68 (make unique)
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
data_path = 'G://Commodity Research//Gas//data//use4.xlsm'  # CHANGED: use3 -> use4
filename  = 'DNB Markets EUROPEAN GAS BALANCE.xlsx'
#data = 'DNB European Gas Market Balance Copy.xlsm'


#%% Load data, get demand components 
dataset   = pd.read_excel(data_path,sheet_name='TickerList',skiprows=8,usecols=range(1,9))

# FIX 1: Add missing column if it doesn't exist
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
     
    # FIX 2: Make the tuple unique by adding row number to last element
    temp_df.columns = pd.MultiIndex.from_tuples([(information.Description,information['Replace blanks with #N/A'],information.Category,information['Region from'],information['Region to'],information.Ticker,f"{information.Ticker}_{row}")])

    if row == 0:
        full_data = temp_df
    else:
        full_data = pd.merge(full_data,temp_df,left_index=True, right_index=True,how="outer")
   

# REST OF THE CODE REMAINS EXACTLY THE SAME...