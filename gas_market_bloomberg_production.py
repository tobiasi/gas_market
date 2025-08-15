# -*- coding: utf-8 -*-
"""
Gas Market Bloomberg Production Processor
Downloads Bloomberg data using tickers from use4.xlsx and processes into master files
"""

import pandas as pd
import numpy as np
from xbbg import blp
import json

def load_ticker_data():
    """Load ticker list and normalization factors from use4.xlsx"""
    data_path = 'use4.xlsx'
    
    # Read ticker list from TickerList sheet
    dataset = pd.read_excel(data_path, sheet_name='TickerList', skiprows=8)
    
    # Extract normalization factors
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    # Get unique tickers
    tickers = list(set(dataset.Ticker.dropna().to_list()))
    
    return dataset, tickers, norm_factors

def download_bloomberg_data(tickers):
    """Download Bloomberg data for tickers"""
    data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)
    return data

def apply_normalization(data, norm_factors):
    """Apply normalization factors to Bloomberg data"""
    for ticker in data.columns:
        if ticker in norm_factors:
            norm_factor = norm_factors[ticker]
            data[ticker] = data[ticker] * norm_factor
    return data

def create_multiindex_data(data, dataset):
    """Create MultiIndex DataFrame from Bloomberg data"""
    successful_tickers = []
    level_0 = []  # Description
    level_1 = []  # Empty
    level_2 = []  # Category  
    level_3 = []  # Region from
    level_4 = []  # Region to
    
    # Process each ticker in dataset
    for index, row in dataset.iterrows():
        ticker = row.get('Ticker')
        
        if pd.isna(ticker) or ticker not in data.columns:
            continue
        
        successful_tickers.append(ticker)
        level_0.append(row.get('Description', ''))
        level_1.append('')  # Empty level
        level_2.append(row.get('Category', ''))
        level_3.append(row.get('Region from', ''))
        level_4.append(row.get('Region to', ''))
    
    # Create MultiIndex DataFrame
    if successful_tickers:
        ticker_data = data[successful_tickers]
        multi_index = pd.MultiIndex.from_arrays([level_0, level_1, level_2, level_3, level_4])
        full_data = pd.DataFrame(ticker_data.values, index=ticker_data.index, columns=multi_index)
        
        # Apply data cleaning
        full_data.index = pd.to_datetime(full_data.index)
        full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]
        full_data = full_data.copy()
        full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')
        
        return full_data
    else:
        raise ValueError("No successful tickers to process")

def process_demand_data(full_data):
    """Process demand data into country breakdown"""
    index = full_data.index
    country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    # Create countries DataFrame
    demand_1 = ['Demand'] * len(country_list) + ['Demand (Net)']
    demand_2 = [''] * (len(country_list) + 1) 
    demand_3 = country_list + ['Island of Ireland']
    
    countries = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))
    
    # Populate individual countries
    for country in country_list:
        country_mask = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(3) == country)
        country_total = full_data.iloc[:, country_mask].sum(axis=1, skipna=True)
        countries[('Demand', '', country)] = country_total
    
    # Handle Island of Ireland separately
    ireland_mask = (full_data.columns.get_level_values(2) == 'Demand (Net)') & (full_data.columns.get_level_values(3) == 'Island of Ireland')
    ireland_total = full_data.iloc[:, ireland_mask].sum(axis=1, skipna=True) 
    countries[('Demand (Net)', '', 'Island of Ireland')] = ireland_total
    
    return countries

def process_category_data(full_data):
    """Process Industrial, LDZ, and Gas-to-Power data"""
    index = full_data.index
    
    # Process Industrial data
    demand_1 = ['Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand','Demand', 'Intermediate Calculation','Demand','Demand', 'Intermediate Calculation', 'Demand']
    demand_2 = ['France','Belgium','Italy','GB','Netherlands','Netherlands','Germany','#Germany','Germany','#Netherlands', '#Netherlands', 'Netherlands']
    demand_3 = ['Industrial','Industrial','Industrial','Industrial','Industrial and Power','Zebra','Industrial and Power','Gas-to-Power','Industrial (calculated)','Industrial', 'Gas-to-Power', 'Industrial (calculated to 30/6/22 then actual)']
    
    industry = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))
    
    for a, b, c in zip(demand_1, demand_2, demand_3):
        l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=True)
        industry.iloc[:, (industry.columns.get_level_values(0)==a) & (industry.columns.get_level_values(2)==c) & (industry.columns.get_level_values(1)==b)] = l
    
    # Apply Industrial calculations
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
    industry[pd.MultiIndex.from_tuples([('Demand','-','Total')])] = pd.DataFrame(industry.iloc[:,[0,1,2,3,8,11]].sum(axis=1, skipna=True))
    
    # Process Gas-to-Power data
    demand_1 = ['Demand', 'Demand', 'Demand' ,'Demand' ,'Demand', 'Intermediate Calculation' ,'Demand' , 'Demand', 'Demand', 'Intermediate Calculation', 'Demand']
    demand_2 = ['France','Belgium','Italy','GB','Germany', '#Germany','Germany','Netherlands','#Netherlands','#Netherlands','Netherlands']
    demand_3 = ['Gas-to-Power', 'Gas-to-Power' ,'Gas-to-Power' ,'Gas-to-Power', 'Industrial and Power', 'Gas-to-Power', 'Gas-to-Power (calculated)' , 'Industrial and Power', 'Gas-to-Power' ,'Gas-to-Power' ,'Gas-to-Power (calculated to 30/6/22 then actual)']
    
    gtp = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))
    
    for a, b, c in zip(demand_1, demand_2, demand_3):
        l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=True)
        gtp.iloc[:, (gtp.columns.get_level_values(0)==a) & (gtp.columns.get_level_values(2)==c) & (gtp.columns.get_level_values(1)==b)] = l
    
    # Apply GTP calculations
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
    gtp[pd.MultiIndex.from_tuples([('Demand','','Total')])] = pd.DataFrame(gtp.iloc[:,[0,1,2,3,6,10]].sum(axis=1, skipna=True))
    
    # Process LDZ data
    demand_1 = ['Demand' ,'Demand' ,'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand', 'Demand (Net)', 'Demand' ,'Delta', 'Delta']
    demand_2 = ['France', 'Belgium' ,'Italy' ,'Italy' ,'Netherlands', 'GB', 'Austria' ,'Germany' ,'Switzerland', 'Luxembourg', 'Island of Ireland' , '#Germany' ,'Germany', 'Germany']
    demand_3 = ['LDZ', 'LDZ', 'LDZ', 'Other', 'LDZ', 'LDZ', 'Austria', 'LDZ', 'Switzerland', 'Luxembourg', 'Island of Ireland', 'Implied LDZ','', 'Cumulative']
    
    ldz = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))
    
    for a, b, c in zip(demand_1, demand_2, demand_3):
        l = full_data.iloc[:, (full_data.columns.get_level_values(2)==a) & (full_data.columns.get_level_values(4)==c) & (full_data.columns.get_level_values(3)==b)].sum(axis=1, skipna=True)
        ldz.iloc[:, (ldz.columns.get_level_values(0)==a) & (ldz.columns.get_level_values(2)==c) & (ldz.columns.get_level_values(1)==b)] = l
    
    ldz[pd.MultiIndex.from_tuples([('Demand','','Total')])] = pd.DataFrame(ldz.iloc[:,[0,1,2,3,4,5,6,7,8,9,10]].sum(axis=1, skipna=True))
    
    return industry, gtp, ldz

def create_demand_master(countries, industry, gtp, ldz):
    """Create demand master DataFrame in simple format"""
    index = countries.index
    
    # Extract country data
    country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    demand_data = pd.DataFrame(index=index)
    
    # Add individual countries
    for country in country_list:
        if ('Demand', '', country) in countries.columns:
            demand_data[country] = countries[('Demand', '', country)]
    
    # Add totals
    demand_data['Total'] = industry[('Demand','-','Total')] if ('Demand','-','Total') in industry.columns else 0
    demand_data['Industrial'] = industry[('Demand','-','Total')] if ('Demand','-','Total') in industry.columns else 0
    demand_data['LDZ'] = ldz[('Demand','','Total')] if ('Demand','','Total') in ldz.columns else 0
    demand_data['Gas-to-Power'] = gtp[('Demand','','Total')] if ('Demand','','Total') in gtp.columns else 0
    
    return demand_data

def save_master_files(demand_data):
    """Save master files"""
    # Reset index to make Date a column
    final_demand_df = demand_data.copy()
    final_demand_df.reset_index(inplace=True)
    final_demand_df.rename(columns={'index': 'Date'}, inplace=True)
    
    # Save Excel file
    master_file = 'European_Gas_Market_Master.xlsx'
    with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
        final_demand_df.to_excel(writer, sheet_name='Demand', index=False)
    
    # Save CSV file
    final_demand_df.to_csv('European_Gas_Demand_Master.csv', index=False)
    
    return master_file

def main():
    """Main execution function"""
    try:
        # Load ticker data
        dataset, tickers, norm_factors = load_ticker_data()
        
        # Download Bloomberg data
        data = download_bloomberg_data(tickers)
        
        # Apply normalization
        data = apply_normalization(data, norm_factors)
        
        # Create MultiIndex data
        full_data = create_multiindex_data(data, dataset)
        
        # Process demand data
        countries = process_demand_data(full_data)
        
        # Process category data
        industry, gtp, ldz = process_category_data(full_data)
        
        # Create demand master
        demand_data = create_demand_master(countries, industry, gtp, ldz)
        
        # Save files
        master_file = save_master_files(demand_data)
        
        print(f"✅ Processing complete: {master_file}")
        print(f"📊 Demand data: {demand_data.shape}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)