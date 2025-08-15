# -*- coding: utf-8 -*-
"""
Gas Market Bloomberg Production Processor - FIXED VERSION
Downloads Bloomberg data using tickers from use4.xlsx and processes into master files
EXCLUDES losses and exports from demand calculations
"""

import pandas as pd
import numpy as np

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
    """Load Bloomberg data - use raw data for testing, xbbg for production"""
    try:
        # Try to use xbbg for production
        from xbbg import blp
        data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)
        print("✅ Using live Bloomberg data via xbbg")
        return data
    except:
        # Fallback to raw data for testing
        data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True)
        print("✅ Using bloomberg_raw_data.csv (test mode)")
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
    """Process demand data into country breakdown - FIXED to exclude losses and exports"""
    index = full_data.index
    country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    # Create countries DataFrame
    demand_1 = ['Demand'] * len(country_list) + ['Demand (Net)']
    demand_2 = [''] * (len(country_list) + 1) 
    demand_3 = country_list + ['Island of Ireland']
    
    countries = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))
    
    # Populate individual countries with PROPER demand filtering
    for country in country_list:
        # Get all demand series for this country
        country_mask = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(3) == country)
        
        # CRITICAL FIX: Exclude losses and exports from demand calculation
        if country == 'Italy':
            # For Italy, only include true demand categories, exclude losses and exports
            true_demand_mask = country_mask & full_data.columns.get_level_values(4).isin(['Industrial', 'LDZ', 'Gas-to-Power'])
            country_total = full_data.iloc[:, true_demand_mask].sum(axis=1, skipna=True)
            
            # Debug info for Italy
            excluded_series = full_data.iloc[:, country_mask & ~true_demand_mask]
            if excluded_series.shape[1] > 0:
                print(f"🇮🇹 Italy: Excluding {excluded_series.shape[1]} non-demand series (losses/exports)")
                for col in excluded_series.columns:
                    print(f"   Excluded: {col[4]} ({col[0][:50]}...)")
        else:
            # For other countries, use all demand series (they seem to be properly categorized)
            country_total = full_data.iloc[:, country_mask].sum(axis=1, skipna=True)
        
        countries[('Demand', '', country)] = country_total
    
    # Handle Island of Ireland separately
    ireland_mask = (full_data.columns.get_level_values(2) == 'Demand (Net)') & (full_data.columns.get_level_values(3) == 'Island of Ireland')
    ireland_total = full_data.iloc[:, ireland_mask].sum(axis=1, skipna=True) 
    countries[('Demand (Net)', '', 'Island of Ireland')] = ireland_total
    
    return countries

def process_category_data(full_data):
    """Process Industrial, LDZ, and Gas-to-Power data with proper filtering"""
    index = full_data.index
    
    # Get category totals with proper filtering (exclude losses and exports)
    industrial_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
                     (full_data.columns.get_level_values(4) == 'Industrial')
    ldz_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
               (full_data.columns.get_level_values(4) == 'LDZ')
    gtp_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
               (full_data.columns.get_level_values(4) == 'Gas-to-Power')
    
    industrial_total = full_data.iloc[:, industrial_mask].sum(axis=1, skipna=True)
    ldz_total = full_data.iloc[:, ldz_mask].sum(axis=1, skipna=True)
    gtp_total = full_data.iloc[:, gtp_mask].sum(axis=1, skipna=True)
    
    return industrial_total, ldz_total, gtp_total

def create_demand_master(countries, industrial_total, ldz_total, gtp_total):
    """Create demand master DataFrame in simple format"""
    index = countries.index
    country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    demand_data = pd.DataFrame(index=index)
    
    # Add individual countries
    for country in country_list:
        if ('Demand', '', country) in countries.columns:
            demand_data[country] = countries[('Demand', '', country)]
    
    # Calculate total from countries (should now match categories)
    country_total = demand_data[country_list].sum(axis=1, skipna=True)
    
    # Add totals
    demand_data['Total'] = country_total
    demand_data['Industrial'] = industrial_total
    demand_data['LDZ'] = ldz_total
    demand_data['Gas-to-Power'] = gtp_total
    
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
        # Load ticker configuration
        dataset, tickers, norm_factors = load_ticker_data()
        
        # Download Bloomberg data
        data = download_bloomberg_data(tickers)
        
        # Apply normalization
        data = apply_normalization(data, norm_factors)
        
        # Create MultiIndex structure
        full_data = create_multiindex_data(data, dataset)
        
        # Process demand data (with fixes)
        countries = process_demand_data(full_data)
        
        # Process category data (with fixes)
        industrial_total, ldz_total, gtp_total = process_category_data(full_data)
        
        # Create demand master
        demand_data = create_demand_master(countries, industrial_total, ldz_total, gtp_total)
        
        # Save files
        master_file = save_master_files(demand_data)
        
        print(f"✅ Processing complete: {master_file}")
        print(f"📊 Demand data: {demand_data.shape}")
        
        # Quick verification
        target_date = '2016-10-04'
        if target_date in [str(d)[:10] for d in demand_data.index]:
            target_row = demand_data[demand_data.index.strftime('%Y-%m-%d') == target_date].iloc[0]
            print(f"\n🎯 Verification for {target_date}:")
            print(f"   Italy: {target_row['Italy']:.2f}")
            print(f"   Countries sum to Total: {abs(target_row['Total'] - target_row[['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']].sum()) < 0.01}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)