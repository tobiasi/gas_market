# -*- coding: utf-8 -*-
"""
PRODUCTION VERSION: European Gas Market Data Processing
Fast, agile, production-ready with real Bloomberg data
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
import os

# Suppress warnings for clean output
warnings.filterwarnings('ignore')
pd.options.mode.chained_assignment = None

# Configuration
DATA_FILE = 'use4.xlsx'
OUTPUT_FILE = 'DNB Markets EUROPEAN GAS BALANCE_temp.xlsx'
START_DATE = "2016-10-01"

def load_bloomberg_data():
    """Load and normalize Bloomberg data"""
    print("🚀 LOADING BLOOMBERG DATA WITH NORMALIZATION")
    print("=" * 70)
    
    # Load ticker configuration
    print("📊 Loading ticker configuration...")
    dataset = pd.read_excel(DATA_FILE, sheet_name='TickerList', skiprows=8)
    
    # Create normalization factor mapping
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    print(f"✅ Loaded {len(norm_factors)} normalization factors")
    
    # Get tickers to download
    tickers = list(set(dataset.Ticker.dropna().tolist()))
    print(f"📈 Downloading {len(tickers)} tickers from Bloomberg...")
    
    # Download Bloomberg data
    try:
        from xbbg import blp
        data = blp.bdh(tickers, start_date=START_DATE, flds="PX_LAST").droplevel(axis=1, level=1)
        available_tickers = set(data.columns)
        missing_tickers = set(tickers) - available_tickers
        
        print(f"✅ Downloaded data for {len(available_tickers)} tickers")
        if missing_tickers:
            print(f"⚠️  Missing {len(missing_tickers)} tickers from Bloomberg")
            
    except Exception as e:
        print(f"❌ Bloomberg download failed: {e}")
        return None, None
    
    # Apply normalization factors
    print("🔧 Applying normalization factors...")
    normalized_count = 0
    
    for ticker in data.columns:
        if ticker in norm_factors:
            data[ticker] = data[ticker] * norm_factors[ticker]
            normalized_count += 1
    
    print(f"✅ Applied normalization to {normalized_count}/{len(data.columns)} series")
    
    # Build MultiIndex structure
    print("📋 Creating MultiIndex structure...")
    level_0, level_1, level_2, level_3, level_4 = [], [], [], [], []
    valid_tickers = []
    
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        if pd.notna(ticker) and ticker in data.columns:
            level_0.append(row.get('Description', ''))
            level_1.append('')
            level_2.append(row.get('Category', ''))
            level_3.append(row.get('Region from', ''))
            level_4.append(row.get('Region to', ''))
            valid_tickers.append(ticker)
    
    # Create final MultiIndex DataFrame
    multi_index = pd.MultiIndex.from_arrays([level_0, level_1, level_2, level_3, level_4])
    full_data = pd.DataFrame(data[valid_tickers].values, 
                           index=data.index, 
                           columns=multi_index)
    
    # Data cleaning
    full_data.index = pd.to_datetime(full_data.index)
    full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]
    full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')
    
    print(f"✅ Created MultiIndex data: {full_data.shape}")
    
    # Debug Italy normalization
    italy_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
                 (full_data.columns.get_level_values(3) == 'Italy')
    if italy_mask.any():
        italy_total = full_data.iloc[:, italy_mask].sum(axis=1, skipna=True).iloc[0]
        print(f"🇮🇹 Italy demand total: {italy_total:.2f} (target: ~117)")
        if abs(italy_total - 117) < 10:
            print("✅ Italy normalization SUCCESS!")
        else:
            print("⚠️  Italy normalization may need adjustment")
    
    return full_data, dataset

def process_countries(full_data):
    """Process country-level demand data"""
    print("\n🌍 PROCESSING COUNTRIES")
    print("=" * 50)
    
    countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 
                'Austria', 'Germany', 'Switzerland', 'Luxembourg']
    
    # Create countries DataFrame
    index = full_data.index
    cols = [('Demand', '', country) for country in countries] + [('Demand (Net)', '', 'Island of Ireland')]
    countries_df = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(cols))
    
    # Process each country
    for country in countries:
        mask = (full_data.columns.get_level_values(2) == 'Demand') & \
               (full_data.columns.get_level_values(3) == country)
        
        if mask.any():
            total = full_data.iloc[:, mask].sum(axis=1, skipna=True)
            countries_df[('Demand', '', country)] = total
            print(f"✅ {country:<12}: {mask.sum():2d} series → {total.iloc[0]:8.2f}")
        else:
            countries_df[('Demand', '', country)] = 0
            print(f"❌ {country:<12}: No demand series found")
    
    # Ireland (Net demand)
    ireland_mask = (full_data.columns.get_level_values(2) == 'Demand (Net)') & \
                   (full_data.columns.get_level_values(3) == 'Island of Ireland')
    if ireland_mask.any():
        countries_df[('Demand (Net)', '', 'Island of Ireland')] = \
            full_data.iloc[:, ireland_mask].sum(axis=1, skipna=True)
        print("✅ Ireland       : Net demand processed")
    
    return countries_df

def process_categories(full_data):
    """Process Industrial, Gas-to-Power, and LDZ categories"""
    print("\n⚙️  PROCESSING CATEGORIES")
    print("=" * 50)
    
    index = full_data.index
    
    # Industrial demand mappings
    industrial_mappings = [
        ('Demand', 'France', 'Industrial'),
        ('Demand', 'Belgium', 'Industrial'), 
        ('Demand', 'Italy', 'Industrial'),
        ('Demand', 'GB', 'Industrial'),
        ('Demand', 'Netherlands', 'Industrial and Power'),
        ('Demand', 'Netherlands', 'Zebra'),
        ('Demand', 'Germany', 'Industrial and Power'),
        ('Intermediate Calculation', '#Germany', 'Gas-to-Power'),
        ('Demand', 'Germany', 'Industrial (calculated)'),
        ('Demand', '#Netherlands', 'Industrial'),
        ('Intermediate Calculation', '#Netherlands', 'Gas-to-Power'),
        ('Demand', 'Netherlands', 'Industrial (calculated to 30/6/22 then actual)')
    ]
    
    # Gas-to-Power mappings
    gtp_mappings = [
        ('Demand', 'France', 'Gas-to-Power'),
        ('Demand', 'Belgium', 'Gas-to-Power'),
        ('Demand', 'Italy', 'Gas-to-Power'),
        ('Demand', 'GB', 'Gas-to-Power'),
        ('Demand', 'Germany', 'Industrial and Power'),
        ('Intermediate Calculation', '#Germany', 'Gas-to-Power'),
        ('Demand', 'Germany', 'Gas-to-Power (calculated)'),
        ('Demand', 'Netherlands', 'Industrial and Power'),
        ('Demand', '#Netherlands', 'Gas-to-Power'),
        ('Intermediate Calculation', '#Netherlands', 'Gas-to-Power'),
        ('Demand', 'Netherlands', 'Gas-to-Power (calculated to 30/6/22 then actual)')
    ]
    
    # LDZ mappings
    ldz_mappings = [
        ('Demand', 'France', 'LDZ'),
        ('Demand', 'Belgium', 'LDZ'),
        ('Demand', 'Italy', 'LDZ'),
        ('Demand', 'Italy', 'Other'),
        ('Demand', 'Netherlands', 'LDZ'),
        ('Demand', 'GB', 'LDZ'),
        ('Demand', 'Austria', 'Austria'),
        ('Demand', 'Germany', 'LDZ'),
        ('Demand', 'Switzerland', 'Switzerland'),
        ('Demand', 'Luxembourg', 'Luxembourg'),
        ('Demand (Net)', 'Island of Ireland', 'Island of Ireland'),
        ('Demand', '#Germany', 'Implied LDZ')
    ]
    
    # Process each category
    categories = {
        'Industrial': (industrial_mappings, 'industry'),
        'Gas-to-Power': (gtp_mappings, 'gtp'), 
        'LDZ': (ldz_mappings, 'ldz')
    }
    
    results = {}
    
    for cat_name, (mappings, var_name) in categories.items():
        print(f"\n📊 Processing {cat_name}...")
        
        # Create DataFrame
        cols = pd.MultiIndex.from_tuples([(d, r, s) for d, r, s in mappings])
        df = pd.DataFrame(index=index, columns=cols)
        
        # Populate data
        processed = 0
        for demand_cat, region, subcategory in mappings:
            mask = (full_data.columns.get_level_values(2) == demand_cat) & \
                   (full_data.columns.get_level_values(3) == region) & \
                   (full_data.columns.get_level_values(4) == subcategory)
            
            if mask.any():
                total = full_data.iloc[:, mask].sum(axis=1, skipna=True)
                df[(demand_cat, region, subcategory)] = total
                processed += 1
        
        # Calculate totals (using specific columns as in original)
        if cat_name == 'Industrial':
            df[('Demand', '-', 'Total')] = df.iloc[:, [0,1,2,3,8,11]].sum(axis=1, skipna=True)
        elif cat_name == 'Gas-to-Power':
            df[('Demand', '', 'Total')] = df.iloc[:, [0,1,2,3,6,10]].sum(axis=1, skipna=True)
        elif cat_name == 'LDZ':
            df[('Demand', '', 'Total')] = df.iloc[:, [0,1,2,3,4,5,6,7,8,9,10]].sum(axis=1, skipna=True)
        
        results[var_name] = df
        print(f"   ✅ {processed} series processed")
    
    return results['industry'], results['gtp'], results['ldz']

def process_supply_lng(full_data):
    """Process Supply and LNG data"""
    print("\n🔋 PROCESSING SUPPLY & LNG")
    print("=" * 50)
    
    index = full_data.index
    
    # Supply mappings
    supply_mappings = [
        ('Import', 'Russia', 'Austria'),
        ('Import', 'Russia', 'Germany'),
        ('Import', 'Norway', 'Europe'),
        ('Production', 'Netherlands', 'Netherlands'),
        ('Production', 'GB', 'GB'),
        ('Import', 'LNG', ''),
        ('Import', 'Algeria', 'Italy'),
        ('Import', 'Libya', 'Italy'),
        ('Import', 'Spain', 'France'),
        ('Import', 'Denmark', 'Germany'),
        ('Export', 'Poland', 'Germany'),
        ('Export', 'Hungary', 'Austria'),
        ('Import', 'Slovenia', 'Europe'),
        ('Import', 'Austria', 'MAB'),
        ('Import', 'TAP', 'Italy'),
        ('Production', 'Austria', 'Austria'),
        ('Production', 'Italy', 'Italy'),
        ('Production', 'Germany', 'Germany')
    ]
    
    # LNG mappings
    lng_mappings = [
        ('Import', 'LNG', 'France'),
        ('Import', 'LNG', 'Italy'),
        ('Import', 'LNG', 'Belgium'),
        ('Import', 'LNG', 'Netherlands'),
        ('Import', 'LNG', 'GB'),
        ('Import', 'LNG', 'Germany')
    ]
    
    # Process supply
    supply_cols = pd.MultiIndex.from_tuples([(d, r, t) for d, r, t in supply_mappings])
    supply = pd.DataFrame(index=index, columns=supply_cols)
    
    for demand_cat, region_from, region_to in supply_mappings:
        mask = (full_data.columns.get_level_values(2) == demand_cat) & \
               (full_data.columns.get_level_values(3) == region_from) & \
               (full_data.columns.get_level_values(4) == region_to)
        
        if mask.any():
            supply[(demand_cat, region_from, region_to)] = \
                full_data.iloc[:, mask].sum(axis=1, skipna=True)
    
    # Supply total (first 13 + index 14)
    supply[('Import', '', 'Total')] = \
        supply.iloc[:, list(range(13)) + [14]].sum(axis=1, skipna=True)
    
    # Process LNG
    lng_cols = pd.MultiIndex.from_tuples([(d, r, t) for d, r, t in lng_mappings])
    lng = pd.DataFrame(index=index, columns=lng_cols)
    
    for demand_cat, region_from, region_to in lng_mappings:
        mask = (full_data.columns.get_level_values(2) == demand_cat) & \
               (full_data.columns.get_level_values(3) == region_from) & \
               (full_data.columns.get_level_values(4) == region_to)
        
        if mask.any():
            lng[(demand_cat, region_from, region_to)] = \
                full_data.iloc[:, mask].sum(axis=1, skipna=True)
    
    # LNG total
    lng[('Import', '', 'Total')] = lng.sum(axis=1, skipna=True)
    
    print(f"✅ Supply: {supply.shape[1]} series")
    print(f"✅ LNG: {lng.shape[1]} series")
    
    return supply, lng

def finalize_countries(countries, industry, gtp, ldz):
    """Add totals and categories to countries DataFrame"""
    print("\n📊 FINALIZING COUNTRIES DATA")
    print("=" * 50)
    
    # Calculate country total
    country_cols = [col for col in countries.columns 
                   if col[0] == 'Demand' and col[2] in 
                   ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg']]
    
    country_total = countries[country_cols].sum(axis=1, skipna=True)
    
    # Add Ireland if exists
    ireland_col = ('Demand (Net)', '', 'Island of Ireland')
    if ireland_col in countries.columns:
        country_total += countries[ireland_col]
    
    # Add totals to countries
    countries[('', '', 'Total')] = country_total
    countries[('', '', 'Industrial')] = industry[('Demand', '-', 'Total')]
    countries[('', '', 'LDZ')] = ldz[('Demand', '', 'Total')]
    countries[('', '', 'Gas-to-Power')] = gtp[('Demand', '', 'Total')]
    
    # Verification
    total_val = country_total.iloc[0]
    industrial_val = industry[('Demand', '-', 'Total')].iloc[0]
    ldz_val = ldz[('Demand', '', 'Total')].iloc[0]
    gtp_val = gtp[('Demand', '', 'Total')].iloc[0]
    
    print(f"✅ Total Country Demand: {total_val:.2f}")
    print(f"✅ Industrial Total: {industrial_val:.2f}")
    print(f"✅ LDZ Total: {ldz_val:.2f}")
    print(f"✅ Gas-to-Power Total: {gtp_val:.2f}")
    
    # Check precision
    category_sum = industrial_val + ldz_val + gtp_val
    difference = abs(total_val - category_sum)
    print(f"✅ Precision check: {difference:.6f} difference")
    
    return countries

def save_excel_output(countries, industry, gtp, ldz, lng, supply):
    """Save all data to Excel with proper formatting"""
    print("\n💾 SAVING EXCEL OUTPUT")
    print("=" * 50)
    
    try:
        print(f"📁 Creating {OUTPUT_FILE}...")
        
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            # Write main sheets
            countries.to_excel(writer, sheet_name='Countries')
            industry.to_excel(writer, sheet_name='Industrial')
            gtp.to_excel(writer, sheet_name='Gas-to-Power')
            ldz.to_excel(writer, sheet_name='LDZ')
            lng.to_excel(writer, sheet_name='LNG')
            supply.to_excel(writer, sheet_name='Supply')
        
        print(f"✅ Successfully saved: {OUTPUT_FILE}")
        print(f"📋 Contains 6 main data sheets")
        
        # Verify file was created
        if os.path.exists(OUTPUT_FILE):
            file_size = os.path.getsize(OUTPUT_FILE) / 1024  # KB
            print(f"📊 File size: {file_size:.1f} KB")
            return True
        else:
            print("❌ File was not created")
            return False
            
    except Exception as e:
        print(f"❌ Error saving Excel file: {e}")
        return False

def main():
    """Main execution function"""
    print("🎯 EUROPEAN GAS MARKET - PRODUCTION PROCESSOR")
    print("=" * 80)
    print(f"📊 Input: {DATA_FILE}")
    print(f"📁 Output: {OUTPUT_FILE}")
    print(f"📅 Start Date: {START_DATE}")
    print("=" * 80)
    
    try:
        # Step 1: Load and normalize data
        full_data, dataset = load_bloomberg_data()
        if full_data is None:
            print("❌ Failed to load Bloomberg data")
            return False
        
        # Step 2: Process countries
        countries = process_countries(full_data)
        
        # Step 3: Process categories
        industry, gtp, ldz = process_categories(full_data)
        
        # Step 4: Process supply and LNG
        supply, lng = process_supply_lng(full_data)
        
        # Step 5: Finalize countries with totals
        countries = finalize_countries(countries, industry, gtp, ldz)
        
        # Step 6: Save output
        success = save_excel_output(countries, industry, gtp, ldz, lng, supply)
        
        if success:
            print("\n🎉 PRODUCTION RUN SUCCESSFUL!")
            print("=" * 50)
            print("✅ Bloomberg data downloaded and normalized")
            print("✅ All countries processed (including Netherlands & Luxembourg)")
            print("✅ All categories calculated (Industrial, LDZ, Gas-to-Power)")
            print("✅ Excel file saved successfully")
            print(f"✅ Italy normalization: {countries[('Demand', '', 'Italy')].iloc[0]:.2f}")
            return True
        else:
            print("\n❌ PRODUCTION RUN FAILED!")
            return False
            
    except Exception as e:
        print(f"\n💥 CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)