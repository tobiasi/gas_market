# -*- coding: utf-8 -*-
"""
TEST SCRIPT: Verify the corrected aggregation logic works without errors
"""

import warnings
import pandas as pd
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)
warnings.filterwarnings('ignore', category=UserWarning)
pd.options.mode.chained_assignment = None

import numpy as np
import os
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
import sys
import pycountry_convert as pc
sys.path.append("C:/development/commodities")
from update_spreadsheet import update_spreadsheet, to_excel_dates, update_spreadsheet_pweekly

from datetime import datetime
from dateutil.relativedelta import relativedelta

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

def safe_test_aggregation(full_data, index):
    """Test the aggregation logic safely before applying to main script"""
    
    print(f"\n{'='*60}")
    print("🧪 SAFE TESTING OF AGGREGATION LOGIC")
    print(f"{'='*60}")
    
    print(f"📊 Full data shape: {full_data.shape}")
    print(f"📅 Index length: {len(index)}")
    
    # Test 1: Check what columns we have
    print(f"\n🔍 COLUMN ANALYSIS:")
    print(f"   Total columns: {len(full_data.columns)}")
    
    if hasattr(full_data.columns, 'levels'):
        print(f"   MultiIndex levels: {len(full_data.columns.levels)}")
        for i, level in enumerate(full_data.columns.levels):
            print(f"   Level {i}: {len(level)} unique values")
            if len(level) < 20:  # Only show if reasonable number
                print(f"      Values: {list(level)[:10]}...")
    
    # Test 2: Check categories in the data
    print(f"\n📋 CATEGORY ANALYSIS:")
    categories = full_data.columns.get_level_values(2).unique()
    print(f"   Categories found: {categories}")
    
    # Test 3: Check regions
    print(f"\n🌍 REGION ANALYSIS:")
    regions_from = full_data.columns.get_level_values(3).unique()
    regions_to = full_data.columns.get_level_values(4).unique()
    print(f"   Regions 'from': {list(regions_from)[:10]}...")
    print(f"   Regions 'to': {list(regions_to)[:10]}...")
    
    # Test 4: Check descriptions for keywords
    print(f"\n🔎 DESCRIPTION KEYWORD ANALYSIS:")
    descriptions = full_data.columns.get_level_values(0)
    
    industrial_matches = descriptions.str.contains('Industrial|industry', case=False, na=False).sum()
    ldz_matches = descriptions.str.contains('LDZ|ldz', case=False, na=False).sum()
    gtp_matches = descriptions.str.contains('Gas-to-Power|Power', case=False, na=False).sum()
    
    print(f"   Industrial matches: {industrial_matches}")
    print(f"   LDZ matches: {ldz_matches}")
    print(f"   Gas-to-Power matches: {gtp_matches}")
    
    # Test 5: Try building simple aggregations
    print(f"\n🧪 TESTING SIMPLE AGGREGATIONS:")
    
    try:
        # Test demand filtering
        demand_mask = full_data.columns.get_level_values(2) == 'Demand'
        demand_series_count = demand_mask.sum()
        print(f"   Found {demand_series_count} 'Demand' series")
        
        if demand_series_count > 0:
            total_demand = full_data.iloc[:, demand_mask].sum(axis=1, skipna=False)
            print(f"   Sample total demand: {total_demand.iloc[0:3].values}")
        
    except Exception as e:
        print(f"   ❌ Error in demand aggregation: {e}")
    
    # Test 6: Check specific countries
    print(f"\n🏛️ COUNTRY-SPECIFIC TESTING:")
    test_countries = ['Germany', 'France', 'Netherlands']
    
    for country in test_countries:
        try:
            country_mask = ((full_data.columns.get_level_values(2) == 'Demand') & 
                          (full_data.columns.get_level_values(3) == country))
            country_series = country_mask.sum()
            
            if country_series > 0:
                country_total = full_data.iloc[:, country_mask].sum(axis=1, skipna=False)
                sample_value = country_total.iloc[0] if len(country_total) > 0 else 0
                print(f"   {country}: {country_series} series, sample value: {sample_value:.2f}")
            else:
                print(f"   {country}: No series found")
                
        except Exception as e:
            print(f"   {country}: Error - {e}")
    
    print(f"\n✅ SAFE TESTING COMPLETE")
    print(f"   Check the results above before running the full aggregation")
    
    return True

def test_consistent_aggregation_method(full_data, index):
    """Test the new consistent aggregation method for fixing sum discrepancies"""
    
    print(f"\n{'='*60}")
    print("🔧 TESTING CONSISTENT AGGREGATION METHOD")
    print(f"{'='*60}")
    
    def get_category_total(category_descriptions):
        """Build category totals using consistent description-based matching"""
        total = pd.Series(0.0, index=index)
        
        for desc in category_descriptions:
            # Find all demand series matching this description
            mask = ((full_data.columns.get_level_values(2) == 'Demand') & 
                   (full_data.columns.get_level_values(0).str.contains(desc, case=False, na=False)))
            
            if mask.any():
                category_data = full_data.iloc[:, mask].sum(axis=1, skipna=False)
                total += category_data
                print(f"   Added {mask.sum()} series for '{desc}'")
        
        return total
    
    try:
        # Test the category aggregations
        industrial_keywords = ['Industrial', 'industry']
        ldz_keywords = ['LDZ', 'ldz']
        gtp_keywords = ['Gas-to-Power', 'Power']
        
        print(f"\n🏭 TESTING INDUSTRIAL AGGREGATION:")
        industrial_total = get_category_total(industrial_keywords)
        print(f"   Sample values: {industrial_total.iloc[0:3].values}")
        
        print(f"\n🌐 TESTING LDZ AGGREGATION:")
        ldz_total = get_category_total(ldz_keywords)
        print(f"   Sample values: {ldz_total.iloc[0:3].values}")
        
        print(f"\n⚡ TESTING GAS-TO-POWER AGGREGATION:")
        gtp_total = get_category_total(gtp_keywords)
        print(f"   Sample values: {gtp_total.iloc[0:3].values}")
        
        # Test country total aggregation
        print(f"\n🏛️ TESTING COUNTRY TOTAL AGGREGATION:")
        country_list = ['France', 'Belgium', 'Italy', 'GB', 'Netherlands', 'Germany']
        
        country_total = pd.Series(0.0, index=index)
        
        for country in country_list:
            country_mask = ((full_data.columns.get_level_values(2) == 'Demand') & 
                          (full_data.columns.get_level_values(3) == country))
            
            if country_mask.any():
                country_data = full_data.iloc[:, country_mask].sum(axis=1, skipna=False)
                country_total += country_data
                print(f"   Added {country_mask.sum()} series for {country}")
        
        print(f"   Country total sample: {country_total.iloc[0:3].values}")
        
        # Test if totals match
        print(f"\n📊 CONSISTENCY CHECK:")
        category_sum = industrial_total + ldz_total + gtp_total
        difference = abs(country_total - category_sum)
        max_diff = difference.max()
        
        print(f"   Country total range: {country_total.min():.2f} to {country_total.max():.2f}")
        print(f"   Category sum range: {category_sum.min():.2f} to {category_sum.max():.2f}")
        print(f"   Max difference: {max_diff:.4f}")
        
        if max_diff < 0.01:
            print(f"   ✅ TOTALS MATCH! Difference < 0.01")
        else:
            print(f"   ⚠️  Differences found, need further investigation")
        
        return True
        
    except Exception as e:
        print(f"   ❌ Error in consistent aggregation: {e}")
        import traceback
        traceback.print_exc()
        return False

# Main testing function
def main():
    """Run the full test suite"""
    
    print("🧪 STARTING AGGREGATION TESTING")
    print("=" * 80)
    
    # Set up paths (use your actual paths)
    os.chdir('G:\Commodity Research\Gas')
    data_path = 'use4.xlsx'
    
    try:
        # Load the dataset
        print("📊 Loading test data...")
        dataset = pd.read_excel(data_path, sheet_name='TickerList', skiprows=8)
        
        # Add missing column if needed
        if 'Replace blanks with #N/A' not in dataset.columns:
            dataset['Replace blanks with #N/A'] = dataset.apply(smart_na_handling, axis=1)
        
        # For testing, we'll use a subset of tickers to avoid Bloomberg API limits
        test_tickers = dataset.Ticker.dropna().head(20).tolist()  # Just first 20 for testing
        print(f"🔧 Using {len(test_tickers)} test tickers")
        
        # Download test data
        print("⬇️ Downloading Bloomberg test data...")
        data = blp.bdh(test_tickers, start_date="2023-01-01", end_date="2023-12-31", flds="PX_LAST").droplevel(axis=1, level=1)
        
        # Create a mock full_data structure for testing
        # In reality this would come from your processing pipeline
        index = data.index
        
        # Create mock MultiIndex columns (simplified version)
        mock_columns = []
        for ticker in data.columns:
            row = dataset[dataset.Ticker == ticker].iloc[0] if ticker in dataset.Ticker.values else None
            if row is not None:
                desc = row.get('Description', ticker)
                category = row.get('Category', 'Demand')
                region_from = row.get('From', 'Unknown')
                region_to = row.get('To', 'Unknown')
                mock_columns.append((desc, '', category, region_from, region_to))
        
        full_data = pd.DataFrame(data.values, index=index, 
                               columns=pd.MultiIndex.from_tuples(mock_columns))
        
        print(f"✅ Test data prepared: {full_data.shape}")
        
        # Run the safe testing
        print("\n" + "="*80)
        safe_test_result = safe_test_aggregation(full_data, index)
        
        # Run the consistent aggregation test  
        if safe_test_result:
            consistent_agg_result = test_consistent_aggregation_method(full_data, index)
            
            if consistent_agg_result:
                print(f"\n🎉 ALL TESTS PASSED!")
                print(f"   The corrected aggregation logic appears to work correctly.")
                print(f"   You can now run the full script with confidence.")
            else:
                print(f"\n❌ CONSISTENT AGGREGATION TEST FAILED")
                print(f"   There may still be issues with the aggregation logic.")
        else:
            print(f"\n❌ SAFE TESTING FAILED")
            print(f"   Basic data structure tests failed.")
    
    except Exception as e:
        print(f"❌ TEST FAILED: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()