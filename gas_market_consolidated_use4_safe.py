# -*- coding: utf-8 -*-
"""
SAFE VERSION - With extensive error checking and debugging
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

# [Include all the imports and setup from the main script...]

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

# Add this to your main script to test first:
print("🧪 Run safe_test_aggregation(full_data, index) first to validate the logic!")