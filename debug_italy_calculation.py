# -*- coding: utf-8 -*-
"""
DEBUG: Identify why Italy calculation is wrong
"""

import pandas as pd
import numpy as np

def debug_italy_calculation():
    """Debug Italy calculation step by step"""
    
    print(f"\n{'='*80}")
    print("🔍 DEBUG ITALY CALCULATION")
    print(f"{'='*80}")
    
    # Load ticker configuration
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    
    # Find Italy DEMAND series
    italy_demand = dataset[
        (dataset['Description'].str.contains('Italy', case=False, na=False)) &
        (dataset['Category'] == 'Demand')
    ]
    
    print(f"🇮🇹 ITALY DEMAND SERIES FOUND:")
    print(f"   Count: {len(italy_demand)}")
    
    for _, row in italy_demand.iterrows():
        ticker = row['Ticker']
        desc = row['Description']
        norm_factor = row.get('Normalization factor', 1.0)
        region_from = row.get('Region from', '')
        region_to = row.get('Region to', '')
        
        print(f"\n📊 {ticker}:")
        print(f"   Description: {desc}")
        print(f"   Normalization: {norm_factor}")
        print(f"   Region from: {region_from}")
        print(f"   Region to: {region_to}")
        
        # Simulate what would happen with dummy data
        dummy_value = 1000  # Our dummy data is random 0-1000
        normalized_value = dummy_value * norm_factor
        print(f"   Dummy value: {dummy_value}")
        print(f"   After normalization: {normalized_value:.2f}")
    
    # Check if we should be looking at Region from = Italy
    italy_region_from = dataset[
        (dataset['Region from'] == 'Italy') &
        (dataset['Category'] == 'Demand')
    ]
    
    print(f"\n🎯 SERIES WHERE REGION FROM = 'ITALY':")
    print(f"   Count: {len(italy_region_from)}")
    
    if len(italy_region_from) > 0:
        print("   Series found:")
        for _, row in italy_region_from.iterrows():
            ticker = row['Ticker']
            desc = row['Description'][:60]
            norm_factor = row.get('Normalization factor', 1.0)
            region_to = row.get('Region to', '')
            print(f"     {ticker}: {desc}... (norm: {norm_factor}, to: {region_to})")
    else:
        print("   ❌ No series found with Region from = 'Italy'")
    
    # Check what the target Excel is actually using
    print(f"\n📋 READING TARGET EXCEL ITALY COLUMN:")
    
    try:
        target_sheet = pd.read_excel('DNB Markets EUROPEAN GAS BALANCE.xlsx', 
                                   sheet_name='Demand', header=None)
        
        # Find Italy column (should be column 3 based on earlier analysis)
        italy_col = 3
        italy_header = target_sheet.iloc[2, italy_col]  # Row 2 has headers
        
        if italy_header == 'Italy':
            print(f"   ✅ Italy column found at position {italy_col}")
            
            # Get some sample values
            sample_values = []
            for row in range(3, min(8, target_sheet.shape[0])):
                val = target_sheet.iloc[row, italy_col]
                if pd.notna(val):
                    sample_values.append(float(val))
            
            print(f"   Sample Italy values: {[f'{v:.2f}' for v in sample_values[:5]]}")
            print(f"   Average: {np.mean(sample_values):.2f}")
            
            # This tells us what the target should be
            target_italy_first = sample_values[0] if sample_values else None
            if target_italy_first:
                print(f"   🎯 Target Italy first value: {target_italy_first:.2f}")
                
                # Now let's see what we need to get this value
                print(f"\n💡 REVERSE ENGINEERING:")
                print(f"   If we have ~5 Italy series each contributing...")
                italy_series_count = len(italy_demand)
                if italy_series_count > 0:
                    per_series_target = target_italy_first / italy_series_count
                    print(f"   Per series target: ~{per_series_target:.2f}")
                    
                    # What normalization would give us this?
                    dummy_val = 500  # Assume dummy values average ~500
                    needed_norm = per_series_target / dummy_val
                    print(f"   Needed normalization: ~{needed_norm:.6f}")
                    
                    # Compare to actual normalization
                    actual_norm = italy_demand['Normalization factor'].iloc[0] if len(italy_demand) > 0 else None
                    if actual_norm:
                        print(f"   Actual normalization: {actual_norm:.6f}")
                        ratio = actual_norm / needed_norm if needed_norm != 0 else 0
                        print(f"   Ratio (actual/needed): {ratio:.2f}")
        
    except Exception as e:
        print(f"   ❌ Could not read target Excel: {e}")
    
    # Check normalization factors distribution
    print(f"\n📊 NORMALIZATION FACTORS ANALYSIS:")
    all_factors = dataset['Normalization factor'].dropna()
    print(f"   Total factors: {len(all_factors)}")
    print(f"   Unique factors: {len(all_factors.unique())}")
    print(f"   Range: {all_factors.min():.6f} to {all_factors.max():.6f}")
    
    # Most common factors
    factor_counts = all_factors.value_counts().head(10)
    print(f"   Most common factors:")
    for factor, count in factor_counts.items():
        print(f"     {factor:.9f}: {count} series")
    
    # The issue might be that we're using dummy data that's too high
    print(f"\n🔧 POTENTIAL SOLUTIONS:")
    print(f"   1. Use real Bloomberg data instead of dummy data")
    print(f"   2. Check if dummy data range is too high")
    print(f"   3. Verify we're using correct aggregation logic")
    print(f"   4. Check if normalization is being double-applied")

if __name__ == "__main__":
    debug_italy_calculation()