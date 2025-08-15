# -*- coding: utf-8 -*-
"""
TEST: Check if aggregation logic is working correctly with normalized data
"""

import pandas as pd
import numpy as np

def test_aggregation_logic():
    """Test if sum operations are working correctly"""
    
    print(f"\n{'='*80}")
    print("🧮 TESTING AGGREGATION LOGIC")
    print(f"{'='*80}")
    
    # Create test MultiIndex DataFrame similar to full_data
    np.random.seed(42)  # For reproducible results
    
    # Create test MultiIndex structure
    countries = ['Netherlands', 'Luxembourg', 'France', 'Italy']
    categories = ['Demand', 'Import', 'Production']
    
    # Create MultiIndex columns (similar to main script)
    level_0 = []  # Description
    level_1 = []  # Empty
    level_2 = []  # Category
    level_3 = []  # Region from
    level_4 = []  # Region to
    
    for country in countries:
        for category in categories:
            for i in range(2):  # 2 series per country-category combination
                level_0.append(f'{country} {category} Series {i+1}')
                level_1.append('')
                level_2.append(category)
                level_3.append(country)
                level_4.append(f'{category} Target')
    
    multi_index = pd.MultiIndex.from_arrays([level_0, level_1, level_2, level_3, level_4])
    
    # Create test data with known values
    test_data = pd.DataFrame(
        np.random.rand(100, len(multi_index)) * 100,  # Random data 0-100
        columns=multi_index
    )
    
    print(f"📊 Created test DataFrame: {test_data.shape}")
    print(f"   MultiIndex levels: {len(multi_index.levels)}")
    print(f"   Categories: {set(level_2)}")
    print(f"   Countries: {set(level_3)}")
    
    # Test the same aggregation logic as main script
    print(f"\n🧮 TESTING AGGREGATION FOR EACH COUNTRY:")
    
    country_results = {}
    
    for country in countries:
        # Same logic as main script
        country_mask = (test_data.columns.get_level_values(2) == 'Demand') & (test_data.columns.get_level_values(3) == country)
        country_total = test_data.iloc[:, country_mask].sum(axis=1, skipna=False)
        
        num_series_found = country_mask.sum()
        first_value = country_total.iloc[0]
        
        country_results[country] = {
            'series_count': num_series_found,
            'first_value': first_value,
            'is_nan': np.isnan(first_value),
            'is_zero': first_value == 0.0
        }
        
        print(f"   {country}:")
        print(f"      Series found: {num_series_found}")
        print(f"      First value: {first_value:.2f}")
        print(f"      Is NaN: {np.isnan(first_value)}")
        print(f"      Is Zero: {first_value == 0.0}")
    
    # Test with NaN values
    print(f"\n🧪 TESTING WITH NaN VALUES:")
    
    # Create a copy with some NaN values
    test_data_nan = test_data.copy()
    
    # Insert NaN values in Netherlands data
    netherlands_mask = (test_data_nan.columns.get_level_values(3) == 'Netherlands')
    netherlands_cols = test_data_nan.columns[netherlands_mask]
    
    # Set some values to NaN
    test_data_nan.iloc[0:10, netherlands_mask] = np.nan
    
    print(f"   Inserted NaN values in Netherlands data (first 10 rows)")
    
    # Test aggregation with NaN
    for country in ['Netherlands', 'Luxembourg']:
        country_mask = (test_data_nan.columns.get_level_values(2) == 'Demand') & (test_data_nan.columns.get_level_values(3) == country)
        
        # Test different skipna settings
        total_skip_na = test_data_nan.iloc[:, country_mask].sum(axis=1, skipna=True)
        total_no_skip = test_data_nan.iloc[:, country_mask].sum(axis=1, skipna=False)
        
        print(f"   {country}:")
        print(f"      With skipna=True: {total_skip_na.iloc[0]:.2f}")
        print(f"      With skipna=False: {total_no_skip.iloc[0]:.2f}")
        print(f"      skipna=True is NaN: {np.isnan(total_skip_na.iloc[0])}")
        print(f"      skipna=False is NaN: {np.isnan(total_no_skip.iloc[0])}")
    
    # Test normalization impact
    print(f"\n🔧 TESTING NORMALIZATION IMPACT:")
    
    # Apply normalization factors similar to main script
    normalization_factors = {
        'Netherlands': 0.090909091,  # 1/11
        'Luxembourg': 0.090909091,   # 1/11
        'France': 0.090909091,       # 1/11
        'Italy': 0.090909091         # 1/11
    }
    
    test_data_normalized = test_data.copy()
    
    # Apply normalization (simulate what main script does)
    for col_idx, col in enumerate(test_data_normalized.columns):
        country = col[3]  # Region from is level 3
        if country in normalization_factors:
            factor = normalization_factors[country]
            test_data_normalized.iloc[:, col_idx] = test_data_normalized.iloc[:, col_idx] * factor
    
    print(f"   Applied normalization factors")
    
    # Test aggregation after normalization
    for country in ['Netherlands', 'Luxembourg']:
        country_mask = (test_data_normalized.columns.get_level_values(2) == 'Demand') & (test_data_normalized.columns.get_level_values(3) == country)
        
        original_total = test_data.iloc[:, country_mask].sum(axis=1, skipna=False).iloc[0]
        normalized_total = test_data_normalized.iloc[:, country_mask].sum(axis=1, skipna=False).iloc[0]
        
        expected_total = original_total * normalization_factors[country]
        
        print(f"   {country}:")
        print(f"      Original: {original_total:.2f}")
        print(f"      Normalized: {normalized_total:.2f}")
        print(f"      Expected: {expected_total:.2f}")
        print(f"      Match: {'✅' if abs(normalized_total - expected_total) < 0.01 else '❌'}")

if __name__ == "__main__":
    test_aggregation_logic()