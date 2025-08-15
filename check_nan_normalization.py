# -*- coding: utf-8 -*-
"""
CHECK: Investigate NaN normalization factors that might cause missing data
"""

import pandas as pd
import numpy as np

def check_nan_normalization():
    """Check for NaN normalization factors"""
    
    print(f"\n{'='*80}")
    print("🔍 CHECKING NaN NORMALIZATION FACTORS")
    print(f"{'='*80}")
    
    # Read the dataset
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    
    # Check entries with NaN normalization factors
    nan_norm = dataset[dataset['Normalization factor'].isna()]
    print(f"⚠️  Entries with NaN normalization factors: {len(nan_norm)}")
    
    if len(nan_norm) > 0:
        print(f"\nNaN normalization entries:")
        for idx, row in nan_norm.iterrows():
            desc = str(row.get('Description', 'No description'))[:60]
            category = str(row.get('Category', 'No category'))
            ticker = str(row.get('Ticker', 'No ticker'))
            print(f"   Row {idx}: {desc}...")
            print(f"      Ticker: {ticker}")
            print(f"      Category: {category}")
            print()
    
    # Check what happens when we create normalization mapping
    print(f"🔧 CREATING NORMALIZATION MAPPING:")
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    print(f"   Total tickers in dataset: {len(dataset)}")
    print(f"   Tickers with valid normalization: {len(norm_factors)}")
    print(f"   Missing normalization: {len(dataset) - len(norm_factors)}")
    
    # Check specific countries
    print(f"\n🌍 CHECKING SPECIFIC COUNTRIES:")
    
    target_countries = ['Netherlands', 'Luxembourg']
    
    for country in target_countries:
        country_data = dataset[
            (dataset['Region from'] == country) | 
            (dataset['Region to'] == country) |
            (dataset['Description'].str.contains(country, case=False, na=False))
        ]
        
        print(f"\n🔍 {country}:")
        print(f"   Total entries: {len(country_data)}")
        
        # Check normalization factors
        valid_norm = country_data[pd.notna(country_data['Normalization factor'])]
        print(f"   With valid normalization: {len(valid_norm)}")
        
        nan_norm_country = country_data[pd.isna(country_data['Normalization factor'])]
        print(f"   With NaN normalization: {len(nan_norm_country)}")
        
        # Check demand entries specifically
        demand_entries = country_data[country_data['Category'] == 'Demand']
        print(f"   Demand entries: {len(demand_entries)}")
        
        demand_with_norm = demand_entries[pd.notna(demand_entries['Normalization factor'])]
        print(f"   Demand with valid normalization: {len(demand_with_norm)}")
        
        if len(demand_with_norm) > 0:
            print(f"   Sample demand entries with normalization:")
            for _, row in demand_with_norm.head(3).iterrows():
                desc = str(row['Description'])[:50]
                norm_val = row['Normalization factor']
                ticker = row['Ticker']
                print(f"      {desc}... = {norm_val} ({ticker})")
    
    return norm_factors

if __name__ == "__main__":
    factors = check_nan_normalization()