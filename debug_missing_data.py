# -*- coding: utf-8 -*-
"""
DEBUG: Check why Netherlands, Luxembourg, totals are missing after normalization
"""

import pandas as pd
import numpy as np

def debug_missing_data():
    """Debug missing data issues after normalization"""
    
    print(f"\n{'='*80}")
    print("🔍 DEBUGGING MISSING DATA AFTER NORMALIZATION")
    print(f"{'='*80}")
    
    # Read the use4.xlsx data like the main script
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    
    print(f"📊 Dataset loaded: {dataset.shape}")
    
    # Check what countries are in the data
    print(f"\n🌍 COUNTRIES IN DATASET:")
    countries_in_data = dataset['Region from'].value_counts()
    print("Region from counts:")
    for country, count in countries_in_data.head(15).items():
        print(f"   {country}: {count}")
    
    region_to_counts = dataset['Region to'].value_counts()
    print(f"\nRegion to counts:")
    for country, count in region_to_counts.head(15).items():
        print(f"   {country}: {count}")
    
    # Check categories
    print(f"\n📋 CATEGORIES IN DATASET:")
    categories = dataset['Category'].value_counts()
    for cat, count in categories.items():
        print(f"   {cat}: {count}")
    
    # Check specific problem areas
    print(f"\n🔍 CHECKING PROBLEM AREAS:")
    
    # Netherlands
    netherlands_data = dataset[
        (dataset['Region from'] == 'Netherlands') | 
        (dataset['Region to'] == 'Netherlands') |
        (dataset['Description'].str.contains('Netherlands', case=False, na=False))
    ]
    print(f"\n🇳🇱 Netherlands entries: {len(netherlands_data)}")
    if len(netherlands_data) > 0:
        print("Netherlands categories:")
        print(netherlands_data['Category'].value_counts())
    
    # Luxembourg  
    luxembourg_data = dataset[
        (dataset['Region from'] == 'Luxembourg') | 
        (dataset['Region to'] == 'Luxembourg') |
        (dataset['Description'].str.contains('Luxembourg', case=False, na=False))
    ]
    print(f"\n🇱🇺 Luxembourg entries: {len(luxembourg_data)}")
    if len(luxembourg_data) > 0:
        print("Luxembourg categories:")
        print(luxembourg_data['Category'].value_counts())
        print("Luxembourg sample:")
        for _, row in luxembourg_data.head(3).iterrows():
            print(f"   {row['Description'][:60]}...")
            print(f"   Category: {row['Category']}, Region from: {row['Region from']}, Region to: {row['Region to']}")
    
    # Check normalization factors for these countries
    print(f"\n🔧 NORMALIZATION FACTORS:")
    
    # Netherlands normalization
    netherlands_norm = netherlands_data['Normalization factor'].dropna().unique()
    print(f"Netherlands normalization factors: {netherlands_norm}")
    
    # Luxembourg normalization
    luxembourg_norm = luxembourg_data['Normalization factor'].dropna().unique()
    print(f"Luxembourg normalization factors: {luxembourg_norm}")
    
    # Check if any data has NaN normalization factors
    nan_norm = dataset[dataset['Normalization factor'].isna()]
    print(f"\n⚠️  Entries with NaN normalization factors: {len(nan_norm)}")
    
    if len(nan_norm) > 0:
        print("Sample entries with NaN normalization:")
        for _, row in nan_norm.head(5).iterrows():
            print(f"   {row['Description'][:60]}...")
            print(f"   Category: {row['Category']}")
    
    # Check the countries list that the script is using
    expected_countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    print(f"\n📝 EXPECTED vs ACTUAL COUNTRY MATCHING:")
    for country in expected_countries:
        # Check exact matches
        exact_from = len(dataset[dataset['Region from'] == country])
        exact_to = len(dataset[dataset['Region to'] == country])
        desc_contains = len(dataset[dataset['Description'].str.contains(country, case=False, na=False)])
        
        print(f"   {country}:")
        print(f"      Region from: {exact_from}")
        print(f"      Region to: {exact_to}")
        print(f"      In description: {desc_contains}")
        
        if exact_from == 0 and exact_to == 0 and desc_contains == 0:
            print(f"      ❌ NO DATA FOUND!")
    
    return dataset

if __name__ == "__main__":
    data = debug_missing_data()