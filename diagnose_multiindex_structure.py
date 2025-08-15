# -*- coding: utf-8 -*-
"""
DIAGNOSE: Check MultiIndex structure after normalization to find missing data
"""

import pandas as pd
import numpy as np
from collections import defaultdict

def diagnose_multiindex_structure():
    """Diagnose MultiIndex structure issues"""
    
    print(f"\n{'='*80}")
    print("🔍 DIAGNOSING MULTIINDEX STRUCTURE AFTER NORMALIZATION")
    print(f"{'='*80}")
    
    # Simulate the same data loading as the main script
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    
    # Create normalization factors
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    print(f"📊 Created normalization mapping for {len(norm_factors)} tickers")
    
    # Simulate the MultiIndex creation (without Bloomberg data)
    level_0 = []  # Description
    level_1 = []  # Empty
    level_2 = []  # Category  
    level_3 = []  # Region from
    level_4 = []  # Region to
    
    successful_tickers = []
    
    for index, row in dataset.iterrows():
        ticker = row.get('Ticker')
        
        if pd.isna(ticker):
            continue
            
        # Add to successful list (simulated)
        successful_tickers.append(ticker)
        
        # Build MultiIndex levels
        level_0.append(row.get('Description', ''))
        level_1.append('')  # Empty level as in original structure
        level_2.append(row.get('Category', ''))
        level_3.append(row.get('Region from', ''))
        level_4.append(row.get('Region to', ''))
    
    print(f"✅ Processed {len(successful_tickers)} tickers for MultiIndex")
    
    # Create MultiIndex (simulated)
    if successful_tickers:
        multi_index = pd.MultiIndex.from_arrays([level_0, level_1, level_2, level_3, level_4])
        print(f"✅ Created MultiIndex with {len(multi_index)} columns")
        
        # Analyze the MultiIndex structure
        print(f"\n📋 MULTIINDEX ANALYSIS:")
        print(f"   Level 0 (Description): {len(set(level_0))} unique values")
        print(f"   Level 1 (Empty): {len(set(level_1))} unique values") 
        print(f"   Level 2 (Category): {len(set(level_2))} unique values")
        print(f"   Level 3 (Region from): {len(set(level_3))} unique values")
        print(f"   Level 4 (Region to): {len(set(level_4))} unique values")
        
        # Check categories
        categories = set(level_2)
        print(f"\n📊 CATEGORIES IN MULTIINDEX:")
        for cat in sorted(categories):
            count = level_2.count(cat)
            print(f"   {cat}: {count} series")
        
        # Check countries in Region from (level 3)
        countries_from = set(level_3)
        print(f"\n🌍 COUNTRIES IN 'REGION FROM' (Level 3):")
        for country in sorted(countries_from):
            if country and country != '':
                count = level_3.count(country)
                print(f"   {country}: {count} series")
        
        # Check countries in Region to (level 4)
        countries_to = set(level_4)
        print(f"\n🌍 COUNTRIES IN 'REGION TO' (Level 4):")
        for country in sorted(countries_to):
            if country and country != '':
                count = level_4.count(country)
                print(f"   {country}: {count} series")
        
        # Check specific combinations for target countries
        target_countries = ['Netherlands', 'Luxembourg']
        
        print(f"\n🎯 TARGET COUNTRIES ANALYSIS:")
        for country in target_countries:
            print(f"\n🔍 {country}:")
            
            # Check for exact matches in Region from
            from_matches = [i for i, x in enumerate(level_3) if x == country]
            print(f"   Region from matches: {len(from_matches)}")
            
            # Check for exact matches in Region to  
            to_matches = [i for i, x in enumerate(level_4) if x == country]
            print(f"   Region to matches: {len(to_matches)}")
            
            # Check for DEMAND category specifically
            demand_from = [i for i in from_matches if level_2[i] == 'Demand']
            demand_to = [i for i in to_matches if level_2[i] == 'Demand']
            
            print(f"   DEMAND + Region from: {len(demand_from)}")
            print(f"   DEMAND + Region to: {len(demand_to)}")
            
            if len(demand_from) > 0:
                print(f"   Sample DEMAND series (Region from):")
                for i in demand_from[:3]:
                    desc = level_0[i][:60] + "..." if len(level_0[i]) > 60 else level_0[i]
                    print(f"      {desc}")
            
            if len(demand_to) > 0:
                print(f"   Sample DEMAND series (Region to):")
                for i in demand_to[:3]:
                    desc = level_0[i][:60] + "..." if len(level_0[i]) > 60 else level_0[i]
                    print(f"      {desc}")
        
        # Check the specific logic used in the main script
        print(f"\n🧮 SIMULATING MAIN SCRIPT LOGIC:")
        country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
        
        for country in country_list:
            # Simulate: (level_2 == 'Demand') & (level_3 == country)
            matches = []
            for i in range(len(level_2)):
                if level_2[i] == 'Demand' and level_3[i] == country:
                    matches.append(i)
            
            print(f"   {country}: {len(matches)} DEMAND series where Region from = '{country}'")
            
            if len(matches) == 0 and country in ['Netherlands', 'Luxembourg']:
                print(f"   ⚠️  PROBLEM: No matches found for {country}!")
                
                # Check alternative matching strategies
                alt_matches = []
                for i in range(len(level_2)):
                    if level_2[i] == 'Demand' and level_4[i] == country:
                        alt_matches.append(i)
                
                print(f"   Alternative (Region to): {len(alt_matches)} matches")
                
                # Check for partial matches in description
                desc_matches = []
                for i in range(len(level_0)):
                    if level_2[i] == 'Demand' and country.lower() in level_0[i].lower():
                        desc_matches.append(i)
                
                print(f"   Description contains '{country}': {len(desc_matches)} matches")

if __name__ == "__main__":
    diagnose_multiindex_structure()