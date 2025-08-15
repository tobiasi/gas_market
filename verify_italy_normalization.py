# -*- coding: utf-8 -*-
"""
VERIFY: Check Italy normalization factors from tickerlist_tab.csv
"""

import pandas as pd

def verify_italy_normalization():
    """Verify Italy normalization factors"""
    
    print(f"\n{'='*80}")
    print("🎯 ITALY NORMALIZATION VERIFICATION")
    print(f"{'='*80}")
    
    # Read the ticker list
    ticker_list = pd.read_csv('tickerlist_tab.csv', skiprows=8)
    
    print(f"📊 Ticker list loaded: {ticker_list.shape}")
    
    # Find Italy DEMAND series
    italy_demand = ticker_list[
        (ticker_list['Description'].str.contains('Italy', case=False, na=False)) &
        (ticker_list['Category'] == 'Demand')
    ]
    
    print(f"\n🇮🇹 ITALY DEMAND SERIES:")
    print(f"   Found {len(italy_demand)} Italy DEMAND series")
    
    italy_values = {
        'Italy Gas Consumption Losses': 12.29,
        'Italy Gas Industrial Demand': 358.90,
        'Italy Gas Domestic Demand': 331.17,
        'Italy Gas Power Generation Demand': 560.01,
        'Italy Gas Total Exports': 29.31
    }
    
    print(f"\n📋 ITALY DEMAND SERIES WITH NORMALIZATION:")
    total_normalized = 0
    total_raw = 0
    
    for _, row in italy_demand.iterrows():
        desc = row['Description']
        norm_factor = row['Normalization factor']
        ticker = row['Ticker']
        
        # Match with our debug values
        raw_value = 0
        for key, value in italy_values.items():
            if key in desc:
                raw_value = value
                break
        
        normalized_value = raw_value * norm_factor
        total_normalized += normalized_value
        total_raw += raw_value
        
        print(f"   {desc}")
        print(f"     Ticker: {ticker}")
        print(f"     Raw value: {raw_value:.2f}")
        print(f"     Norm factor: {norm_factor}")
        print(f"     Normalized: {normalized_value:.2f}")
        print()
    
    print(f"📊 TOTALS:")
    print(f"   Raw Italy total: {total_raw:.2f}")
    print(f"   Normalized Italy total: {total_normalized:.2f}")
    print(f"   Expected (~117): {'✅ MATCH!' if abs(total_normalized - 117) < 5 else '❌ Different'}")
    
    # Check the normalization factor
    common_factor = italy_demand['Normalization factor'].iloc[0] if len(italy_demand) > 0 else None
    print(f"\n🔍 NORMALIZATION FACTOR:")
    print(f"   Common factor: {common_factor}")
    print(f"   This is: 1/{1/common_factor:.1f} = {common_factor:.9f}")
    
    return italy_demand

if __name__ == "__main__":
    italy_data = verify_italy_normalization()