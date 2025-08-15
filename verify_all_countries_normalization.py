# -*- coding: utf-8 -*-
"""
VERIFY: Check normalization factors for all countries
"""

import pandas as pd
import numpy as np

def verify_all_normalization():
    """Verify normalization factors for all countries"""
    
    print(f"\n{'='*80}")
    print("🌍 ALL COUNTRIES NORMALIZATION VERIFICATION")
    print(f"{'='*80}")
    
    # Read the ticker list
    ticker_list = pd.read_csv('tickerlist_tab.csv', skiprows=8)
    
    print(f"📊 Ticker list loaded: {ticker_list.shape}")
    
    # Define countries to check
    countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Germany', 'Austria', 'Switzerland', 'Luxembourg']
    
    print(f"\n🔍 NORMALIZATION FACTORS BY COUNTRY:")
    print(f"{'Country':<12} {'Series':<6} {'Factors':<30} {'Impact':<20}")
    print("-" * 80)
    
    normalization_summary = {}
    
    for country in countries:
        # Find series for this country
        country_mask = ticker_list['Description'].str.contains(country, case=False, na=False)
        country_data = ticker_list[country_mask]
        
        if len(country_data) > 0:
            # Get unique normalization factors
            norm_factors = country_data['Normalization factor'].dropna().unique()
            norm_factors_str = ', '.join([f'{f:.6f}' for f in sorted(norm_factors)])
            
            # Determine impact
            has_scaling = any(abs(f - 1.0) > 0.001 for f in norm_factors if not np.isnan(f))
            impact = "🔧 SCALED" if has_scaling else "✅ No scaling"
            
            print(f"{country:<12} {len(country_data):<6} {norm_factors_str:<30} {impact:<20}")
            
            normalization_summary[country] = {
                'series_count': len(country_data),
                'factors': norm_factors,
                'has_scaling': has_scaling
            }
        else:
            print(f"{country:<12} {0:<6} {'No series found':<30} {'N/A':<20}")
            normalization_summary[country] = {
                'series_count': 0,
                'factors': [],
                'has_scaling': False
            }
    
    print(f"\n📋 DETAILED ANALYSIS:")
    
    for country, info in normalization_summary.items():
        if info['has_scaling']:
            print(f"\n🔧 {country} - NORMALIZATION APPLIED:")
            country_series = ticker_list[ticker_list['Description'].str.contains(country, case=False, na=False)]
            
            for _, row in country_series.iterrows():
                desc = str(row['Description'])[:60]
                norm_factor = row['Normalization factor']
                category = row.get('Category', 'Unknown')
                
                if pd.notna(norm_factor) and abs(norm_factor - 1.0) > 0.001:
                    scaling_desc = f"1/{1/norm_factor:.1f}" if norm_factor != 0 else "Zero"
                    print(f"   📊 {desc}...")
                    print(f"      Factor: {norm_factor:.9f} ({scaling_desc})")
                    print(f"      Category: {category}")
                    print()
    
    print(f"\n💡 SUMMARY:")
    scaled_countries = [c for c, info in normalization_summary.items() if info['has_scaling']]
    unscaled_countries = [c for c, info in normalization_summary.items() if info['series_count'] > 0 and not info['has_scaling']]
    no_series_countries = [c for c, info in normalization_summary.items() if info['series_count'] == 0]
    
    print(f"   🔧 Countries with scaling: {', '.join(scaled_countries)}")
    print(f"   ✅ Countries without scaling: {', '.join(unscaled_countries)}")
    print(f"   ❓ Countries with no series: {', '.join(no_series_countries)}")
    
    print(f"\n🎯 CONCLUSION:")
    if scaled_countries:
        print(f"   ✅ Normalization will fix scaling issues in {len(scaled_countries)} countries")
        print(f"   🔧 The normalized script applies factors to ALL Bloomberg data")
        print(f"   🎯 This should resolve country vs category discrepancies across Europe")
    else:
        print(f"   ⚠️  No scaling factors found - normalization may not help")
    
    return normalization_summary

if __name__ == "__main__":
    summary = verify_all_normalization()