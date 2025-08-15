# -*- coding: utf-8 -*-
"""
INVESTIGATION SCRIPT: Understand why demand totals don't match
"""

# Add this code to the END of your gas_market_consolidated_use4.py script
# or run it separately after the main script completes

def investigate_demand_differences(full_data, countries, industrial_total, ldz_total, gtp_total):
    """Deep dive into why demand totals don't match"""
    
    print(f"\n{'='*80}")
    print("🔍 DETAILED INVESTIGATION OF DEMAND DIFFERENCES")
    print(f"{'='*80}")
    
    # 1. Analyze ALL demand series
    print(f"\n📊 STEP 1: ANALYZE ALL DEMAND SERIES")
    all_demand_mask = full_data.columns.get_level_values(2) == 'Demand'
    all_demand_series = full_data.iloc[:, all_demand_mask]
    
    print(f"   Total 'Demand' series in dataset: {all_demand_series.shape[1]}")
    print(f"   Sample series names:")
    for i, col in enumerate(all_demand_series.columns[:10]):
        desc = col[0][:60] + "..." if len(col[0]) > 60 else col[0]
        region = col[3] if len(col) > 3 else "Unknown"
        sector = col[4] if len(col) > 4 else "Unknown"
        print(f"     {i+1:2d}. {desc} | {region} | {sector}")
    
    if all_demand_series.shape[1] > 10:
        print(f"     ... and {all_demand_series.shape[1] - 10} more")
    
    # 2. Group by country to see what we're missing
    print(f"\n🌍 STEP 2: GROUP DEMAND SERIES BY COUNTRY")
    countries_in_data = full_data.columns.get_level_values(3).unique()
    countries_in_data = [c for c in countries_in_data if c not in ['', 'nan', None]]
    
    country_totals = {}
    for country in sorted(countries_in_data):
        country_mask = ((full_data.columns.get_level_values(2) == 'Demand') & 
                       (full_data.columns.get_level_values(3) == country))
        if country_mask.any():
            country_total = full_data.iloc[:, country_mask].sum(axis=1, skipna=False)
            country_totals[country] = {
                'series_count': country_mask.sum(),
                'sample_value': country_total.iloc[0],
                'series_names': full_data.columns[country_mask].get_level_values(0).tolist()
            }
    
    print(f"   Countries found in demand data:")
    total_series_accounted = 0
    for country, data in country_totals.items():
        print(f"     {country}: {data['series_count']} series, value={data['sample_value']:.1f}")
        total_series_accounted += data['series_count']
        
    print(f"   Total series accounted for: {total_series_accounted} / {all_demand_series.shape[1]}")
    
    # 3. Group by sector to see category breakdown
    print(f"\n🏭 STEP 3: GROUP DEMAND SERIES BY SECTOR/CATEGORY")
    sectors_in_data = full_data.columns.get_level_values(4).unique()
    sectors_in_data = [s for s in sectors_in_data if s not in ['', 'nan', None]]
    
    sector_totals = {}
    for sector in sorted(sectors_in_data):
        sector_mask = ((full_data.columns.get_level_values(2) == 'Demand') & 
                      (full_data.columns.get_level_values(4) == sector))
        if sector_mask.any():
            sector_total = full_data.iloc[:, sector_mask].sum(axis=1, skipna=False)
            sector_totals[sector] = {
                'series_count': sector_mask.sum(),
                'sample_value': sector_total.iloc[0],
                'series_names': full_data.columns[sector_mask].get_level_values(0).tolist()[:5]  # First 5
            }
    
    print(f"   Sectors found in demand data:")
    for sector, data in sector_totals.items():
        print(f"     {sector}: {data['series_count']} series, value={data['sample_value']:.1f}")
        print(f"       Examples: {', '.join(data['series_names'][:3])}...")
    
    # 4. Find the missing series
    print(f"\n🔍 STEP 4: IDENTIFY MISSING SERIES")
    
    # What countries are we including in our country_total?
    country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    included_countries = set(country_list)
    all_countries = set(country_totals.keys())
    missing_countries = all_countries - included_countries
    
    print(f"   Countries we're including: {sorted(included_countries)}")
    print(f"   Countries in data but NOT included: {sorted(missing_countries)}")
    
    if missing_countries:
        print(f"   Missing countries breakdown:")
        missing_total = 0
        for country in sorted(missing_countries):
            if country in country_totals:
                data = country_totals[country]
                missing_total += data['sample_value']
                print(f"     {country}: {data['series_count']} series, value={data['sample_value']:.1f}")
        print(f"   Total value from missing countries: {missing_total:.1f}")
    
    # 5. Check category keywords coverage
    print(f"\n🔎 STEP 5: CHECK CATEGORY KEYWORD COVERAGE")
    
    # Check how many series our keywords are capturing
    industrial_keywords = ['Industrial', 'industry', 'industrial and power']
    ldz_keywords = ['LDZ', 'ldz', 'Low Distribution Zone']  
    gtp_keywords = ['Gas-to-Power', 'gas-to-power', 'Power', 'electricity']
    
    all_keywords = industrial_keywords + ldz_keywords + gtp_keywords
    
    captured_series = set()
    for keyword in all_keywords:
        mask = full_data.columns.get_level_values(0).str.contains(keyword, case=False, na=False)
        captured_series.update(full_data.columns[mask].get_level_values(0))
    
    all_demand_descriptions = set(all_demand_series.columns.get_level_values(0))
    uncaptured_series = all_demand_descriptions - captured_series
    
    print(f"   Total demand series: {len(all_demand_descriptions)}")
    print(f"   Captured by keywords: {len(captured_series)}")
    print(f"   NOT captured: {len(uncaptured_series)}")
    
    if uncaptured_series:
        print(f"   Uncaptured series examples:")
        for i, series in enumerate(sorted(uncaptured_series)[:10]):
            print(f"     {i+1:2d}. {series[:80]}...")
    
    # 6. Recommendation
    print(f"\n💡 RECOMMENDATIONS:")
    if missing_countries:
        print(f"   1. Add missing countries to country_list: {sorted(missing_countries)}")
    if uncaptured_series:
        print(f"   2. Add keywords to capture uncaptured series")
        print(f"   3. Review if uncaptured series should be included in totals")
    
    print(f"\n📝 SUMMARY:")
    print(f"   The difference likely comes from:")
    print(f"   - Countries not included in country_list")
    print(f"   - Demand series not captured by category keywords")
    print(f"   - Different aggregation scopes (country vs category)")

# Add this to your main script:
print("\n" + "="*80)
print("🔍 To investigate the differences, add this to your script:")
print("investigate_demand_differences(full_data, countries, industrial_total, ldz_total, gtp_total)")
print("="*80)