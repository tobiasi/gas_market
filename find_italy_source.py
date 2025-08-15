# -*- coding: utf-8 -*-
"""
INVESTIGATION: Find where our Italy values are actually coming from
"""

import pandas as pd

def find_italy_source():
    """Find the source of our Italy demand values"""
    
    print(f"\n{'='*80}")
    print("🔍 ITALY VALUE SOURCE INVESTIGATION")
    print(f"{'='*80}")
    
    # Read our MultiTicker to understand what Italy data we actually have
    try:
        our_mt = pd.read_csv('our_multiticker.csv', low_memory=False)
        
        print(f"📊 Our MultiTicker loaded: {our_mt.shape}")
        
        # Check all metadata rows to understand structure
        print(f"\n📋 MULTITICKER STRUCTURE:")
        for i in range(min(6, our_mt.shape[0])):
            print(f"   Row {i}: {str(our_mt.iloc[i, :5].values)[:100]}...")
        
        # Find where countries are listed (should be in one of the metadata rows)
        country_row = None
        category_row = None
        
        for row_idx in range(min(6, our_mt.shape[0])):
            row_data = our_mt.iloc[row_idx]
            # Look for Italy mentions
            italy_mentions = sum(1 for val in row_data if 'Italy' in str(val))
            print(f"   Row {row_idx} has {italy_mentions} Italy mentions")
            
            if italy_mentions > 0:
                if country_row is None:
                    country_row = row_idx
                    print(f"     → Using as country row")
        
        # Also check for category information
        for row_idx in range(min(6, our_mt.shape[0])):
            row_data = our_mt.iloc[row_idx]
            demand_mentions = sum(1 for val in row_data if 'Demand' in str(val))
            if demand_mentions > 5:  # Should have many demand series
                category_row = row_idx
                print(f"   Row {row_idx} has {demand_mentions} Demand mentions (category row)")
                break
        
        # Now find Italy series in our data
        if country_row is not None:
            print(f"\n🇮🇹 ITALY SERIES IN OUR DATA (Country Row {country_row}):")
            
            countries = our_mt.iloc[country_row]
            italy_columns = []
            
            for i, country in enumerate(countries):
                if 'Italy' in str(country):
                    series_name = our_mt.iloc[0, i] if i < len(our_mt.iloc[0]) else "Unknown"
                    category = our_mt.iloc[category_row, i] if category_row is not None and i < len(our_mt.iloc[category_row]) else "Unknown"
                    
                    italy_columns.append({
                        'column': i,
                        'name': series_name,
                        'country': country,
                        'category': category
                    })
                    
                    print(f"   Col {i}: {series_name}")
                    print(f"      Country: {country} | Category: {category}")
            
            if not italy_columns:
                print("   ❌ NO ITALY SERIES FOUND IN OUR MULTITICKER!")
                
                # If no Italy series, where are our Italy values coming from?
                print(f"\n🤔 MYSTERY: Our daily output shows Italy values but no Italy series!")
                print(f"   This suggests our Italy calculation is NOT coming from Italy-specific tickers")
                print(f"   Possible sources:")
                print(f"   1. Cross-border flows TO Italy")
                print(f"   2. Calculated/derived values")
                print(f"   3. Different country classification")
                print(f"   4. Error in our country aggregation logic")
        
        # Let's also check what countries we DO have
        print(f"\n🌍 COUNTRIES WE ACTUALLY HAVE:")
        if country_row is not None:
            countries = our_mt.iloc[country_row]
            unique_countries = set()
            
            for country in countries:
                country_str = str(country).strip()
                if country_str not in ['nan', '', 'Unknown', '#N/A'] and len(country_str) > 1:
                    unique_countries.add(country_str)
            
            sorted_countries = sorted(unique_countries)
            print(f"   Found {len(sorted_countries)} unique countries:")
            for i, country in enumerate(sorted_countries):
                print(f"   {i+1:2d}. {country}")
        
        # Check if we have any series that SHOULD be Italy but are named differently
        print(f"\n🔍 SEARCHING FOR ITALY-LIKE SERIES:")
        italy_like_terms = ['IT', 'ITA', 'Italian', 'italia', 'SNAM', 'TAP']  # Common Italy abbreviations
        
        series_names = our_mt.iloc[0] if our_mt.shape[0] > 0 else []
        
        for i, name in enumerate(series_names):
            name_str = str(name).lower()
            for term in italy_like_terms:
                if term.lower() in name_str:
                    country = our_mt.iloc[country_row, i] if country_row is not None else "Unknown"
                    category = our_mt.iloc[category_row, i] if category_row is not None else "Unknown"
                    print(f"   Col {i}: {name}")
                    print(f"      Contains '{term}' | Country: {country} | Category: {category}")
                    break
        
        return italy_columns if 'italy_columns' in locals() else []
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return []

def analyze_italy_calculation_logic():
    """Analyze how our current code calculates Italy values"""
    
    print(f"\n{'='*60}")
    print("🧮 ITALY CALCULATION LOGIC ANALYSIS") 
    print(f"{'='*60}")
    
    print(f"📝 Current Italy calculation logic:")
    print(f"   In gas_market_consolidated_use4.py line ~335:")
    print(f"   country_mask = (full_data.columns.get_level_values(2) == 'Demand') & ")
    print(f"                  (full_data.columns.get_level_values(3) == 'Italy')")
    print(f"   country_total = full_data.iloc[:, country_mask].sum(axis=1, skipna=False)")
    print(f"")
    print(f"💡 This means we're summing ALL series where:")
    print(f"   - Level 2 (Category) = 'Demand'")  
    print(f"   - Level 3 (Country) = 'Italy'")
    print(f"")
    print(f"🤔 But if we have 0 Italy series, this should return 0!")
    print(f"   The fact that we get ~117 suggests:")
    print(f"   1. We have series with Level 3 = 'Italy' that we're not seeing")
    print(f"   2. Or there's an error in our MultiIndex structure")
    print(f"   3. Or our aggregation is picking up wrong series")

if __name__ == "__main__":
    italy_series = find_italy_source()
    analyze_italy_calculation_logic()
    
    print(f"\n{'='*80}")
    print("🎯 RECOMMENDED NEXT STEPS:")
    print(f"{'='*80}")
    print(f"1. Add diagnose_italy_demand(full_data) to your main script")
    print(f"2. Run the script and check what series are being summed for Italy")
    print(f"3. Compare the MultiIndex structure between target and our data")
    print(f"4. Check if Bloomberg tickers are missing or miscategorized")