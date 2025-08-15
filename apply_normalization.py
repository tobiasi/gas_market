# -*- coding: utf-8 -*-
"""
APPLY NORMALIZATION: Test applying normalization factors to Italy series
"""

import pandas as pd
import numpy as np

def test_normalization():
    """Test applying normalization factors to see if we get ~117"""
    
    print(f"\n{'='*80}")
    print("🔧 NORMALIZATION FACTOR TEST")
    print(f"{'='*80}")
    
    # Our current Italy values from debug output
    italy_values = {
        'Italy Gas Consumption Losses': 12.29,
        'Italy Gas Industrial Demand': 358.90,
        'Italy Gas Domestic Demand': 331.17,
        'Italy Gas Power Generation Demand': 560.01,
        'Italy Gas Total Exports': 29.31
    }
    
    raw_total = sum(italy_values.values())
    print(f"📊 CURRENT ITALY CALCULATION:")
    print(f"   Raw total: {raw_total:.2f}")
    
    # Test different normalization factors
    test_factors = [0.1, 0.09, 0.0905, 0.095, 0.085]
    
    print(f"\n🧪 NORMALIZATION FACTOR TESTS:")
    print(f"{'Factor':<8} {'Result':<10} {'Match ~117?':<12}")
    print("-" * 35)
    
    for factor in test_factors:
        normalized_total = raw_total * factor
        match = "✅" if abs(normalized_total - 117) < 5 else "❌"
        print(f"{factor:<8} {normalized_total:<10.2f} {match:<12}")
    
    # Try to read normalization factors from use4.xlsx
    print(f"\n📋 READING USE4.XLSX NORMALIZATION FACTORS:")
    try:
        # Try different methods to read use4.xlsx
        use4 = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8, engine='openpyxl')
        
        # Find Italy entries
        italy_mask = use4['Description'].str.contains('Italy', case=False, na=False)
        italy_entries = use4[italy_mask]
        
        print(f"   Found {len(italy_entries)} Italy entries")
        
        if 'Normalization factor' in italy_entries.columns:
            print(f"   Italy normalization factors:")
            for _, row in italy_entries.iterrows():
                desc = str(row['Description'])[:50]
                norm_factor = row['Normalization factor']
                print(f"     {desc}... = {norm_factor}")
            
            # Calculate weighted normalization
            if len(italy_entries) > 0:
                factors = italy_entries['Normalization factor'].dropna()
                if len(factors) > 0:
                    avg_factor = factors.mean()
                    print(f"\n   Average normalization factor: {avg_factor:.6f}")
                    
                    normalized_with_avg = raw_total * avg_factor
                    print(f"   Italy with avg factor: {normalized_with_avg:.2f}")
                    print(f"   Close to 117? {'✅' if abs(normalized_with_avg - 117) < 10 else '❌'}")
        
    except Exception as e:
        print(f"   ❌ Could not read use4.xlsx: {e}")
        print(f"   Try manually checking normalization factors in Excel")
    
    print(f"\n💡 SOLUTION:")
    print(f"   1. Read 'Normalization factor' column from use4.xlsx") 
    print(f"   2. Apply factors to each series BEFORE aggregation")
    print(f"   3. This should convert 1,291.68 → ~117")
    print(f"")
    print(f"🔧 IMPLEMENTATION:")
    print(f"   In gas_market_consolidated_use4.py:")
    print(f"   1. Read normalization factors from dataset")
    print(f"   2. Apply factor to each ticker: data[ticker] *= factor")
    print(f"   3. Then proceed with normal aggregation")

if __name__ == "__main__":
    test_normalization()