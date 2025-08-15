# -*- coding: utf-8 -*-
"""
CHECK: What scaling factors should Italy series have?
"""

import pandas as pd

def check_italy_scaling():
    """Check scaling factors for Italy series"""
    
    print(f"\n{'='*80}")
    print("🔍 ITALY SCALING FACTORS INVESTIGATION")
    print(f"{'='*80}")
    
    try:
        # Read target MultiTicker
        target_mt = pd.read_csv('the_multiticker_we_are_trying_to_replicate.csv', low_memory=False)
        
        print(f"📊 Target MultiTicker shape: {target_mt.shape}")
        
        # Find scaling factor row (should be row 2)
        scaling_row = 2
        if target_mt.shape[0] > scaling_row:
            print(f"\n🔍 ROW {scaling_row} (Scaling Factor):")
            scaling_factors = target_mt.iloc[scaling_row]
            print(f"   First 10 scaling factors: {scaling_factors.iloc[:10].tolist()}")
        
        # Find Italy series in target
        italy_columns = []
        if target_mt.shape[0] > 4:  # Row 4 should be Metadata 2 (countries)
            for i, country in enumerate(target_mt.iloc[4]):
                if 'Italy' in str(country):
                    series_name = target_mt.iloc[0, i] if i < len(target_mt.iloc[0]) else "Unknown"
                    scaling_factor = target_mt.iloc[scaling_row, i] if i < len(target_mt.iloc[scaling_row]) else "Unknown"
                    category = target_mt.iloc[3, i] if i < len(target_mt.iloc[3]) else "Unknown"
                    
                    italy_columns.append({
                        'column': i,
                        'name': series_name,
                        'country': country,
                        'category': category,
                        'scaling_factor': scaling_factor
                    })
        
        print(f"\n🇮🇹 ITALY SERIES SCALING FACTORS:")
        print(f"   Found {len(italy_columns)} Italy series")
        
        for i, info in enumerate(italy_columns):
            print(f"   {i+1:2d}. {info['name'][:50]}...")
            print(f"        Country: {info['country']} | Category: {info['category']} | Scaling: {info['scaling_factor']}")
        
        # Check if all Italy DEMAND series have same scaling factor
        demand_italy = [info for info in italy_columns if info['category'] == 'Demand']
        print(f"\n🎯 ITALY DEMAND SERIES SCALING:")
        print(f"   Italy DEMAND series: {len(demand_italy)}")
        
        if demand_italy:
            scaling_factors = [info['scaling_factor'] for info in demand_italy]
            unique_scalings = list(set(str(s) for s in scaling_factors))
            print(f"   Unique scaling factors: {unique_scalings}")
            
            if len(unique_scalings) == 1:
                common_scaling = unique_scalings[0]
                print(f"   ✅ All Italy DEMAND series use scaling factor: {common_scaling}")
                
                # Test the scaling
                if common_scaling != 'Unknown':
                    try:
                        factor = float(common_scaling)
                        our_italy = 1291.68  # From debug output
                        scaled_italy = our_italy * factor
                        print(f"   📊 SCALING TEST:")
                        print(f"      Our Italy value: {our_italy}")
                        print(f"      Scaling factor: {factor}")
                        print(f"      Scaled result: {scaled_italy:.2f}")
                        print(f"      Expected (~117): Close? {'✅' if abs(scaled_italy - 117) < 20 else '❌'}")
                    except ValueError:
                        print(f"   ❌ Could not convert scaling factor to number: {common_scaling}")
            else:
                print(f"   ⚠️  Different scaling factors for Italy DEMAND series")
                for info in demand_italy:
                    print(f"      {info['scaling_factor']}: {info['name'][:40]}...")
        
        return italy_columns
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return None

def check_use4_normalization():
    """Check normalization factors in use4.xlsx"""
    
    print(f"\n{'='*60}")
    print("📊 USE4.XLSX NORMALIZATION FACTORS")
    print(f"{'='*60}")
    
    try:
        # Read use4.xlsx
        use4 = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
        
        print(f"📋 Use4 shape: {use4.shape}")
        
        # Find Italy entries
        italy_entries = use4[use4['Description'].str.contains('Italy', case=False, na=False)]
        
        print(f"\n🇮🇹 ITALY ENTRIES IN USE4:")
        print(f"   Found {len(italy_entries)} Italy entries")
        
        if len(italy_entries) > 0:
            print(f"   {'Description':<50} {'Norm Factor':<12} {'Category':<15}")
            print("-" * 80)
            
            for _, row in italy_entries.iterrows():
                desc = str(row['Description'])[:47] + "..." if len(str(row['Description'])) > 47 else str(row['Description'])
                norm_factor = row.get('Normalization factor', 'N/A')
                category = row.get('Category', 'N/A')
                print(f"   {desc:<50} {norm_factor:<12} {category:<15}")
        
        # Check if normalization factors are consistent
        if len(italy_entries) > 0 and 'Normalization factor' in italy_entries.columns:
            norm_factors = italy_entries['Normalization factor'].dropna().unique()
            print(f"\n📊 ITALY NORMALIZATION FACTORS:")
            print(f"   Unique factors: {norm_factors}")
            
            if len(norm_factors) == 1:
                factor = norm_factors[0]
                print(f"   ✅ All Italy entries use factor: {factor}")
                
                # Test scaling
                try:
                    factor_float = float(factor)
                    our_italy = 1291.68
                    scaled_italy = our_italy * factor_float
                    print(f"   📊 NORMALIZATION TEST:")
                    print(f"      Our Italy value: {our_italy}")
                    print(f"      Normalization factor: {factor_float}")
                    print(f"      Scaled result: {scaled_italy:.2f}")
                    print(f"      Expected (~117): Close? {'✅' if abs(scaled_italy - 117) < 20 else '❌'}")
                except ValueError:
                    print(f"   ❌ Could not convert factor to number: {factor}")
        
        return italy_entries
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

if __name__ == "__main__":
    target_italy = check_italy_scaling()
    use4_italy = check_use4_normalization()
    
    print(f"\n{'='*80}")
    print("🎯 SCALING CONCLUSION")
    print(f"{'='*80}")
    print(f"We need to apply scaling/normalization factors to match the target!")
    print(f"Check which method (target scaling vs use4 normalization) gives ~117")