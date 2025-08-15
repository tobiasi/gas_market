# -*- coding: utf-8 -*-
"""
COMPARISON: Compare Italy series between target and our MultiTicker data
"""

import pandas as pd

def compare_italy_multiticker():
    """Compare Italy series in MultiTicker files"""
    
    print(f"\n{'='*80}")
    print("🇮🇹 ITALY MULTITICKER COMPARISON")
    print(f"{'='*80}")
    
    try:
        # Read MultiTicker files
        target_mt = pd.read_csv('the_multiticker_we_are_trying_to_replicate.csv')
        our_mt = pd.read_csv('our_multiticker.csv')
        
        print(f"📊 MultiTicker files loaded:")
        print(f"   Target MultiTicker shape: {target_mt.shape}")
        print(f"   Our MultiTicker shape: {our_mt.shape}")
        
        # Find Italy-related series in TARGET MultiTicker
        print(f"\n🎯 ITALY SERIES IN TARGET MULTITICKER:")
        
        # Look in the Metadata 2 row (row 4) for Italy entries
        if target_mt.shape[0] > 4:
            target_countries = target_mt.iloc[4]  # Metadata 2 row
            italy_cols_target = []
            
            for i, country in enumerate(target_countries):
                if str(country).strip() in ['Italy', '#Italy']:
                    series_name = target_mt.iloc[0, i] if i < len(target_mt.iloc[0]) else "Unknown"
                    category = target_mt.iloc[3, i] if target_mt.shape[0] > 3 and i < len(target_mt.iloc[3]) else "Unknown"
                    italy_cols_target.append({
                        'column': i,
                        'name': series_name,
                        'country': country,
                        'category': category
                    })
            
            print(f"   Found {len(italy_cols_target)} Italy series in target:")
            for item in italy_cols_target:
                print(f"     Col {item['column']}: {item['name'][:60]}...")
                print(f"         Country: {item['country']} | Category: {item['category']}")
        
        # Find Italy-related series in OUR MultiTicker  
        print(f"\n🔧 ITALY SERIES IN OUR MULTITICKER:")
        
        if our_mt.shape[0] > 4:
            our_countries = our_mt.iloc[4]  # Metadata 2 row
            italy_cols_ours = []
            
            for i, country in enumerate(our_countries):
                if str(country).strip() in ['Italy', '#Italy']:
                    series_name = our_mt.iloc[0, i] if i < len(our_mt.iloc[0]) else "Unknown"
                    category = our_mt.iloc[3, i] if our_mt.shape[0] > 3 and i < len(our_mt.iloc[3]) else "Unknown"
                    italy_cols_ours.append({
                        'column': i,
                        'name': series_name,
                        'country': country,
                        'category': category
                    })
            
            print(f"   Found {len(italy_cols_ours)} Italy series in ours:")
            for item in italy_cols_ours:
                print(f"     Col {item['column']}: {item['name'][:60]}...")
                print(f"         Country: {item['country']} | Category: {item['category']}")
        
        # Compare the series lists
        print(f"\n🔍 COMPARISON ANALYSIS:")
        print(f"   Target Italy series: {len(italy_cols_target)}")
        print(f"   Our Italy series: {len(italy_cols_ours)}")
        
        # Find differences in series names
        target_names = set(item['name'] for item in italy_cols_target)
        our_names = set(item['name'] for item in italy_cols_ours)
        
        missing_in_ours = target_names - our_names
        extra_in_ours = our_names - target_names
        
        if missing_in_ours:
            print(f"\n❌ MISSING IN OUR DATA:")
            for name in missing_in_ours:
                print(f"     {name[:70]}...")
        
        if extra_in_ours:
            print(f"\n➕ EXTRA IN OUR DATA:")
            for name in extra_in_ours:
                print(f"     {name[:70]}...")
        
        # Focus on DEMAND category specifically
        print(f"\n🎯 DEMAND CATEGORY ANALYSIS:")
        
        target_demand_italy = [item for item in italy_cols_target if item['category'] == 'Demand']
        our_demand_italy = [item for item in italy_cols_ours if item['category'] == 'Demand']
        
        print(f"   Target Italy DEMAND series: {len(target_demand_italy)}")
        for item in target_demand_italy:
            print(f"     {item['name'][:60]}...")
        
        print(f"\n   Our Italy DEMAND series: {len(our_demand_italy)}")
        for item in our_demand_italy:
            print(f"     {item['name'][:60]}...")
        
        # Check specific data values for first few dates
        if italy_cols_target and italy_cols_ours:
            print(f"\n📊 DATA VALUE COMPARISON (first Italy DEMAND series):")
            
            # Get first DEMAND series from each
            target_demand_col = next((item['column'] for item in target_demand_italy), None)
            our_demand_col = next((item['column'] for item in our_demand_italy), None)
            
            if target_demand_col is not None and our_demand_col is not None:
                print(f"   Comparing column {target_demand_col} (target) vs {our_demand_col} (ours)")
                
                # Compare first 10 data rows (skip headers)
                data_start_row = 5  # After headers
                for i in range(data_start_row, min(data_start_row + 10, target_mt.shape[0], our_mt.shape[0])):
                    target_val = target_mt.iloc[i, target_demand_col] if target_demand_col < target_mt.shape[1] else "N/A"
                    our_val = our_mt.iloc[i, our_demand_col] if our_demand_col < our_mt.shape[1] else "N/A"
                    
                    try:
                        target_num = float(target_val) if target_val != "N/A" else 0
                        our_num = float(our_val) if our_val != "N/A" else 0
                        diff = our_num - target_num
                        print(f"     Row {i}: Target={target_num:.2f}, Ours={our_num:.2f}, Diff={diff:.2f}")
                    except:
                        print(f"     Row {i}: Target={target_val}, Ours={our_val}")
        
        return italy_cols_target, italy_cols_ours
        
    except Exception as e:
        print(f"❌ Error in comparison: {e}")
        import traceback
        traceback.print_exc()
        return None, None

if __name__ == "__main__":
    target_italy, our_italy = compare_italy_multiticker()