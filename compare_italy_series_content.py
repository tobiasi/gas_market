# -*- coding: utf-8 -*-
"""
ITALY CONTENT: Compare the actual Italy series between target and our data
"""

import pandas as pd

def compare_italy_series_content():
    """Compare the actual Italy series values between files"""
    
    print(f"\n{'='*80}")
    print("🇮🇹 ITALY SERIES CONTENT COMPARISON")
    print(f"{'='*80}")
    
    try:
        target_mt = pd.read_csv('the_multiticker_we_are_trying_to_replicate.csv', low_memory=False)
        our_mt = pd.read_csv('our_multiticker.csv', low_memory=False)
        
        # Find Italy columns in both files
        print(f"🔍 FINDING ITALY SERIES:")
        
        # Find Italy in target (check row 4 - Metadata 2)
        target_italy_cols = []
        if target_mt.shape[0] > 4:
            for i, country in enumerate(target_mt.iloc[4]):
                if 'Italy' in str(country):
                    series_name = target_mt.iloc[0, i] if i < len(target_mt.iloc[0]) else "Unknown"
                    category = target_mt.iloc[3, i] if i < len(target_mt.iloc[3]) else "Unknown"
                    target_italy_cols.append({
                        'column': i,
                        'name': series_name,
                        'country': country,
                        'category': category
                    })
        
        # Find Italy in our data (check row 2 - we found this earlier)
        our_italy_cols = []
        if our_mt.shape[0] > 2:
            for i, country in enumerate(our_mt.iloc[2]):
                if 'Italy' in str(country):
                    series_name = our_mt.iloc[0, i] if i < len(our_mt.iloc[0]) else "Unknown"
                    category = our_mt.iloc[1, i] if i < len(our_mt.iloc[1]) else "Unknown"
                    our_italy_cols.append({
                        'column': i,
                        'name': series_name,
                        'country': country,
                        'category': category
                    })
        
        print(f"   Target Italy series: {len(target_italy_cols)}")
        print(f"   Our Italy series: {len(our_italy_cols)}")
        
        # Focus on DEMAND category only
        target_demand_italy = [col for col in target_italy_cols if col['category'] == 'Demand']
        our_demand_italy = [col for col in our_italy_cols if col['category'] == 'Demand']
        
        print(f"\n🎯 ITALY DEMAND SERIES:")
        print(f"   Target Italy DEMAND: {len(target_demand_italy)}")
        print(f"   Our Italy DEMAND: {len(our_demand_italy)}")
        
        # Compare the actual data values for Italy DEMAND series
        print(f"\n📊 ITALY DEMAND VALUES COMPARISON:")
        
        # Sum all Italy DEMAND series for first 10 dates
        data_start = 5  # Skip headers
        
        print(f"{'Date':<12} {'Target Sum':<12} {'Our Sum':<12} {'Difference':<12}")
        print("-" * 50)
        
        for row in range(data_start, min(data_start + 10, target_mt.shape[0], our_mt.shape[0])):
            # Sum target Italy DEMAND values
            target_sum = 0
            for col_info in target_demand_italy:
                try:
                    val = float(target_mt.iloc[row, col_info['column']])
                    target_sum += val
                except (ValueError, TypeError):
                    pass
            
            # Sum our Italy DEMAND values  
            our_sum = 0
            for col_info in our_demand_italy:
                try:
                    val = float(our_mt.iloc[row, col_info['column']])
                    our_sum += val
                except (ValueError, TypeError):
                    pass
            
            difference = our_sum - target_sum
            date_str = str(our_mt.iloc[row, 0])[:10] if row < our_mt.shape[0] else "Unknown"
            
            print(f"{date_str:<12} {target_sum:<12.2f} {our_sum:<12.2f} {difference:<12.2f}")
        
        # Analyze individual Italy DEMAND series
        print(f"\n🔍 INDIVIDUAL ITALY DEMAND SERIES ANALYSIS:")
        
        print(f"\n📋 TARGET ITALY DEMAND SERIES:")
        for i, col_info in enumerate(target_demand_italy):
            sample_val = 0
            try:
                sample_val = float(target_mt.iloc[data_start, col_info['column']])
            except (ValueError, TypeError):
                pass
            print(f"   {i+1}. Col {col_info['column']}: {col_info['name']} = {sample_val:.2f}")
        
        print(f"\n📋 OUR ITALY DEMAND SERIES:")
        for i, col_info in enumerate(our_demand_italy):
            sample_val = 0
            try:
                sample_val = float(our_mt.iloc[data_start, col_info['column']])
            except (ValueError, TypeError):
                pass
            print(f"   {i+1}. Col {col_info['column']}: {col_info['name']} = {sample_val:.2f}")
        
        return target_demand_italy, our_demand_italy
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return None, None

if __name__ == "__main__":
    target_italy, our_italy = compare_italy_series_content()