# -*- coding: utf-8 -*-
"""
IDENTIFY: Find the 22 extra series we have that Excel doesn't
"""

import pandas as pd

def identify_extra_series():
    """Identify which 22 series we have that Excel doesn't"""
    
    print(f"\n{'='*80}")
    print("🔍 IDENTIFYING EXTRA SERIES")
    print(f"{'='*80}")
    
    try:
        target_mt = pd.read_csv('the_multiticker_we_are_trying_to_replicate.csv', low_memory=False)
        our_mt = pd.read_csv('our_multiticker.csv', low_memory=False)
        
        print(f"📊 Loaded files:")
        print(f"   Target: {target_mt.shape[1]} series")
        print(f"   Ours: {our_mt.shape[1]} series")
        print(f"   Extra: {our_mt.shape[1] - target_mt.shape[1]} series")
        
        # Get series names from first row (should be ticker names)
        target_series = target_mt.iloc[0].fillna('').astype(str).tolist()
        our_series = our_mt.iloc[0].fillna('').astype(str).tolist()
        
        # Clean up series names (remove empty/nan)
        target_series_clean = [s.strip() for s in target_series if s.strip() and s.strip().lower() != 'nan']
        our_series_clean = [s.strip() for s in our_series if s.strip() and s.strip().lower() != 'nan']
        
        print(f"\n📋 Clean series names:")
        print(f"   Target clean: {len(target_series_clean)} series")
        print(f"   Ours clean: {len(our_series_clean)} series")
        
        # Find differences
        target_set = set(target_series_clean)
        our_set = set(our_series_clean)
        
        extra_in_ours = our_set - target_set
        missing_in_ours = target_set - our_set
        
        print(f"\n🔍 SERIES DIFFERENCES:")
        print(f"   Extra in our data: {len(extra_in_ours)}")
        print(f"   Missing in our data: {len(missing_in_ours)}")
        
        if extra_in_ours:
            print(f"\n➕ EXTRA SERIES IN OUR DATA:")
            for i, series in enumerate(sorted(extra_in_ours), 1):
                print(f"   {i:2d}. {series}")
                
                # Find which column this series is in our data
                try:
                    col_idx = our_series.index(series)
                    
                    # Get metadata for this series
                    category = our_mt.iloc[1, col_idx] if our_mt.shape[0] > 1 else "Unknown"
                    country = our_mt.iloc[2, col_idx] if our_mt.shape[0] > 2 else "Unknown"
                    sector = our_mt.iloc[3, col_idx] if our_mt.shape[0] > 3 else "Unknown"
                    
                    print(f"       Category: {category} | Country: {country} | Sector: {sector}")
                    
                    # Check if this is Italy-related
                    if 'Italy' in str(country) or 'italy' in str(series).lower():
                        print(f"       🇮🇹 ITALY-RELATED! This could explain the +3 difference")
                    
                except (ValueError, IndexError):
                    print(f"       Could not find metadata")
        
        if missing_in_ours:
            print(f"\n❌ MISSING SERIES IN OUR DATA:")
            for i, series in enumerate(sorted(missing_in_ours), 1):
                print(f"   {i:2d}. {series}")
                
                # Find metadata in target data
                try:
                    col_idx = target_series.index(series)
                    category = target_mt.iloc[3, col_idx] if target_mt.shape[0] > 3 else "Unknown"  # Row 3 = Metadata 1
                    country = target_mt.iloc[4, col_idx] if target_mt.shape[0] > 4 else "Unknown"  # Row 4 = Metadata 2
                    
                    print(f"       Category: {category} | Country: {country}")
                    
                    if 'Italy' in str(country) or 'italy' in str(series).lower():
                        print(f"       🇮🇹 MISSING ITALY SERIES! This could explain the -3 difference")
                        
                except (ValueError, IndexError):
                    print(f"       Could not find metadata")
        
        # Focus on Italy-related differences
        print(f"\n🇮🇹 ITALY-SPECIFIC ANALYSIS:")
        
        italy_extra = [s for s in extra_in_ours if 'italy' in s.lower() or 'IT' in s]
        italy_missing = [s for s in missing_in_ours if 'italy' in s.lower() or 'IT' in s]
        
        print(f"   Extra Italy-related series: {len(italy_extra)}")
        for series in italy_extra:
            print(f"     + {series}")
            
        print(f"   Missing Italy-related series: {len(italy_missing)}")
        for series in italy_missing:
            print(f"     - {series}")
        
        # Calculate net effect
        net_italy_difference = len(italy_extra) - len(italy_missing)
        print(f"\n📊 NET ITALY SERIES DIFFERENCE: {net_italy_difference:+d}")
        
        if net_italy_difference > 0:
            print(f"   → We have MORE Italy series than Excel (+{net_italy_difference})")
            print(f"   → This likely explains why our Italy values are +3 units higher")
        elif net_italy_difference < 0:
            print(f"   → We have FEWER Italy series than Excel ({net_italy_difference})")
            print(f"   → This would explain lower values, not higher")
        else:
            print(f"   → Same number of Italy series, but different ones")
            print(f"   → The +3 difference comes from different Italy series content")
        
        return extra_in_ours, missing_in_ours
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return None, None

if __name__ == "__main__":
    extra, missing = identify_extra_series()