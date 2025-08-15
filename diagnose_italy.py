# -*- coding: utf-8 -*-
"""
DIAGNOSTIC: Check Italy demand series specifically
"""

def diagnose_italy_demand(full_data):
    """Diagnose what Italy demand series are being captured"""
    
    print(f"\n{'='*60}")
    print("🇮🇹 ITALY DEMAND DIAGNOSTIC")
    print(f"{'='*60}")
    
    # Find ALL series with Italy in the country field
    italy_mask = full_data.columns.get_level_values(3) == 'Italy'
    italy_series = full_data.columns[italy_mask]
    
    print(f"📊 ALL ITALY SERIES IN DATA:")
    print(f"   Total Italy series found: {len(italy_series)}")
    
    for i, col in enumerate(italy_series):
        desc = col[0][:60] + "..." if len(col[0]) > 60 else col[0]
        category = col[2] if len(col) > 2 else "Unknown"
        region_from = col[3] if len(col) > 3 else "Unknown"
        sector = col[4] if len(col) > 4 else "Unknown"
        print(f"   {i+1:2d}. {desc}")
        print(f"       Category: {category} | Region: {region_from} | Sector: {sector}")
    
    # Check specifically DEMAND category Italy series
    italy_demand_mask = ((full_data.columns.get_level_values(2) == 'Demand') & 
                        (full_data.columns.get_level_values(3) == 'Italy'))
    italy_demand_series = full_data.columns[italy_demand_mask]
    
    print(f"\n🎯 ITALY 'DEMAND' SERIES (what gets summed):")
    print(f"   Italy Demand series found: {len(italy_demand_series)}")
    
    for i, col in enumerate(italy_demand_series):
        desc = col[0][:60] + "..." if len(col[0]) > 60 else col[0]
        sector = col[4] if len(col) > 4 else "Unknown"
        sample_value = full_data[col].iloc[0] if not full_data[col].empty else 0
        print(f"   {i+1:2d}. {desc}")
        print(f"       Sector: {sector} | Sample value: {sample_value:.2f}")
    
    # Check for other categories that might be missing
    print(f"\n🔍 OTHER ITALY CATEGORIES (not included in sum):")
    other_italy_mask = ((full_data.columns.get_level_values(3) == 'Italy') & 
                       (full_data.columns.get_level_values(2) != 'Demand'))
    other_italy_series = full_data.columns[other_italy_mask]
    
    for i, col in enumerate(other_italy_series):
        desc = col[0][:50] + "..." if len(col[0]) > 50 else col[0]
        category = col[2] if len(col) > 2 else "Unknown"
        sample_value = full_data[col].iloc[0] if not full_data[col].empty else 0
        print(f"   {i+1:2d}. {desc}")
        print(f"       Category: {category} | Sample value: {sample_value:.2f}")
    
    # Calculate the current Italy total
    if len(italy_demand_series) > 0:
        italy_total = full_data.iloc[:, italy_demand_mask].sum(axis=1, skipna=False)
        sample_total = italy_total.iloc[0]
        print(f"\n📈 CURRENT ITALY TOTAL:")
        print(f"   Sample total (first date): {sample_total:.2f}")
        print(f"   Based on {len(italy_demand_series)} demand series")
    else:
        print(f"\n❌ NO ITALY DEMAND SERIES FOUND!")
    
    # Check for potential missing series patterns
    print(f"\n💡 POTENTIAL ISSUES TO CHECK:")
    print(f"   1. Are there Italy series with different category names?")
    print(f"   2. Are there Italy series with slightly different country names?")
    print(f"   3. Are there missing Italy series in the data entirely?")
    print(f"   4. Check Excel vs Python calculation for same date")
    
    return italy_demand_series

# Add this to your main script after full_data is created:
print("\n" + "="*60)
print("🇮🇹 To diagnose Italy, add this to your script:")
print("italy_series = diagnose_italy_demand(full_data)")
print("="*60)