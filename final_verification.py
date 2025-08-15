# Final verification to check which countries we might be missing
import pandas as pd
import numpy as np

def final_verification():
    """Final check to ensure we have all demand countries"""
    
    print("🔎 FINAL VERIFICATION - COMPLETE COUNTRY ANALYSIS")
    print("=" * 70)
    
    try:
        # Load the corrected master file
        master_file = 'European_Gas_Market_Master.xlsx'
        demand_df = pd.read_excel(master_file, sheet_name='Demand')
        demand_df['Date'] = pd.to_datetime(demand_df['Date'])
        
        print(f"✅ Loaded corrected demand data: {demand_df.shape}")
        
        # Get all country columns (excluding categories and totals)
        all_countries = [col for col in demand_df.columns 
                        if col not in ['Date', 'Total', 'Industrial', 'LDZ', 'Gas-to-Power']]
        
        print(f"\n📊 ALL COUNTRY COLUMNS FOUND:")
        for i, country in enumerate(all_countries, 1):
            print(f"  {i:2d}. {country}")
        
        # Check specific date: 2016-10-04
        target_date = '2016-10-04'
        target_row = demand_df[demand_df['Date'].dt.strftime('%Y-%m-%d') == target_date].iloc[0]
        
        print(f"\n📅 VALUES FOR {target_date}:")
        countries_sum = 0
        for country in all_countries:
            val = target_row[country]
            countries_sum += val if pd.notna(val) else 0
            print(f"  {country:15}: {val:8.3f}")
        
        total_val = target_row['Total']
        categories_sum = target_row['Industrial'] + target_row['LDZ'] + target_row['Gas-to-Power']
        
        print(f"\n📊 SUMMARY:")
        print(f"  All Countries Sum: {countries_sum:8.3f}")
        print(f"  Categories Sum:    {categories_sum:8.3f}")
        print(f"  Total Column:      {total_val:8.3f}")
        print(f"  Countries vs Total diff: {abs(total_val - countries_sum):8.3f}")
        print(f"  Categories vs Total diff: {abs(total_val - categories_sum):8.3f}")
        
        if abs(total_val - countries_sum) < 0.001:
            print(f"  ✅ Countries perfectly sum to Total!")
        elif abs(total_val - categories_sum) < 0.001:
            print(f"  ✅ Categories perfectly sum to Total!")
            print(f"  💡 This confirms: Total = Industrial + LDZ + Gas-to-Power")
            print(f"     Countries provide geographical breakdown")
            print(f"     Categories provide consumption type breakdown")
        
        # Statistical check across all data
        print(f"\n📈 STATISTICAL ANALYSIS (all {len(demand_df)} rows):")
        
        # Calculate sums for all rows
        countries_sum_series = demand_df[all_countries].sum(axis=1)
        categories_sum_series = demand_df[['Industrial', 'LDZ', 'Gas-to-Power']].sum(axis=1)
        total_series = demand_df['Total']
        
        countries_diff = abs(total_series - countries_sum_series)
        categories_diff = abs(total_series - categories_sum_series)
        
        print(f"  Countries vs Total:")
        print(f"    Mean diff: {countries_diff.mean():.3f}")
        print(f"    Max diff:  {countries_diff.max():.3f}")
        print(f"    Perfect matches (< 0.001): {(countries_diff < 0.001).sum()}/{len(demand_df)}")
        
        print(f"  Categories vs Total:")
        print(f"    Mean diff: {categories_diff.mean():.3f}")
        print(f"    Max diff:  {categories_diff.max():.3f}")
        print(f"    Perfect matches (< 0.001): {(categories_diff < 0.001).sum()}/{len(demand_df)}")
        
        # Final conclusion
        if (categories_diff < 0.001).sum() > (countries_diff < 0.001).sum():
            print(f"\n💡 FINAL CONCLUSION:")
            print(f"   ✅ Total column represents CONSUMPTION CATEGORIES (Industrial + LDZ + Gas-to-Power)")
            print(f"   📍 Country columns provide geographical distribution of that same total")
            print(f"   🎯 Both are valid views of European gas demand - just different breakdowns!")
            print(f"   📋 Your original observation was correct: in our system they don't sum")
            print(f"      because Total = Categories, not Countries")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    final_verification()