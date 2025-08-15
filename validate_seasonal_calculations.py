# Comprehensive validation of seasonal calculations
import pandas as pd
import numpy as np

def validate_seasonal_calculations():
    print("🔍 SEASONAL CALCULATIONS VALIDATION")
    print("=" * 80)
    
    # Load seasonal file
    seasonal_file = 'European_Gas_Seasonal_Analysis.xlsx'
    
    # Load a YOY percentage sheet
    italy_yoy_pct = pd.read_excel(seasonal_file, sheet_name='D_YOY_Pct_Italy', index_col=0)
    italy_yoy_monthly_pct = pd.read_excel(seasonal_file, sheet_name='M_YOY_Pct_Italy', index_col=0)
    
    print(f"✅ Loaded seasonal data")
    print(f"📊 Daily YOY Italy shape: {italy_yoy_pct.shape}")
    print(f"📅 Monthly YOY Italy shape: {italy_yoy_monthly_pct.shape}")
    
    # Load master data for validation
    master_file = 'European_Gas_Market_Master.xlsx'
    demand_daily = pd.read_excel(master_file, sheet_name='Demand')
    demand_daily['Date'] = pd.to_datetime(demand_daily['Date'])
    
    print(f"\n📊 VALIDATION: Daily YOY Percentage Calculations")
    print("=" * 60)
    
    # Manual calculation for specific day: October 4th (day 277)
    day_of_year = 277  # October 4th
    
    # Get data for Oct 4th, 2017 and 2018
    oct4_2017 = demand_daily[
        (demand_daily['Date'].dt.dayofyear == day_of_year) & 
        (demand_daily['Date'].dt.year == 2017)
    ]['Italy'].iloc[0]
    
    oct4_2018 = demand_daily[
        (demand_daily['Date'].dt.dayofyear == day_of_year) & 
        (demand_daily['Date'].dt.year == 2018)
    ]['Italy'].iloc[0]
    
    manual_yoy_pct = ((oct4_2018 / oct4_2017) - 1) * 100
    seasonal_yoy_pct = italy_yoy_pct.loc[day_of_year, 2018]
    
    print(f"Day {day_of_year} (Oct 4th):")
    print(f"  2017 Italy: {oct4_2017:.2f}")
    print(f"  2018 Italy: {oct4_2018:.2f}")
    print(f"  Manual YOY %: {manual_yoy_pct:.2f}%")
    print(f"  Seasonal YOY %: {seasonal_yoy_pct:.2f}%")
    print(f"  Match: {'✅' if abs(manual_yoy_pct - seasonal_yoy_pct) < 0.01 else '❌'}")
    
    print(f"\n📊 VALIDATION: Baseline Statistics (2017-2021 Average)")
    print("=" * 60)
    
    # Get 2017-2021 values for day 277
    baseline_years = [2017, 2018, 2019, 2020, 2021]
    day277_values = []
    
    for year in baseline_years:
        try:
            value = italy_yoy_pct.loc[day_of_year, year]
            if pd.notna(value):
                day277_values.append(value)
        except:
            continue
    
    if day277_values:
        manual_avg = np.mean(day277_values)
        manual_min = np.min(day277_values)
        manual_max = np.max(day277_values)
        manual_range = manual_max - manual_min
        
        seasonal_avg = italy_yoy_pct.loc[day_of_year, 'Avg_2018-2021']
        seasonal_min = italy_yoy_pct.loc[day_of_year, 'Min_2018-2021']
        seasonal_max = italy_yoy_pct.loc[day_of_year, 'Max_2018-2021']
        seasonal_range = italy_yoy_pct.loc[day_of_year, 'Range_2018-2021']
        
        print(f"Baseline values for Day {day_of_year}: {day277_values}")
        print(f"Manual Average: {manual_avg:.2f}%, Seasonal: {seasonal_avg:.2f}% {'✅' if abs(manual_avg - seasonal_avg) < 0.01 else '❌'}")
        print(f"Manual Min: {manual_min:.2f}%, Seasonal: {seasonal_min:.2f}% {'✅' if abs(manual_min - seasonal_min) < 0.01 else '❌'}")
        print(f"Manual Max: {manual_max:.2f}%, Seasonal: {seasonal_max:.2f}% {'✅' if abs(manual_max - seasonal_max) < 0.01 else '❌'}")
        print(f"Manual Range: {manual_range:.2f}%, Seasonal: {seasonal_range:.2f}% {'✅' if abs(manual_range - seasonal_range) < 0.01 else '❌'}")
    
    print(f"\n📊 VALIDATION: Monthly YOY Calculations")
    print("=" * 60)
    
    # Validate monthly calculations
    demand_monthly = pd.read_excel(master_file, sheet_name='Demand Monthly')
    demand_monthly['Date'] = pd.to_datetime(demand_monthly['Date'])
    
    # October 2017 vs 2018
    oct_2017 = demand_monthly[
        (demand_monthly['Date'].dt.year == 2017) & 
        (demand_monthly['Date'].dt.month == 10)
    ]['Italy'].iloc[0]
    
    oct_2018 = demand_monthly[
        (demand_monthly['Date'].dt.year == 2018) & 
        (demand_monthly['Date'].dt.month == 10)
    ]['Italy'].iloc[0]
    
    manual_monthly_yoy = ((oct_2018 / oct_2017) - 1) * 100
    seasonal_monthly_yoy = italy_yoy_monthly_pct.loc['Oct', 2018]
    
    print(f"October Monthly YOY:")
    print(f"  2017: {oct_2017:.2f}")
    print(f"  2018: {oct_2018:.2f}")
    print(f"  Manual: {manual_monthly_yoy:.2f}%")
    print(f"  Seasonal: {seasonal_monthly_yoy:.2f}%")
    print(f"  Match: {'✅' if abs(manual_monthly_yoy - seasonal_monthly_yoy) < 0.01 else '❌'}")
    
    print(f"\n📊 DATA QUALITY SUMMARY")
    print("=" * 60)
    
    # Check data ranges and quality
    print("Daily YOY Percentage Data Quality:")
    daily_data = italy_yoy_pct[[2018, 2019, 2020, 2021]].values.flatten()
    daily_data_clean = daily_data[~np.isnan(daily_data)]
    print(f"  Valid values: {len(daily_data_clean)}/{len(daily_data)} ({len(daily_data_clean)/len(daily_data)*100:.1f}%)")
    print(f"  Range: {daily_data_clean.min():.1f}% to {daily_data_clean.max():.1f}%")
    print(f"  Mean: {daily_data_clean.mean():.1f}%")
    
    print("\nMonthly YOY Percentage Data Quality:")
    monthly_data = italy_yoy_monthly_pct[[2018, 2019, 2020, 2021]].values.flatten()
    monthly_data_clean = monthly_data[~np.isnan(monthly_data)]
    print(f"  Valid values: {len(monthly_data_clean)}/{len(monthly_data)} ({len(monthly_data_clean)/len(monthly_data)*100:.1f}%)")
    print(f"  Range: {monthly_data_clean.min():.1f}% to {monthly_data_clean.max():.1f}%")
    print(f"  Mean: {monthly_data_clean.mean():.1f}%")
    
    print(f"\n✅ SEASONAL CALCULATIONS VALIDATION COMPLETE")
    
    return True

if __name__ == "__main__":
    validate_seasonal_calculations()