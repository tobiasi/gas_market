# Validate monthly calculations manually
import pandas as pd

def validate_monthly_calculations():
    print("🔍 MONTHLY CALCULATIONS VALIDATION")
    print("=" * 60)
    
    # Load master file
    master_file = 'European_Gas_Market_Master.xlsx'
    
    # Load daily and monthly data
    demand_daily = pd.read_excel(master_file, sheet_name='Demand')
    demand_monthly = pd.read_excel(master_file, sheet_name='Demand Monthly')
    demand_yoy = pd.read_excel(master_file, sheet_name='Demand Monthly YOY')
    
    demand_daily['Date'] = pd.to_datetime(demand_daily['Date'])
    demand_monthly['Date'] = pd.to_datetime(demand_monthly['Date'])
    
    print("✅ Data loaded successfully")
    print(f"📊 Daily data: {demand_daily.shape}")
    print(f"📅 Monthly data: {demand_monthly.shape}")
    print(f"📈 YOY data: {demand_yoy.shape}")
    
    # Manual verification: October 2016 Italy average
    print(f"\n📊 MANUAL VERIFICATION - October 2016 Italy")
    print("=" * 50)
    
    oct_2016_data = demand_daily[
        (demand_daily['Date'].dt.year == 2016) & 
        (demand_daily['Date'].dt.month == 10)
    ]
    
    manual_italy_avg = oct_2016_data['Italy'].mean()
    monthly_italy_oct2016 = demand_monthly[
        (demand_monthly['Date'].dt.year == 2016) & 
        (demand_monthly['Date'].dt.month == 10)
    ]['Italy'].iloc[0]
    
    print(f"Manual calculation: {manual_italy_avg:.2f}")
    print(f"Monthly aggregation: {monthly_italy_oct2016:.2f}")
    print(f"Match: {'✅' if abs(manual_italy_avg - monthly_italy_oct2016) < 0.01 else '❌'}")
    
    # Manual YOY verification
    print(f"\n📈 MANUAL YOY VERIFICATION - October 2017 vs 2016")
    print("=" * 50)
    
    oct_2017_italy = demand_monthly[
        (demand_monthly['Date'].dt.year == 2017) & 
        (demand_monthly['Date'].dt.month == 10)
    ]['Italy'].iloc[0]
    
    oct_2016_italy = demand_monthly[
        (demand_monthly['Date'].dt.year == 2016) & 
        (demand_monthly['Date'].dt.month == 10)
    ]['Italy'].iloc[0]
    
    manual_yoy_abs = oct_2017_italy - oct_2016_italy
    manual_yoy_pct = ((oct_2017_italy / oct_2016_italy) - 1) * 100
    
    # Find in YOY data
    yoy_oct_2017 = demand_yoy[demand_yoy['Year-Month'] == '2017-10']
    if not yoy_oct_2017.empty:
        system_yoy_abs = yoy_oct_2017['Italy_YOY_Abs'].iloc[0]
        system_yoy_pct = yoy_oct_2017['Italy_YOY_Pct'].iloc[0]
        
        print(f"2016 Oct Italy: {oct_2016_italy:.2f}")
        print(f"2017 Oct Italy: {oct_2017_italy:.2f}")
        print(f"Manual YOY Absolute: {manual_yoy_abs:.2f}")
        print(f"System YOY Absolute: {system_yoy_abs:.2f}")
        print(f"Manual YOY Percent: {manual_yoy_pct:.2f}%")
        print(f"System YOY Percent: {system_yoy_pct:.2f}%")
        print(f"Absolute Match: {'✅' if abs(manual_yoy_abs - system_yoy_abs) < 0.01 else '❌'}")
        print(f"Percent Match: {'✅' if abs(manual_yoy_pct - system_yoy_pct) < 0.01 else '❌'}")
    
    # Summary statistics validation
    print(f"\n📊 SUMMARY STATISTICS VALIDATION")
    print("=" * 50)
    
    # Check that monthly data makes sense
    monthly_ranges = {}
    for col in ['Italy', 'Total', 'Norway', 'LNG']:
        if col in demand_monthly.columns:
            values = demand_monthly[col].dropna()
            monthly_ranges[col] = {
                'min': values.min(),
                'max': values.max(),
                'mean': values.mean(),
                'std': values.std()
            }
            print(f"{col}: min={values.min():.1f}, max={values.max():.1f}, mean={values.mean():.1f}")
    
    print(f"\n✅ MONTHLY CALCULATIONS VALIDATION COMPLETE")
    
    return True

if __name__ == "__main__":
    validate_monthly_calculations()