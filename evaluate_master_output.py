# Comprehensive evaluation of master output
import pandas as pd
import numpy as np

def evaluate_master_output():
    print("📊 COMPREHENSIVE MASTER OUTPUT EVALUATION")
    print("=" * 80)
    
    # Load the master file
    master_file = 'European_Gas_Market_Master.xlsx'
    demand_df = pd.read_excel(master_file, sheet_name='Demand')
    supply_df = pd.read_excel(master_file, sheet_name='Supply')
    
    print(f"✅ Master file loaded: {master_file}")
    print(f"📊 Demand shape: {demand_df.shape}")
    print(f"⛽ Supply shape: {supply_df.shape}")
    
    # Convert dates
    demand_df['Date'] = pd.to_datetime(demand_df['Date'])
    supply_df['Date'] = pd.to_datetime(supply_df['Date'])
    
    print(f"📅 Date range: {demand_df['Date'].min()} to {demand_df['Date'].max()}")
    print(f"📅 Total days: {len(demand_df)} days")
    
    # Check data completeness
    print(f"\n📈 DATA COMPLETENESS ANALYSIS")
    print("=" * 50)
    
    demand_metrics = ['Italy', 'Netherlands', 'GB', 'Austria', 'Total', 'Industrial', 'LDZ', 'Gas-to-Power']
    supply_metrics = ['Norway', 'Russia (Nord Stream)', 'LNG', 'Algeria', 'Total']
    
    print("DEMAND DATA COMPLETENESS:")
    for metric in demand_metrics:
        if metric in demand_df.columns:
            non_null = demand_df[metric].count()
            null_count = demand_df[metric].isnull().sum()
            completeness = (non_null / len(demand_df)) * 100
            print(f"  {metric:<15}: {completeness:6.1f}% ({non_null}/{len(demand_df)} values)")
    
    print("\nSUPPLY DATA COMPLETENESS:")
    for metric in supply_metrics:
        if metric in supply_df.columns:
            non_null = supply_df[metric].count()
            null_count = supply_df[metric].isnull().sum()
            completeness = (non_null / len(supply_df)) * 100
            print(f"  {metric:<20}: {completeness:6.1f}% ({non_null}/{len(supply_df)} values)")
    
    # Statistical summary
    print(f"\n📊 KEY STATISTICAL SUMMARY")
    print("=" * 50)
    
    print("DEMAND STATISTICS (Recent 30 days):")
    recent_demand = demand_df.tail(30)
    for metric in ['Italy', 'Total', 'Industrial']:
        if metric in recent_demand.columns:
            values = recent_demand[metric].dropna()
            if len(values) > 0:
                print(f"  {metric:<12}: avg={values.mean():7.1f}, min={values.min():7.1f}, max={values.max():7.1f}")
    
    print("\nSUPPLY STATISTICS (Recent 30 days):")
    recent_supply = supply_df.tail(30)
    for metric in ['Norway', 'LNG', 'Total']:
        if metric in recent_supply.columns:
            values = recent_supply[metric].dropna()
            if len(values) > 0:
                print(f"  {metric:<12}: avg={values.mean():7.1f}, min={values.min():7.1f}, max={values.max():7.1f}")
    
    # Cross-validation check
    print(f"\n🔍 CROSS-VALIDATION CHECK")
    print("=" * 50)
    
    # Check specific date values
    test_date = '2016-10-04'
    test_row_demand = demand_df[demand_df['Date'] == test_date]
    test_row_supply = supply_df[supply_df['Date'] == test_date]
    
    if not test_row_demand.empty:
        print(f"DEMAND VALUES for {test_date}:")
        row = test_row_demand.iloc[0]
        key_metrics = ['Italy', 'Netherlands', 'GB', 'Total']
        for metric in key_metrics:
            if metric in row.index:
                print(f"  {metric}: {row[metric]:.3f}")
    
    if not test_row_supply.empty:
        print(f"\nSUPPLY VALUES for {test_date}:")
        row = test_row_supply.iloc[0]
        key_metrics = ['Norway', 'Russia (Nord Stream)', 'LNG', 'Total']
        for metric in key_metrics:
            if metric in row.index:
                print(f"  {metric}: {row[metric]:.3f}")
    
    return demand_df, supply_df

if __name__ == "__main__":
    demand_df, supply_df = evaluate_master_output()