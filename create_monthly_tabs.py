# -*- coding: utf-8 -*-
"""
Create monthly aggregation tabs for demand and supply data
- Demand Monthly: Monthly averages for demand
- Supply Monthly: Monthly averages for supply  
- Demand Monthly YOY: Year-over-year comparison for demand
- Supply Monthly YOY: Year-over-year comparison for supply
"""

import pandas as pd
import numpy as np

def load_existing_data():
    """Load existing demand and supply data"""
    
    print("📊 LOADING EXISTING DEMAND AND SUPPLY DATA")
    print("=" * 60)
    
    master_file = 'European_Gas_Market_Master.xlsx'
    
    # Load demand and supply data
    demand_df = pd.read_excel(master_file, sheet_name='Demand')
    supply_df = pd.read_excel(master_file, sheet_name='Supply')
    
    # Convert Date columns to datetime
    demand_df['Date'] = pd.to_datetime(demand_df['Date'])
    supply_df['Date'] = pd.to_datetime(supply_df['Date'])
    
    # Set Date as index for easier resampling
    demand_df.set_index('Date', inplace=True)
    supply_df.set_index('Date', inplace=True)
    
    print(f"✅ Demand data: {demand_df.shape}")
    print(f"✅ Supply data: {supply_df.shape}")
    print(f"📅 Date range: {demand_df.index.min()} to {demand_df.index.max()}")
    
    return demand_df, supply_df

def create_demand_monthly(demand_df):
    """Create monthly averages for demand data"""
    
    print("\n🏠 CREATING DEMAND MONTHLY AVERAGES")
    print("=" * 60)
    
    # Resample to monthly averages
    demand_monthly = demand_df.resample('M').mean()
    
    # Round to 2 decimal places
    demand_monthly = demand_monthly.round(2)
    
    # Create Year-Month column for better readability
    demand_monthly['Year'] = demand_monthly.index.year
    demand_monthly['Month'] = demand_monthly.index.month
    demand_monthly['Year-Month'] = demand_monthly.index.strftime('%Y-%m')
    
    # Reset index to make Date a column again
    demand_monthly.reset_index(inplace=True)
    
    # Reorder columns: Date, Year, Month, Year-Month, then data columns
    date_cols = ['Date', 'Year', 'Month', 'Year-Month']
    data_cols = [col for col in demand_monthly.columns if col not in date_cols]
    demand_monthly = demand_monthly[date_cols + data_cols]
    
    print(f"✅ Created monthly demand averages: {demand_monthly.shape}")
    print(f"📅 Monthly periods: {len(demand_monthly)} months")
    
    # Show sample
    print(f"\n📊 SAMPLE DEMAND MONTHLY DATA:")
    print(demand_monthly.head()[['Year-Month', 'Italy', 'Total', 'Industrial', 'LDZ', 'Gas-to-Power']].to_string())
    
    return demand_monthly

def create_supply_monthly(supply_df):
    """Create monthly averages for supply data"""
    
    print("\n⛽ CREATING SUPPLY MONTHLY AVERAGES")
    print("=" * 60)
    
    # Resample to monthly averages
    supply_monthly = supply_df.resample('M').mean()
    
    # Round to 2 decimal places
    supply_monthly = supply_monthly.round(2)
    
    # Create Year-Month column for better readability
    supply_monthly['Year'] = supply_monthly.index.year
    supply_monthly['Month'] = supply_monthly.index.month
    supply_monthly['Year-Month'] = supply_monthly.index.strftime('%Y-%m')
    
    # Reset index to make Date a column again
    supply_monthly.reset_index(inplace=True)
    
    # Reorder columns: Date, Year, Month, Year-Month, then data columns
    date_cols = ['Date', 'Year', 'Month', 'Year-Month']
    data_cols = [col for col in supply_monthly.columns if col not in date_cols]
    supply_monthly = supply_monthly[date_cols + data_cols]
    
    print(f"✅ Created monthly supply averages: {supply_monthly.shape}")
    print(f"📅 Monthly periods: {len(supply_monthly)} months")
    
    # Show sample
    key_supply_cols = ['Year-Month', 'Norway', 'Russia (Nord Stream)', 'LNG', 'Total']
    available_cols = [col for col in key_supply_cols if col in supply_monthly.columns]
    print(f"\n📊 SAMPLE SUPPLY MONTHLY DATA:")
    print(supply_monthly.head()[available_cols].to_string())
    
    return supply_monthly

def create_demand_monthly_yoy(demand_monthly):
    """Create year-over-year comparison for demand monthly data"""
    
    print("\n📈 CREATING DEMAND MONTHLY YOY")
    print("=" * 60)
    
    # Create a copy for YOY calculations
    demand_yoy = demand_monthly.copy()
    
    # Get data columns (exclude date/time columns)
    date_cols = ['Date', 'Year', 'Month', 'Year-Month']
    data_cols = [col for col in demand_yoy.columns if col not in date_cols]
    
    # For each data column, calculate YOY change
    for col in data_cols:
        # Shift by 12 months (12 rows) to get same month previous year
        demand_yoy[f'{col}_YOY_Abs'] = demand_yoy[col] - demand_yoy[col].shift(12)
        demand_yoy[f'{col}_YOY_Pct'] = ((demand_yoy[col] / demand_yoy[col].shift(12)) - 1) * 100
    
    # Round YOY columns
    yoy_cols = [col for col in demand_yoy.columns if '_YOY_' in col]
    demand_yoy[yoy_cols] = demand_yoy[yoy_cols].round(2)
    
    # Remove rows where YOY calculation is not possible (first 12 months)
    demand_yoy = demand_yoy.dropna(subset=[f'{data_cols[0]}_YOY_Abs'])
    
    print(f"✅ Created demand YOY data: {demand_yoy.shape}")
    print(f"📅 YOY periods: {len(demand_yoy)} months (starting from 2nd year)")
    
    # Show sample YOY data
    print(f"\n📊 SAMPLE DEMAND YOY DATA:")
    sample_cols = ['Year-Month', 'Italy', 'Italy_YOY_Abs', 'Italy_YOY_Pct', 'Total', 'Total_YOY_Pct']
    available_cols = [col for col in sample_cols if col in demand_yoy.columns]
    print(demand_yoy.head()[available_cols].to_string())
    
    return demand_yoy

def create_supply_monthly_yoy(supply_monthly):
    """Create year-over-year comparison for supply monthly data"""
    
    print("\n📈 CREATING SUPPLY MONTHLY YOY")
    print("=" * 60)
    
    # Create a copy for YOY calculations
    supply_yoy = supply_monthly.copy()
    
    # Get data columns (exclude date/time columns)
    date_cols = ['Date', 'Year', 'Month', 'Year-Month']
    data_cols = [col for col in supply_yoy.columns if col not in date_cols]
    
    # For each data column, calculate YOY change
    for col in data_cols:
        # Shift by 12 months (12 rows) to get same month previous year
        supply_yoy[f'{col}_YOY_Abs'] = supply_yoy[col] - supply_yoy[col].shift(12)
        supply_yoy[f'{col}_YOY_Pct'] = ((supply_yoy[col] / supply_yoy[col].shift(12)) - 1) * 100
    
    # Round YOY columns
    yoy_cols = [col for col in supply_yoy.columns if '_YOY_' in col]
    supply_yoy[yoy_cols] = supply_yoy[yoy_cols].round(2)
    
    # Remove rows where YOY calculation is not possible (first 12 months)
    supply_yoy = supply_yoy.dropna(subset=[f'{data_cols[0]}_YOY_Abs'])
    
    print(f"✅ Created supply YOY data: {supply_yoy.shape}")
    print(f"📅 YOY periods: {len(supply_yoy)} months (starting from 2nd year)")
    
    # Show sample YOY data
    sample_cols = ['Year-Month', 'Norway', 'Norway_YOY_Pct', 'Total', 'Total_YOY_Pct']
    available_cols = [col for col in sample_cols if col in supply_yoy.columns]
    print(f"\n📊 SAMPLE SUPPLY YOY DATA:")
    print(supply_yoy.head()[available_cols].to_string())
    
    return supply_yoy

def save_monthly_tabs(demand_monthly, supply_monthly, demand_yoy, supply_yoy):
    """Save all monthly tabs to the master file"""
    
    print(f"\n💾 SAVING MONTHLY TABS TO MASTER FILE")
    print("=" * 60)
    
    master_file = 'European_Gas_Market_Master.xlsx'
    
    # Load existing demand and supply data
    demand_daily = pd.read_excel(master_file, sheet_name='Demand')
    supply_daily = pd.read_excel(master_file, sheet_name='Supply')
    
    # Save all tabs to master file
    with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
        # Original daily data
        demand_daily.to_excel(writer, sheet_name='Demand', index=False)
        supply_daily.to_excel(writer, sheet_name='Supply', index=False)
        
        # New monthly tabs
        demand_monthly.to_excel(writer, sheet_name='Demand Monthly', index=False)
        supply_monthly.to_excel(writer, sheet_name='Supply Monthly', index=False)
        demand_yoy.to_excel(writer, sheet_name='Demand Monthly YOY', index=False)
        supply_yoy.to_excel(writer, sheet_name='Supply Monthly YOY', index=False)
    
    print(f"✅ MASTER FILE UPDATED: {master_file}")
    print(f"📊 Total tabs: 6")
    print(f"   • Demand (daily): {demand_daily.shape}")
    print(f"   • Supply (daily): {supply_daily.shape}")
    print(f"   • Demand Monthly: {demand_monthly.shape}")
    print(f"   • Supply Monthly: {supply_monthly.shape}")
    print(f"   • Demand Monthly YOY: {demand_yoy.shape}")
    print(f"   • Supply Monthly YOY: {supply_yoy.shape}")
    
    # Also save individual CSV files for convenience
    demand_monthly.to_csv('European_Gas_Demand_Monthly.csv', index=False)
    supply_monthly.to_csv('European_Gas_Supply_Monthly.csv', index=False)
    demand_yoy.to_csv('European_Gas_Demand_Monthly_YOY.csv', index=False)
    supply_yoy.to_csv('European_Gas_Supply_Monthly_YOY.csv', index=False)
    
    print(f"✅ Individual CSV files also saved")

def main():
    """Main execution function"""
    
    print("📅 CREATING MONTHLY AGGREGATION TABS")
    print("=" * 80)
    print("Creating monthly averages and YOY comparisons for gas market data")
    print("=" * 80)
    
    try:
        # Load existing data
        demand_df, supply_df = load_existing_data()
        
        # Create monthly averages
        demand_monthly = create_demand_monthly(demand_df)
        supply_monthly = create_supply_monthly(supply_df)
        
        # Create YOY comparisons  
        demand_yoy = create_demand_monthly_yoy(demand_monthly)
        supply_yoy = create_supply_monthly_yoy(supply_monthly)
        
        # Save all tabs
        save_monthly_tabs(demand_monthly, supply_monthly, demand_yoy, supply_yoy)
        
        # Final summary
        print(f"\n🎉 MONTHLY TABS CREATION COMPLETE")
        print("=" * 80)
        print(f"✅ Demand Monthly: Monthly averages for all demand categories")
        print(f"✅ Supply Monthly: Monthly averages for all supply sources")
        print(f"✅ Demand Monthly YOY: Year-over-year changes in demand")
        print(f"✅ Supply Monthly YOY: Year-over-year changes in supply")
        print(f"✅ All tabs integrated into master Excel file")
        print(f"✅ Individual CSV files created for each monthly view")
        
        print(f"\n📁 UPDATED OUTPUT:")
        print(f"   🎯 European_Gas_Market_Master.xlsx (6 tabs total)")
        print(f"   📊 Individual monthly CSV files")
        
        print(f"\n🎯 SYSTEM READY FOR MONTHLY ANALYSIS!")
        
        return True
        
    except Exception as e:
        print(f"\n💥 ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)