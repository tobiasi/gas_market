# Simple test to see if we can run basic analysis without kernel restart
import pandas as pd
import numpy as np
import json
from datetime import datetime
import gc
import psutil
import os

def log_memory():
    """Log current memory usage"""
    process = psutil.Process(os.getpid())
    memory_mb = process.memory_info().rss / 1024 / 1024
    print(f"Memory: {memory_mb:.1f} MB")
    return memory_mb

def simple_gas_analysis():
    """Simple version that should work without memory issues"""
    
    print("🧪 SIMPLE GAS MARKET ANALYSIS TEST")
    print("=" * 60)
    
    # Setup
    analysis_data = {
        'target_values': {'Italy': 151.466, 'Netherlands': 90.493, 'GB': 97.740, 'Total': 767.693},
        'column_mapping': {'Italy': 4, 'Netherlands': 20, 'GB': 21, 'Total': 12}
    }
    
    with open('analysis_results.json', 'w') as f:
        json.dump(analysis_data, f, indent=2)
    
    log_memory()
    
    try:
        # Load data
        print("Loading data...")
        filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
        target_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
        print(f"Data loaded: {target_data.shape}")
        log_memory()
        
        # Extract basic demand data
        print("Extracting demand data...")
        demand_data = []
        dates = []
        
        for i in range(12, min(1012, target_data.shape[0])):  # First 1000 rows only
            date_val = target_data.iloc[i, 1]
            if pd.notna(date_val):
                try:
                    date_parsed = pd.to_datetime(str(date_val))
                    dates.append(date_parsed)
                    
                    demand_data.append({
                        'Italy': target_data.iloc[i, 4],
                        'Netherlands': target_data.iloc[i, 20], 
                        'GB': target_data.iloc[i, 21],
                        'Total': target_data.iloc[i, 12]
                    })
                except:
                    continue
        
        demand_df = pd.DataFrame(demand_data, index=dates)
        print(f"Demand data created: {demand_df.shape}")
        log_memory()
        
        # Clear large data
        del target_data
        gc.collect()
        print("Cleared target_data")
        log_memory()
        
        # Create simple monthly data
        print("Creating monthly data...")
        demand_monthly = demand_df.resample('ME').mean().round(2)
        demand_monthly.reset_index(inplace=True)
        print(f"Monthly data: {demand_monthly.shape}")
        log_memory()
        
        # Simple YOY calculation
        print("Creating YOY data...")
        demand_yoy = demand_monthly.copy()
        for col in ['Italy', 'Netherlands', 'GB', 'Total']:
            if col in demand_yoy.columns:
                demand_yoy[f'{col}_YOY_Pct'] = ((demand_yoy[col] / demand_yoy[col].shift(12)) - 1) * 100
        
        demand_yoy = demand_yoy.round(2)
        print(f"YOY data: {demand_yoy.shape}")
        log_memory()
        
        # Save simple master file
        print("Saving master file...")
        master_file = 'European_Gas_Market_Simple.xlsx'
        
        with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
            demand_df.reset_index().to_excel(writer, sheet_name='Demand', index=False)
            demand_monthly.to_excel(writer, sheet_name='Demand Monthly', index=False)
            demand_yoy.to_excel(writer, sheet_name='Demand YOY', index=False)
        
        print(f"✅ Simple master file saved: {master_file}")
        log_memory()
        
        # Test one simple seasonal plot
        print("Creating one seasonal plot...")
        years = [2017, 2018, 2019, 2020, 2021]
        days = list(range(1, 366))
        seasonal_df = pd.DataFrame(index=days, columns=years)
        
        # Fill with Italy data only
        for _, row in demand_df.head(365).iterrows():  # First year only
            date = row.name
            year = date.year
            if year in years:
                day_of_year = date.dayofyear
                if 1 <= day_of_year <= 365:
                    seasonal_df.at[day_of_year, year] = round(row['Italy'], 2)
        
        print(f"Seasonal plot created: {seasonal_df.shape}")
        log_memory()
        
        # Save seasonal file
        seasonal_file = 'European_Gas_Seasonal_Simple.xlsx'
        with pd.ExcelWriter(seasonal_file, engine='openpyxl') as writer:
            seasonal_df.to_excel(writer, sheet_name='Italy_Daily')
        
        print(f"✅ Simple seasonal file saved: {seasonal_file}")
        log_memory()
        
        # Verify accuracy
        print("\nVerifying accuracy...")
        target_date = '2016-10-04'
        if target_date in [str(d)[:10] for d in demand_df.index]:
            target_row = demand_df[demand_df.index.strftime('%Y-%m-%d') == target_date].iloc[0]
            
            matches = 0
            for key, expected in analysis_data['target_values'].items():
                if key in target_row.index:
                    actual = target_row[key]
                    if abs(actual - expected) < 0.001:
                        matches += 1
                        print(f"✅ {key}: {actual:.3f} (expected: {expected:.3f})")
                    else:
                        print(f"❌ {key}: {actual:.3f} (expected: {expected:.3f})")
            
            accuracy = (matches / len(analysis_data['target_values'])) * 100
            print(f"Accuracy: {accuracy:.1f}%")
        
        print(f"\n🎉 SIMPLE ANALYSIS COMPLETE!")
        print(f"✅ Created {master_file} and {seasonal_file}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = simple_gas_analysis()
    print(f"\nResult: {'SUCCESS' if success else 'FAILED'}")