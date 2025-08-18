# Debug script to check seasonal data structure
import pandas as pd

# Load data to test
master_file = 'European_Gas_Market_Master.xlsx'
demand_daily = pd.read_excel(master_file, sheet_name='Demand')
demand_daily['Date'] = pd.to_datetime(demand_daily['Date'])

demand_columns = ['Italy', 'Netherlands', 'GB']

print("Available columns in demand_daily:")
print(list(demand_daily.columns))
print(f"\nTesting with columns: {demand_columns}")

# Test the seasonal creation logic
seasonal_dfs = {}

for col in demand_columns:
    print(f"\nProcessing {col}...")
    
    # Create empty DataFrame for years 2017-2030
    years = list(range(2017, 2031))
    days = list(range(1, 366))  # 365 days
    seasonal_df = pd.DataFrame(index=days, columns=years)
    seasonal_df.index.name = 'Day_of_Year'
    
    # Fill with a few data points for testing
    for _, row in demand_daily.head(10).iterrows():
        date = row['Date']
        year = date.year
        
        # Skip Feb 29 (leap year day)
        if date.month == 2 and date.day == 29:
            continue
        
        # Calculate day of year (adjusting for leap years)
        if date.month > 2 and pd.Timestamp(year, 12, 31).dayofyear == 366:
            day_of_year = date.dayofyear - 1
        else:
            day_of_year = date.dayofyear
        
        if year in years and 1 <= day_of_year <= 365:
            value = row[col]
            if pd.notna(value):
                seasonal_df.at[day_of_year, year] = round(value, 2)
    
    seasonal_dfs[col] = seasonal_df
    print(f"Added {col} to seasonal_dfs. Keys now: {list(seasonal_dfs.keys())}")

print(f"\nFinal seasonal_dfs keys: {list(seasonal_dfs.keys())}")
print(f"Total keys: {len(seasonal_dfs)}")