#!/usr/bin/env python3
"""
Debug date finding issue.
"""

import pandas as pd
import sys
sys.path.append("C:/development/commodities")

def debug_date_issue():
    print("🔍 DEBUGGING DATE FINDING ISSUE")
    print("=" * 80)
    
    # Load Excel file
    df_full = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    data_start = 19 - 1  # Row 19 (0-based = 18)
    
    # Get dates column (column B)
    dates_raw = df_full.iloc[data_start:, 1].dropna()
    print(f"📊 Found {len(dates_raw)} raw date values")
    print(f"🔍 First 10 raw date values:")
    for i in range(min(10, len(dates_raw))):
        print(f"   {i}: {dates_raw.iloc[i]} (type: {type(dates_raw.iloc[i])})")
    
    # Convert to datetime
    dates_df = pd.to_datetime(dates_raw.values)
    print(f"\n📅 Converted to datetime:")
    print(f"   Date range: {dates_df.min()} to {dates_df.max()}")
    print(f"   Total dates: {len(dates_df)}")
    
    # Check first few dates
    print(f"\n🔍 First 10 converted dates:")
    for i in range(min(10, len(dates_df))):
        date_str = dates_df[i].strftime('%Y-%m-%d')
        print(f"   {i}: {date_str}")
    
    # Try to find target date 
    target_date = '2016-10-01'
    matching_dates = dates_df[dates_df.strftime('%Y-%m-%d') == target_date]
    
    if len(matching_dates) > 0:
        print(f"\n✅ Found {target_date}: {len(matching_dates)} matches")
        target_idx = matching_dates.index[0] 
        print(f"   First match at index: {target_idx}")
    else:
        print(f"\n❌ {target_date} not found")
        
        # Check if it might be in a different format
        print("🔍 Checking for approximate matches:")
        for i in range(min(20, len(dates_df))):
            date_str = dates_df[i].strftime('%Y-%m-%d')
            if '2016-10' in date_str:
                print(f"   Found October 2016 date: {date_str} at index {i}")
    
    # Test with first available date instead
    if len(dates_df) > 0:
        first_date = dates_df[0]
        first_date_str = first_date.strftime('%Y-%m-%d')
        data_row_idx = data_start + 0
        
        print(f"\n🔍 TESTING WITH FIRST AVAILABLE DATE: {first_date_str}")
        print("   Column   Index   Raw Value      Numeric Value")
        print("   " + "-" * 50)
        
        # Test key columns
        test_indices = [3, 4, 7, 8, 22, 23, 33]  # D, E, H, I, W, X, AH
        test_names = ['D', 'E', 'H', 'I', 'W', 'X', 'AH']
        
        for name, col_idx in zip(test_names, test_indices):
            if col_idx < len(df_full.columns):
                raw_value = df_full.iloc[data_row_idx, col_idx]
                
                try:
                    numeric_value = pd.to_numeric(raw_value, errors='coerce')
                    print(f"   {name:8} {col_idx:6d} {str(raw_value):>12} {str(numeric_value):>12}")
                except Exception as e:
                    print(f"   {name:8} {col_idx:6d} {str(raw_value):>12} ERROR: {str(e)}")

if __name__ == "__main__":
    debug_date_issue()