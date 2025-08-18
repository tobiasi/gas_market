#!/usr/bin/env python3
"""
Debug actual data values in matched columns.
"""

import pandas as pd
import sys
sys.path.append("C:/development/commodities")

def debug_data_values():
    print("🔍 DEBUGGING DATA VALUES IN MATCHED COLUMNS")
    print("=" * 80)
    
    # Load Excel file
    df_full = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    # Check structure  
    data_start = 19 - 1  # Row 19 (0-based = 18)
    
    # Get dates column (column B)
    dates = df_full.iloc[data_start:, 1].dropna()
    dates_df = pd.to_datetime(dates.values)
    
    print(f"📊 Data starts at row {data_start + 1}")
    print(f"📅 Date range: {dates_df.min()} to {dates_df.max()}")
    print(f"🔢 Total data rows: {len(dates_df)}")
    
    # Test specific columns that should have matches
    test_columns = [
        {'name': 'MAB_to_Austria', 'excel_cols': ['D', 'E', 'F'], 'indices': [3, 4, 5]},
        {'name': 'Norway_to_Europe', 'excel_cols': ['AH'], 'indices': [33]},  # AH = 34th column (0-based = 33)
        {'name': 'Austria_Production', 'excel_cols': ['H', 'I'], 'indices': [7, 8]},
        {'name': 'Germany_Production', 'excel_cols': ['W', 'X'], 'indices': [22, 23]}
    ]
    
    # Check specific validation date
    target_date = '2016-10-01'
    try:
        target_idx = dates_df[dates_df.dt.strftime('%Y-%m-%d') == target_date].index[0]
        data_row_idx = data_start + target_idx
        print(f"\n📅 Target date {target_date} found at data row {data_row_idx + 1}")
    except:
        print(f"\n❌ Target date {target_date} not found")
        return
    
    print(f"\n🔍 CHECKING DATA VALUES FOR {target_date}:")
    print("   Route                    Excel Col   Index   Raw Value      Numeric Value")
    print("   " + "-" * 75)
    
    for test_col in test_columns:
        print(f"\n   {test_col['name']}:")
        total_value = 0
        
        for excel_col, col_idx in zip(test_col['excel_cols'], test_col['indices']):
            if col_idx < len(df_full.columns):
                raw_value = df_full.iloc[data_row_idx, col_idx]
                
                # Try to convert to numeric
                try:
                    numeric_value = pd.to_numeric(raw_value, errors='coerce')
                    if pd.notna(numeric_value):
                        total_value += numeric_value
                    else:
                        numeric_value = "NaN"
                except:
                    numeric_value = "Error"
                
                print(f"      {excel_col:10} {col_idx:6d} {str(raw_value):>12} {str(numeric_value):>12}")
            else:
                print(f"      {excel_col:10} {col_idx:6d} {'OUT_OF_BOUNDS':>25}")
        
        print(f"      TOTAL: {total_value}")
    
    # Check a few more dates to see if there's any data
    print(f"\n🔍 SAMPLE DATA CHECK - First 5 dates:")
    print("   Date         MAB(D)   MAB(E)   Norway(AH)  Austria(H)  Germany(W)")
    print("   " + "-" * 65)
    
    for i in range(min(5, len(dates_df))):
        data_row = data_start + i
        date_str = dates_df.iloc[i].strftime('%Y-%m-%d')
        
        values = []
        for col_idx in [3, 4, 33, 7, 22]:  # D, E, AH, H, W
            if col_idx < len(df_full.columns):
                raw_val = df_full.iloc[data_row, col_idx]
                try:
                    num_val = pd.to_numeric(raw_val, errors='coerce')
                    values.append(f"{num_val:.1f}" if pd.notna(num_val) else "NaN")
                except:
                    values.append("Err")
            else:
                values.append("OOB")
        
        print(f"   {date_str} {values[0]:>8} {values[1]:>8} {values[2]:>10} {values[3]:>10} {values[4]:>10}")

if __name__ == "__main__":
    debug_data_values()