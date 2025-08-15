# Debug why Bloomberg processing doesn't match our perfect LiveSheet results
import pandas as pd
import numpy as np
import json

def compare_livesheet_vs_bloomberg():
    """Compare the exact same date in LiveSheet vs Bloomberg data"""
    
    print("🔍 DEBUGGING BLOOMBERG vs LIVESHEET MISMATCH")
    print("=" * 70)
    
    # 1. Get the PERFECT values from LiveSheet (our reference)
    print("1️⃣ LOADING LIVESHEET REFERENCE VALUES...")
    with open('corrected_analysis_results.json', 'r') as f:
        analysis = json.load(f)
    target_values = analysis['target_values']
    column_mapping = analysis['column_mapping']
    
    print("Perfect LiveSheet values (2016-10-04):")
    for key, value in target_values.items():
        print(f"  {key}: {value}")
    
    # 2. Load LiveSheet directly and extract the same date
    print("\n2️⃣ EXTRACTING LIVESHEET DATA DIRECTLY...")
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    livesheet_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
    
    # Find the exact row for 2016-10-04
    target_row_idx = None
    for i in range(12, min(50, livesheet_data.shape[0])):
        date_val = livesheet_data.iloc[i, 1]
        if pd.notna(date_val):
            try:
                date_parsed = pd.to_datetime(str(date_val))
                if date_parsed.strftime('%Y-%m-%d') == '2016-10-04':
                    target_row_idx = i
                    break
            except:
                continue
    
    if target_row_idx:
        print(f"Found 2016-10-04 at LiveSheet row {target_row_idx}")
        livesheet_values = {}
        for name, col_idx in column_mapping.items():
            val = livesheet_data.iloc[target_row_idx, col_idx]
            livesheet_values[name] = float(val) if pd.notna(val) else 0
            print(f"  {name}: {val} (column {col_idx})")
    
    # 3. Load Bloomberg raw data for the same date
    print("\n3️⃣ LOADING BLOOMBERG RAW DATA...")
    bloomberg_data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True)
    
    # Find 2016-10-04 in Bloomberg data
    target_date = '2016-10-04'
    if target_date in [str(d)[:10] for d in bloomberg_data.index]:
        bloomberg_row = bloomberg_data[bloomberg_data.index.strftime('%Y-%m-%d') == target_date].iloc[0]
        print(f"Found 2016-10-04 in Bloomberg data")
        print(f"Bloomberg data shape: {bloomberg_data.shape}")
        print(f"Available tickers: {len(bloomberg_data.columns)}")
    else:
        print("❌ 2016-10-04 not found in Bloomberg data!")
        return
    
    # 4. Load normalization factors
    print("\n4️⃣ LOADING NORMALIZATION FACTORS...")
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    print(f"Loaded {len(norm_factors)} normalization factors")
    
    # 5. Find Italy-related tickers and check their values
    print("\n5️⃣ ANALYZING ITALY TICKERS...")
    italy_tickers = []
    for ticker in bloomberg_data.columns:
        if 'SNAM' in ticker and any(keyword in ticker for keyword in ['GIND', 'GLDN', 'GPGE', 'CLGG', 'GOTH']):
            italy_tickers.append(ticker)
    
    print(f"Found {len(italy_tickers)} Italy-related tickers:")
    italy_total_raw = 0
    italy_total_normalized = 0
    
    for ticker in italy_tickers:
        raw_val = bloomberg_row[ticker] if ticker in bloomberg_row.index else 0
        norm_factor = norm_factors.get(ticker, 1.0)
        normalized_val = raw_val * norm_factor
        italy_total_raw += raw_val
        italy_total_normalized += normalized_val
        print(f"  {ticker}: {raw_val:.2f} → {normalized_val:.2f} (factor: {norm_factor:.9f})")
    
    print(f"\nItaly totals:")
    print(f"  Raw Bloomberg sum: {italy_total_raw:.2f}")
    print(f"  Normalized sum: {italy_total_normalized:.2f}")
    print(f"  LiveSheet value: {livesheet_values.get('Italy', 0):.2f}")
    print(f"  Difference: {abs(italy_total_normalized - livesheet_values.get('Italy', 0)):.2f}")
    
    # 6. Check if we're missing some Italy tickers
    print("\n6️⃣ CHECKING FOR MISSING ITALY DATA...")
    
    # Look at use4.xlsx to see what Italy tickers should exist
    italy_dataset_rows = dataset[dataset['Region from'] == 'Italy']
    print(f"Italy-related rows in use4.xlsx: {len(italy_dataset_rows)}")
    
    for _, row in italy_dataset_rows.iterrows():
        ticker = row.get('Ticker')
        category = row.get('Category', '')
        region_to = row.get('Region to', '')
        description = row.get('Description', '')
        
        if 'Demand' in category:
            print(f"  {ticker}: {category} → {region_to} ({description[:50]})")
            
            if ticker in bloomberg_data.columns:
                val = bloomberg_row[ticker]
                norm_factor = norm_factors.get(ticker, 1.0)
                normalized = val * norm_factor
                print(f"    Raw: {val:.2f}, Normalized: {normalized:.2f}")
            else:
                print(f"    ❌ MISSING from Bloomberg data!")
    
    # 7. Compare with our MultiIndex approach
    print("\n7️⃣ DEBUGGING MULTIINDEX APPROACH...")
    
    # Try to replicate the MultiIndex creation and see where it goes wrong
    successful_tickers = []
    level_2 = []  # Category  
    level_3 = []  # Region from
    level_4 = []  # Region to
    
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        if pd.notna(ticker) and ticker in bloomberg_data.columns:
            successful_tickers.append(ticker)
            level_2.append(row.get('Category', ''))
            level_3.append(row.get('Region from', ''))
            level_4.append(row.get('Region to', ''))
    
    # Apply normalization to Bloomberg data
    bloomberg_normalized = bloomberg_data.copy()
    for ticker in bloomberg_normalized.columns:
        if ticker in norm_factors:
            bloomberg_normalized[ticker] = bloomberg_normalized[ticker] * norm_factors[ticker]
    
    # Create MultiIndex structure
    if successful_tickers:
        ticker_data = bloomberg_normalized[successful_tickers]
        multi_index = pd.MultiIndex.from_arrays([[''] * len(successful_tickers), [''] * len(successful_tickers), level_2, level_3, level_4])
        full_data = pd.DataFrame(ticker_data.values, index=ticker_data.index, columns=multi_index)
        
        # Now check Italy demand specifically
        italy_mask = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(3) == 'Italy')
        italy_series = full_data.iloc[:, italy_mask]
        
        print(f"Italy demand series found: {italy_series.shape[1]}")
        
        if italy_series.shape[1] > 0:
            italy_multiindex_total = italy_series.sum(axis=1, skipna=True)
            target_row_multiindex = italy_multiindex_total[italy_multiindex_total.index.strftime('%Y-%m-%d') == target_date].iloc[0]
            
            print(f"Italy MultiIndex total: {target_row_multiindex:.2f}")
            print(f"LiveSheet Italy: {livesheet_values.get('Italy', 0):.2f}")
            print(f"Difference: {abs(target_row_multiindex - livesheet_values.get('Italy', 0)):.2f}")
            
            # Show individual series
            print("Individual Italy demand series:")
            for i, col in enumerate(italy_series.columns):
                val = italy_series.iloc[italy_series.index.strftime('%Y-%m-%d') == target_date, i].iloc[0]
                print(f"  {col[4]}: {val:.2f}")  # region_to is at index 4
        
    print("\n" + "=" * 70)
    print("🎯 ANALYSIS COMPLETE - Check the differences above!")

if __name__ == "__main__":
    compare_livesheet_vs_bloomberg()