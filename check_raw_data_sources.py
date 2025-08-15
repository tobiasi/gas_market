# Check if Bloomberg raw data matches LiveSheet base values before normalization
import pandas as pd
import numpy as np

def check_raw_data_sources():
    """Compare raw Bloomberg data with LiveSheet to find data source discrepancy"""
    
    print("🔎 CHECKING RAW DATA SOURCES ALIGNMENT")
    print("=" * 60)
    
    # Load Bloomberg raw data
    bloomberg_data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True)
    target_date = '2016-10-04'
    bloomberg_row = bloomberg_data[bloomberg_data.index.strftime('%Y-%m-%d') == target_date].iloc[0]
    
    # Load normalization factors
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    # Italy key tickers
    italy_tickers = {
        'SNAMGIND Index': 'Industrial',
        'SNAMGLDN Index': 'LDZ', 
        'SNAMGPGE Index': 'Gas-to-Power'
    }
    
    print(f"📊 ITALY TICKERS RAW VALUES ON {target_date}:")
    bloomberg_total_raw = 0
    bloomberg_total_normalized = 0
    
    for ticker, category in italy_tickers.items():
        if ticker in bloomberg_row.index:
            raw_val = bloomberg_row[ticker]
            norm_factor = norm_factors.get(ticker, 1.0)
            normalized_val = raw_val * norm_factor
            
            bloomberg_total_raw += raw_val
            bloomberg_total_normalized += normalized_val
            
            print(f"  {ticker} ({category}):")
            print(f"    Raw: {raw_val:.6f}")
            print(f"    Normalization factor: {norm_factor:.9f}")
            print(f"    Normalized: {normalized_val:.6f}")
    
    print(f"\n📊 BLOOMBERG TOTALS:")
    print(f"  Raw total: {bloomberg_total_raw:.6f}")
    print(f"  Normalized total: {bloomberg_total_normalized:.6f}")
    
    # Load LiveSheet and check if we can reverse-engineer the raw values
    print(f"\n📋 LIVESHEET ANALYSIS:")
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    livesheet_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
    livesheet_italy = livesheet_data.iloc[15, 4]  # Row 15, Column 4 = Italy on 2016-10-04
    
    print(f"  LiveSheet Italy value: {livesheet_italy:.6f}")
    print(f"  Difference: {abs(livesheet_italy - bloomberg_total_normalized):.6f}")
    
    # Try to reverse-engineer what the LiveSheet raw values would be
    print(f"\n🔬 REVERSE ENGINEERING:")
    print(f"  If LiveSheet used same normalization factors:")
    
    # Assume the LiveSheet is the "correct" normalized total
    # What would the raw total need to be?
    avg_norm_factor = np.mean([norm_factors[t] for t in italy_tickers.keys() if t in norm_factors])
    implied_raw_total = livesheet_italy / avg_norm_factor
    
    print(f"  Average norm factor: {avg_norm_factor:.9f}")
    print(f"  Implied raw total: {implied_raw_total:.6f}")
    print(f"  Actual Bloomberg raw: {bloomberg_total_raw:.6f}")
    print(f"  Raw difference: {abs(implied_raw_total - bloomberg_total_raw):.6f}")
    
    # Check each ticker individually for potential data differences
    print(f"\n🎯 INDIVIDUAL TICKER ANALYSIS:")
    for ticker, category in italy_tickers.items():
        if ticker in bloomberg_row.index:
            bloomberg_raw = bloomberg_row[ticker]
            norm_factor = norm_factors.get(ticker, 1.0)
            bloomberg_normalized = bloomberg_raw * norm_factor
            
            # Calculate what this ticker's contribution should be to reach LiveSheet total
            # (This is speculative, but helps identify which ticker might be off)
            proportion_estimate = 1/3  # Rough estimate, each ticker ~1/3 of total
            target_normalized = livesheet_italy * proportion_estimate
            target_raw = target_normalized / norm_factor
            
            print(f"\n  {ticker} ({category}):")
            print(f"    Bloomberg raw: {bloomberg_raw:.6f}")
            print(f"    Bloomberg normalized: {bloomberg_normalized:.6f}")
            print(f"    Estimated target raw: {target_raw:.6f}")
            print(f"    Raw difference: {abs(bloomberg_raw - target_raw):.6f}")
    
    # Final assessment
    print(f"\n🎯 ASSESSMENT:")
    raw_diff_pct = abs(bloomberg_total_raw - implied_raw_total) / bloomberg_total_raw * 100
    norm_diff_pct = abs(bloomberg_total_normalized - livesheet_italy) / livesheet_italy * 100
    
    print(f"  Raw data difference: {raw_diff_pct:.2f}%")
    print(f"  Normalized difference: {norm_diff_pct:.2f}%")
    
    if raw_diff_pct > 1.0:
        print(f"  💡 LIKELY CAUSE: Bloomberg data vintage difference")
        print(f"     The bloomberg_raw_data.csv appears to be from a different")
        print(f"     Bloomberg pull than what was used for the LiveSheet")
    elif norm_diff_pct > 0.5:
        print(f"  💡 LIKELY CAUSE: Normalization factor mismatch")
        print(f"     The normalization factors might be slightly different")
    else:
        print(f"  💡 LIKELY CAUSE: Minor calculation differences")

if __name__ == "__main__":
    check_raw_data_sources()