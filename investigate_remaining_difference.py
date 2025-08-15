# Investigate the remaining 0.63 difference for Italy
import pandas as pd
import numpy as np
import json

def investigate_remaining_difference():
    """Deep dive into the remaining 0.63 difference for Italy"""
    
    print("🔬 INVESTIGATING REMAINING 0.63 ITALY DIFFERENCE")
    print("=" * 60)
    
    # Load the corrected target values
    with open('corrected_analysis_results.json', 'r') as f:
        analysis = json.load(f)
    target_italy = analysis['target_values']['Italy']  # 151.4659795006551
    
    print(f"Target Italy value: {target_italy:.9f}")
    
    # Load and process Bloomberg data exactly as our fixed system does
    bloomberg_data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True)
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    
    # Get normalization factors
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    # Apply normalization
    bloomberg_normalized = bloomberg_data.copy()
    for ticker in bloomberg_normalized.columns:
        if ticker in norm_factors:
            bloomberg_normalized[ticker] = bloomberg_normalized[ticker] * norm_factors[ticker]
    
    # Get the exact date we're comparing
    target_date = '2016-10-04'
    bloomberg_row = bloomberg_normalized[bloomberg_normalized.index.strftime('%Y-%m-%d') == target_date].iloc[0]
    
    # Get all Italy-related tickers and their exact values
    print(f"\n📊 ALL ITALY TICKERS ON {target_date}:")
    italy_dataset_rows = dataset[dataset['Region from'] == 'Italy']
    
    total_all_italy = 0
    total_demand_only = 0
    total_true_demand = 0
    
    for _, row in italy_dataset_rows.iterrows():
        ticker = row.get('Ticker')
        category = row.get('Category', '')
        region_to = row.get('Region to', '')
        description = row.get('Description', '')
        
        if ticker in bloomberg_row.index:
            val = bloomberg_row[ticker]
            norm_factor = norm_factors.get(ticker, 1.0)
            
            print(f"\n{ticker}:")
            print(f"  Category: {category}")
            print(f"  Region to: {region_to}")
            print(f"  Description: {description[:80]}...")
            print(f"  Raw value: {val/norm_factor:.9f}")
            print(f"  Normalized: {val:.9f}")
            
            total_all_italy += val
            
            if category == 'Demand':
                total_demand_only += val
                print(f"  ✅ INCLUDED in 'Demand' category")
                
                if region_to in ['Industrial', 'LDZ', 'Gas-to-Power']:
                    total_true_demand += val
                    print(f"  ✅ INCLUDED in TRUE DEMAND (Industrial/LDZ/Gas-to-Power)")
                else:
                    print(f"  ❌ EXCLUDED from TRUE DEMAND (region_to: {region_to})")
            else:
                print(f"  ❌ EXCLUDED (not 'Demand' category)")
    
    print(f"\n📊 ITALY TOTALS BREAKDOWN:")
    print(f"  All Italy tickers:        {total_all_italy:.9f}")
    print(f"  Only 'Demand' category:   {total_demand_only:.9f}")
    print(f"  True demand only:         {total_true_demand:.9f}")
    print(f"  LiveSheet target:         {target_italy:.9f}")
    print(f"  Difference from target:   {abs(total_true_demand - target_italy):.9f}")
    
    # Check if the LiveSheet might be using a different calculation
    print(f"\n🤔 POSSIBLE EXPLANATIONS FOR REMAINING DIFFERENCE:")
    diff = abs(total_true_demand - target_italy)
    print(f"  1. Data vintage difference: {diff:.9f}")
    print(f"  2. Different calculation methodology")
    print(f"  3. LiveSheet might exclude additional items")
    print(f"  4. Rounding/precision differences")
    
    # Let's also check what the raw LiveSheet value is for comparison
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    livesheet_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
    
    # Find row 15 (2016-10-04) and column 4 (Italy)
    livesheet_italy_raw = livesheet_data.iloc[15, 4]
    print(f"\n📋 LIVESHEET VERIFICATION:")
    print(f"  Raw LiveSheet Italy value: {livesheet_italy_raw:.9f}")
    print(f"  Our target value:          {target_italy:.9f}")
    print(f"  Match: {abs(livesheet_italy_raw - target_italy) < 0.000001}")
    
    # Check if there are any Bloomberg tickers we might be missing
    print(f"\n🔍 CHECKING FOR MISSING BLOOMBERG TICKERS:")
    demand_italy_tickers = italy_dataset_rows[italy_dataset_rows['Category'] == 'Demand']['Ticker'].tolist()
    missing_tickers = []
    for ticker in demand_italy_tickers:
        if ticker not in bloomberg_data.columns:
            missing_tickers.append(ticker)
    
    if missing_tickers:
        print(f"  ❌ Missing {len(missing_tickers)} Italy demand tickers from Bloomberg data:")
        for ticker in missing_tickers:
            print(f"     {ticker}")
    else:
        print(f"  ✅ All Italy demand tickers present in Bloomberg data")
    
    # Final analysis
    print(f"\n🎯 FINAL ANALYSIS:")
    if diff < 0.01:
        print(f"  ✅ EXCELLENT: Difference of {diff:.9f} is very small")
        print(f"     This is likely due to data vintage or minor methodology differences")
        print(f"     For practical purposes, this is a perfect match!")
    elif diff < 0.1:
        print(f"  ✅ GOOD: Difference of {diff:.9f} is acceptable")
        print(f"     Further investigation needed for perfect precision")
    else:
        print(f"  ❌ SIGNIFICANT: Difference of {diff:.9f} needs investigation")

if __name__ == "__main__":
    investigate_remaining_difference()