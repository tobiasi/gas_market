# -*- coding: utf-8 -*-
"""
Analyze the uploaded Bloomberg raw data file
"""

import pandas as pd
import numpy as np

def analyze_bloomberg_raw_data():
    """Analyze the uploaded Bloomberg raw data"""
    
    print(f"\n{'='*80}")
    print("📊 ANALYZING UPLOADED BLOOMBERG RAW DATA")
    print(f"{'='*80}")
    
    # Check the new LiveSheet file
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    
    try:
        # Get sheet names
        excel_file = pd.ExcelFile(filename)
        print(f"📁 File: {filename}")
        print(f"📋 Available sheets:")
        for i, sheet in enumerate(excel_file.sheet_names, 1):
            print(f"  {i:2d}. {sheet}")
        
        # Look for sheets that might contain raw Bloomberg data
        potential_sheets = ['Raw_Data', 'Bloomberg_Data', 'Data', 'Multiticker', 'Raw', 'bloomberg_raw_data']
        
        data_sheet = None
        for sheet in excel_file.sheet_names:
            if any(keyword.lower() in sheet.lower() for keyword in potential_sheets):
                data_sheet = sheet
                break
        
        if not data_sheet:
            # Try the first sheet or Multiticker
            if 'Multiticker' in excel_file.sheet_names:
                data_sheet = 'Multiticker'
            else:
                data_sheet = excel_file.sheet_names[0]
        
        print(f"\n📊 Reading sheet: '{data_sheet}'")
        
        # Read the data
        raw_data = pd.read_excel(filename, sheet_name=data_sheet, index_col=0)
        print(f"✅ Loaded data shape: {raw_data.shape}")
        print(f"📅 Date range: {raw_data.index[0]} to {raw_data.index[-1]}")
        
        # Check column structure
        print(f"\n🔍 DATA STRUCTURE ANALYSIS:")
        print(f"   Total columns: {len(raw_data.columns)}")
        print(f"   Sample columns: {list(raw_data.columns[:10])}")
        
        # Look for Italy tickers
        italy_tickers = []
        for col in raw_data.columns:
            if any(italy_keyword in str(col).upper() for italy_keyword in ['SNAM', 'ITALY']):
                italy_tickers.append(col)
        
        print(f"\n🇮🇹 ITALY TICKERS FOUND: {len(italy_tickers)}")
        if italy_tickers:
            for ticker in italy_tickers:
                sample_val = raw_data[ticker].iloc[0] if not pd.isna(raw_data[ticker].iloc[0]) else 'NaN'
                print(f"   {ticker}: {sample_val}")
        
        # Save as CSV for easier processing
        output_csv = 'bloomberg_raw_data_uploaded.csv'
        raw_data.to_csv(output_csv)
        print(f"\n💾 Saved as CSV: {output_csv}")
        
        # Quick test with normalization
        print(f"\n🧪 QUICK NORMALIZATION TEST:")
        
        # Load normalization factors
        try:
            dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
            norm_factors = {}
            for _, row in dataset.iterrows():
                ticker = row.get('Ticker')
                norm_factor = row.get('Normalization factor', 1.0)
                if pd.notna(ticker) and pd.notna(norm_factor):
                    norm_factors[ticker] = float(norm_factor)
            
            print(f"   Loaded {len(norm_factors)} normalization factors")
            
            # Test Italy normalization
            italy_raw_total = 0
            italy_normalized_total = 0
            italy_count = 0
            
            for ticker in italy_tickers:
                if ticker in raw_data.columns and ticker in norm_factors:
                    raw_val = raw_data[ticker].iloc[0]
                    if pd.notna(raw_val):
                        norm_val = raw_val * norm_factors[ticker]
                        italy_raw_total += raw_val
                        italy_normalized_total += norm_val
                        italy_count += 1
                        print(f"   {ticker}: {raw_val:.2f} → {norm_val:.2f}")
            
            print(f"\n   Italy totals:")
            print(f"     Raw sum: {italy_raw_total:.2f}")
            print(f"     Normalized sum: {italy_normalized_total:.2f}")
            print(f"     Target: ~115.24")
            print(f"     Match: {'✅' if abs(italy_normalized_total - 115.24) < 10 else '❌'}")
            
        except Exception as e:
            print(f"   ⚠️  Could not test normalization: {e}")
        
        return raw_data
        
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return None

def create_normalized_version(raw_data):
    """Create normalized version of the raw data"""
    
    if raw_data is None:
        return None
    
    print(f"\n🔧 CREATING NORMALIZED VERSION")
    print("=" * 50)
    
    try:
        # Load normalization factors
        dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
        norm_factors = {}
        for _, row in dataset.iterrows():
            ticker = row.get('Ticker')
            norm_factor = row.get('Normalization factor', 1.0)
            if pd.notna(ticker) and pd.notna(norm_factor):
                norm_factors[ticker] = float(norm_factor)
        
        # Apply normalization
        normalized_data = raw_data.copy()
        normalized_count = 0
        
        for ticker in normalized_data.columns:
            if ticker in norm_factors:
                normalized_data[ticker] = normalized_data[ticker] * norm_factors[ticker]
                normalized_count += 1
        
        print(f"✅ Applied normalization to {normalized_count}/{len(normalized_data.columns)} tickers")
        
        # Save normalized version
        output_csv = 'bloomberg_normalized_data_uploaded.csv'
        normalized_data.to_csv(output_csv)
        print(f"✅ Saved normalized data: {output_csv}")
        
        return normalized_data
        
    except Exception as e:
        print(f"❌ Normalization failed: {e}")
        return None

if __name__ == "__main__":
    raw_data = analyze_bloomberg_raw_data()
    normalized_data = create_normalized_version(raw_data)