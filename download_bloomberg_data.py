# -*- coding: utf-8 -*-
"""
Bloomberg Data Downloader - Get exact raw data for testing and development
This script downloads the actual Bloomberg data and saves it to CSV/Excel
for testing without needing Bloomberg API access every time.
"""

import pandas as pd
import numpy as np
import warnings
from datetime import datetime
import os

warnings.filterwarnings('ignore')

def download_and_save_bloomberg_data():
    """Download Bloomberg data and save for offline testing"""
    
    print("📡 BLOOMBERG DATA DOWNLOADER")
    print("=" * 70)
    
    # Load ticker configuration
    print("📊 Loading ticker configuration...")
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    
    # Create normalization mapping
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    print(f"✅ Loaded {len(norm_factors)} normalization factors")
    
    # Get unique tickers
    tickers = list(set(dataset.Ticker.dropna().tolist()))
    print(f"📈 Found {len(tickers)} unique tickers to download")
    
    # Download Bloomberg data
    print("🔄 Downloading from Bloomberg...")
    try:
        from xbbg import blp
        
        print("   Connecting to Bloomberg API...")
        raw_data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST")
        
        # Clean up the data structure
        if len(raw_data.columns.levels) > 1:
            data = raw_data.droplevel(axis=1, level=1)  # Remove field level
        else:
            data = raw_data
        
        print(f"✅ Downloaded data: {data.shape}")
        
        # Check for missing tickers
        available_tickers = set(data.columns)
        requested_tickers = set(tickers)
        missing_tickers = requested_tickers - available_tickers
        
        if missing_tickers:
            print(f"⚠️  Missing {len(missing_tickers)} tickers from Bloomberg:")
            for i, ticker in enumerate(sorted(missing_tickers)[:10], 1):
                print(f"   {i:2d}. {ticker}")
            if len(missing_tickers) > 10:
                print(f"   ... and {len(missing_tickers) - 10} more")
        
        print(f"📊 Successfully downloaded {len(available_tickers)} out of {len(requested_tickers)} tickers")
        
    except ImportError:
        print("❌ xbbg module not available")
        print("   To download real data, install: pip install xbbg")
        return False
        
    except Exception as e:
        print(f"❌ Bloomberg download failed: {e}")
        print("   Check Bloomberg terminal connection and permissions")
        return False
    
    # Save raw data
    print("\n💾 SAVING RAW DATA")
    print("=" * 40)
    
    try:
        # Save as CSV (faster loading)
        raw_csv_file = 'bloomberg_raw_data.csv'
        data.to_csv(raw_csv_file)
        print(f"✅ Saved raw data: {raw_csv_file}")
        
        # Save as Excel (better for inspection)
        raw_excel_file = 'bloomberg_raw_data.xlsx'
        with pd.ExcelWriter(raw_excel_file, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='Raw_Data')
            
            # Also save ticker metadata
            dataset.to_excel(writer, sheet_name='Ticker_Metadata')
            
            # Save normalization factors summary
            norm_df = pd.DataFrame(list(norm_factors.items()), 
                                 columns=['Ticker', 'Normalization_Factor'])
            norm_df.to_excel(writer, sheet_name='Normalization_Factors', index=False)
        
        print(f"✅ Saved with metadata: {raw_excel_file}")
        
    except Exception as e:
        print(f"❌ Save failed: {e}")
        return False
    
    # Apply normalization and save normalized version
    print("\n🔧 CREATING NORMALIZED VERSION")
    print("=" * 40)
    
    try:
        data_normalized = data.copy()
        normalized_count = 0
        
        for ticker in data_normalized.columns:
            if ticker in norm_factors:
                data_normalized[ticker] = data_normalized[ticker] * norm_factors[ticker]
                normalized_count += 1
        
        print(f"✅ Applied normalization to {normalized_count}/{len(data_normalized.columns)} tickers")
        
        # Save normalized data
        normalized_csv = 'bloomberg_normalized_data.csv'
        data_normalized.to_csv(normalized_csv)
        print(f"✅ Saved normalized data: {normalized_csv}")
        
        # Quick verification - check Italy values
        italy_tickers = [ticker for ticker in data_normalized.columns 
                        if any(italy_word in str(ticker).upper() for italy_word in ['SNAM', 'ITALY'])]
        
        if italy_tickers:
            print(f"\n🇮🇹 ITALY VERIFICATION:")
            italy_sample = data_normalized[italy_tickers].iloc[0]  # First row
            italy_total = italy_sample.sum()
            print(f"   Italy tickers found: {len(italy_tickers)}")
            print(f"   Sample Italy total: {italy_total:.2f}")
            print(f"   Expected target: ~115.24")
            print(f"   Match: {'✅' if abs(italy_total - 115) < 20 else '❌'}")
            
            # Show breakdown
            print(f"   Breakdown:")
            for ticker in italy_tickers:
                value = italy_sample[ticker]
                print(f"      {ticker}: {value:.2f}")
        
    except Exception as e:
        print(f"⚠️  Normalization failed: {e}")
    
    # Create data info summary
    print("\n📋 DATA SUMMARY")
    print("=" * 40)
    
    print(f"📊 Data range: {data.index[0]} to {data.index[-1]}")
    print(f"📊 Total days: {len(data)}")
    print(f"📊 Total tickers: {len(data.columns)}")
    print(f"📊 File size: ~{os.path.getsize(raw_csv_file) / 1024 / 1024:.1f} MB" if os.path.exists(raw_csv_file) else "")
    
    # Data quality check
    null_counts = data.isnull().sum()
    problematic_tickers = null_counts[null_counts > len(data) * 0.1]  # >10% nulls
    
    if len(problematic_tickers) > 0:
        print(f"⚠️  {len(problematic_tickers)} tickers with >10% missing data:")
        for ticker in problematic_tickers.head(5).index:
            pct_missing = (null_counts[ticker] / len(data)) * 100
            print(f"      {ticker}: {pct_missing:.1f}% missing")
    else:
        print("✅ Data quality looks good (<10% missing values)")
    
    print("\n🎯 USAGE")
    print("=" * 40)
    print("To use this data in other scripts:")
    print("   data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True)")
    print("   normalized_data = pd.read_csv('bloomberg_normalized_data.csv', index_col=0, parse_dates=True)")
    print("")
    print("Files created:")
    print("   📁 bloomberg_raw_data.csv - Raw Bloomberg data")
    print("   📁 bloomberg_normalized_data.csv - With normalization applied")
    print("   📁 bloomberg_raw_data.xlsx - Excel version with metadata")
    
    return True

def create_data_loader():
    """Create a helper function to load the downloaded data"""
    
    loader_code = '''
def load_bloomberg_data(normalized=True, start_date=None, end_date=None):
    """
    Load downloaded Bloomberg data
    
    Args:
        normalized (bool): Load normalized data if True, raw data if False
        start_date (str): Filter data from this date (YYYY-MM-DD)
        end_date (str): Filter data to this date (YYYY-MM-DD)
    
    Returns:
        pd.DataFrame: Bloomberg data
    """
    import pandas as pd
    
    filename = 'bloomberg_normalized_data.csv' if normalized else 'bloomberg_raw_data.csv'
    
    try:
        data = pd.read_csv(filename, index_col=0, parse_dates=True)
        
        # Apply date filters if provided
        if start_date:
            data = data[data.index >= start_date]
        if end_date:
            data = data[data.index <= end_date]
            
        print(f"✅ Loaded {'normalized' if normalized else 'raw'} data: {data.shape}")
        return data
        
    except FileNotFoundError:
        print(f"❌ File not found: {filename}")
        print("   Run download_bloomberg_data.py first to download the data")
        return None
    except Exception as e:
        print(f"❌ Error loading data: {e}")
        return None

# Usage examples:
# raw_data = load_bloomberg_data(normalized=False)
# normalized_data = load_bloomberg_data(normalized=True)
# recent_data = load_bloomberg_data(normalized=True, start_date='2023-01-01')
'''
    
    with open('bloomberg_data_loader.py', 'w') as f:
        f.write(loader_code)
    
    print("✅ Created bloomberg_data_loader.py helper module")

def main():
    """Main execution"""
    
    print("🎯 BLOOMBERG DATA DOWNLOADER & SAVER")
    print("=" * 80)
    print("Downloads real Bloomberg data for offline development and testing")
    print("=" * 80)
    
    success = download_and_save_bloomberg_data()
    
    if success:
        create_data_loader()
        
        print("\n🎉 SUCCESS!")
        print("=" * 50)
        print("✅ Bloomberg data downloaded and saved")
        print("✅ Normalized version created")
        print("✅ Helper loader function created")
        print("✅ Ready for offline development!")
        
        return True
    else:
        print("\n❌ FAILED!")
        print("=" * 50)
        print("Could not download Bloomberg data")
        print("Check Bloomberg connection and try again")
        
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)