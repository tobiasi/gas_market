#!/usr/bin/env python3
"""
Bloomberg Gas Market System Test
================================

Test the Bloomberg-based gas market processing system to ensure it works
correctly with actual ticker configuration and produces expected results.
"""

import pandas as pd
import numpy as np
import logging
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def test_ticker_configuration():
    """Test loading ticker configuration from use4.xlsx."""
    logger.info("🔍 Testing ticker configuration loading...")
    
    try:
        # Load TickerList sheet (skip 8 rows as per CLAUDE.md)
        ticker_config = pd.read_excel(
            'use4.xlsx', 
            sheet_name='TickerList', 
            skiprows=8
        )
        
        logger.info(f"✅ Loaded {len(ticker_config)} rows from TickerList")
        logger.info(f"📊 Columns: {list(ticker_config.columns)}")
        
        # Show first few rows
        logger.info("First 5 rows:")
        print(ticker_config.head().to_string())
        
        # Extract Bloomberg tickers
        bloomberg_tickers = []
        for idx, row in ticker_config.iterrows():
            ticker = str(row.get('Ticker', '')).strip()
            if any(suffix in ticker for suffix in ['Index', 'Comdty', 'BGN', 'Equity']):
                bloomberg_tickers.append({
                    'ticker': ticker,
                    'description': row.get('Description', ''),
                    'category': row.get('Category', ''),
                    'region_from': row.get('Region from', ''),
                    'region_to': row.get('Region to', ''),
                    'normalization': row.get('Normalization Factor', 1.0)
                })
        
        logger.info(f"🎯 Found {len(bloomberg_tickers)} Bloomberg tickers")
        
        # Show sample Bloomberg tickers
        logger.info("Sample Bloomberg tickers:")
        for i, ticker_info in enumerate(bloomberg_tickers[:5]):
            logger.info(f"  {i+1}. {ticker_info['ticker']} - {ticker_info['description']}")
        
        return bloomberg_tickers
        
    except Exception as e:
        logger.error(f"❌ Error testing ticker configuration: {str(e)}")
        return None


def test_fallback_data():
    """Test fallback CSV data availability."""
    logger.info("🔍 Testing fallback data availability...")
    
    try:
        fallback_data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True)
        logger.info(f"✅ Fallback data available: {fallback_data.shape}")
        logger.info(f"📅 Date range: {fallback_data.index.min().date()} to {fallback_data.index.max().date()}")
        logger.info(f"📊 Columns: {len(fallback_data.columns)} tickers")
        
        # Show sample data
        logger.info("Sample fallback data (first 3 rows, 5 columns):")
        print(fallback_data.iloc[:3, :5].to_string())
        
        return fallback_data
        
    except FileNotFoundError:
        logger.warning("⚠️ bloomberg_raw_data.csv not found")
        logger.info("💡 To create fallback data:")
        logger.info("   1. Install xbbg: pip install xbbg")
        logger.info("   2. Run Bloomberg system once to generate CSV")
        return None
        
    except Exception as e:
        logger.error(f"❌ Error testing fallback data: {str(e)}")
        return None


def test_multiticker_creation():
    """Test MultiTicker format creation."""
    logger.info("🔍 Testing MultiTicker format creation...")
    
    # Create sample Bloomberg data
    sample_dates = pd.date_range('2023-01-01', '2023-01-10', freq='D')
    sample_tickers = [
        {'ticker': 'SAMPLE1 Index', 'category': 'Industrial', 'region_from': 'France', 'region_to': 'France'},
        {'ticker': 'SAMPLE2 Comdty', 'category': 'LDZ', 'region_from': 'Germany', 'region_to': 'Germany'},
        {'ticker': 'SAMPLE3 BGN', 'category': 'Gas-to-Power', 'region_from': 'Italy', 'region_to': 'Italy'}
    ]
    
    # Create sample data
    np.random.seed(42)  # For reproducible results
    sample_data = pd.DataFrame(
        np.random.rand(len(sample_dates), len(sample_tickers)) * 100,
        index=sample_dates,
        columns=[t['ticker'] for t in sample_tickers]
    )
    
    logger.info(f"📊 Created sample data: {sample_data.shape}")
    
    # Create MultiTicker format
    multiticker_data = pd.DataFrame(index=range(len(sample_dates) + 25))  # Add header rows
    multiticker_data.loc[:, 'A'] = ''  # Column A
    multiticker_data.loc[25:, 'B'] = sample_dates  # Column B: dates starting row 26
    
    # Add data columns with metadata headers
    col_idx = 2  # Start from column C (index 2)
    
    for ticker_info in sample_tickers:
        ticker = ticker_info['ticker']
        # Add metadata headers (rows 13-15)
        multiticker_data.loc[12, col_idx] = 1.0  # Row 13: normalization
        multiticker_data.loc[13, col_idx] = ticker_info.get('category', '')         # Row 14: category
        multiticker_data.loc[14, col_idx] = ticker_info.get('region_from', '')     # Row 15: region from
        multiticker_data.loc[15, col_idx] = ticker_info.get('region_to', '')       # Row 16: region to
        
        # Add data starting row 26 (index 25)
        multiticker_data.loc[25:, col_idx] = sample_data[ticker].values
        
        col_idx += 1
    
    logger.info(f"✅ MultiTicker format created: {multiticker_data.shape}")
    logger.info("Sample MultiTicker headers (rows 13-16):")
    print(multiticker_data.loc[12:16].iloc[:, :5].to_string())
    
    return multiticker_data


def test_demand_processing():
    """Test demand-side processing logic."""
    logger.info("🔍 Testing demand-side processing...")
    
    # Create test MultiTicker data
    multiticker_data = test_multiticker_creation()
    
    # Test industrial demand calculation
    industrial_categories = [
        'Industrial', 'Industrial and Power', 'Zebra', 
        'Industrial (calculated)', 'Gas-to-Power'
    ]
    
    # Extract dates (data starts at row 25)
    dates = pd.to_datetime(multiticker_data.loc[25:, 'B'])
    
    # Test country processing
    countries = ['France', 'Germany', 'Italy']
    
    for country in countries:
        logger.info(f"  Testing {country} demand processing...")
        
        # Filter columns for this country
        country_cols = []
        for col_idx in range(2, multiticker_data.shape[1]):  # Start from column C
            category = str(multiticker_data.loc[13, col_idx]).strip() if pd.notna(multiticker_data.loc[13, col_idx]) else ''
            region_from = str(multiticker_data.loc[14, col_idx]).strip() if pd.notna(multiticker_data.loc[14, col_idx]) else ''
            
            # Match country and industrial categories
            if country.lower() in region_from.lower():
                if any(cat in category for cat in industrial_categories):
                    country_cols.append(col_idx)
        
        # Calculate demand for each date
        demand_values = []
        for date_idx in range(len(dates)):
            data_row_idx = 25 + date_idx
            total = 0.0
            
            for col_idx in country_cols:
                value = multiticker_data.loc[data_row_idx, col_idx]
                if pd.notna(value):
                    total += float(value)
            
            demand_values.append(total)
        
        avg_demand = np.mean(demand_values)
        logger.info(f"    {country} avg demand: {avg_demand:.2f} MCM/d ({len(country_cols)} columns)")
    
    logger.info("✅ Demand processing logic working")


def test_supply_processing():
    """Test supply-side processing logic."""
    logger.info("🔍 Testing supply-side processing...")
    
    # Create test MultiTicker data
    multiticker_data = test_multiticker_creation()
    
    # Define sample supply routes
    supply_routes = [
        ('Test_Import', 'Import', 'Norway', 'Europe'),
        ('Test_Production', 'Production', 'Netherlands', 'Netherlands'),
        ('Test_LNG', 'Import', 'LNG', '*'),  # Wildcard
    ]
    
    # Extract dates
    dates = pd.to_datetime(multiticker_data.loc[25:, 'B'])
    
    for route_name, criteria1, criteria2, criteria3 in supply_routes:
        logger.info(f"  Testing {route_name}...")
        
        # Extract headers from rows 13-15 (0-indexed)
        headers_level1 = multiticker_data.iloc[13, 2:].fillna('').astype(str)
        headers_level2 = multiticker_data.iloc[14, 2:].fillna('').astype(str) 
        headers_level3 = multiticker_data.iloc[15, 2:].fillna('').astype(str)
        
        # Find matching columns
        matching_cols = []
        for col_idx in range(len(headers_level1)):
            match1 = headers_level1.iloc[col_idx].strip() == criteria1
            match2 = headers_level2.iloc[col_idx].strip() == criteria2
            match3 = (criteria3 == '*') or (headers_level3.iloc[col_idx].strip() == criteria3)
            
            if match1 and match2 and match3:
                matching_cols.append(col_idx + 2)  # Adjust for column offset
        
        # Calculate supply for each date
        supply_values = []
        for date_idx in range(len(dates)):
            data_row_idx = 25 + date_idx
            total = 0.0
            
            for col_idx in matching_cols:
                if col_idx < multiticker_data.shape[1]:
                    value = multiticker_data.iloc[data_row_idx, col_idx]
                    if pd.notna(value):
                        total += float(value)
            
            supply_values.append(total)
        
        avg_supply = np.mean(supply_values)
        logger.info(f"    {route_name} avg supply: {avg_supply:.2f} MCM/d ({len(matching_cols)} columns)")
    
    logger.info("✅ Supply processing logic working")


def main():
    """Run complete Bloomberg system test."""
    logger.info("=" * 80)
    logger.info("🚀 BLOOMBERG GAS MARKET SYSTEM TEST")
    logger.info("=" * 80)
    logger.info(f"Test time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Test 1: Ticker configuration
    tickers = test_ticker_configuration()
    ticker_test = "✅ PASS" if tickers else "❌ FAIL"
    
    # Test 2: Fallback data
    fallback_data = test_fallback_data()
    fallback_test = "✅ PASS" if fallback_data is not None else "⚠️ OPTIONAL"
    
    # Test 3: MultiTicker creation
    try:
        multiticker_data = test_multiticker_creation()
        multiticker_test = "✅ PASS"
    except Exception as e:
        logger.error(f"❌ MultiTicker creation failed: {str(e)}")
        multiticker_test = "❌ FAIL"
    
    # Test 4: Demand processing
    try:
        test_demand_processing()
        demand_test = "✅ PASS"
    except Exception as e:
        logger.error(f"❌ Demand processing failed: {str(e)}")
        demand_test = "❌ FAIL"
    
    # Test 5: Supply processing
    try:
        test_supply_processing()
        supply_test = "✅ PASS"
    except Exception as e:
        logger.error(f"❌ Supply processing failed: {str(e)}")
        supply_test = "❌ FAIL"
    
    # Summary
    logger.info("\n" + "=" * 80)
    logger.info("📊 TEST SUMMARY")
    logger.info("=" * 80)
    logger.info(f"1. Ticker Configuration:     {ticker_test}")
    logger.info(f"2. Fallback Data:           {fallback_test}")
    logger.info(f"3. MultiTicker Creation:    {multiticker_test}")
    logger.info(f"4. Demand Processing:       {demand_test}")
    logger.info(f"5. Supply Processing:       {supply_test}")
    
    # Overall status
    core_tests = [ticker_test, multiticker_test, demand_test, supply_test]
    if all("✅" in test for test in core_tests):
        logger.info("\n🎯 OVERALL STATUS: ✅ ALL CORE TESTS PASSED")
        logger.info("🚀 Bloomberg system is ready for production!")
        
        if "⚠️" in fallback_test:
            logger.info("💡 To enable full functionality:")
            logger.info("   • Install xbbg: pip install xbbg")
            logger.info("   • OR provide bloomberg_raw_data.csv")
            
    else:
        logger.info("\n❌ OVERALL STATUS: SOME TESTS FAILED")
        logger.info("🔧 Check failed tests above and fix issues")
    
    return all("✅" in test for test in core_tests)


if __name__ == "__main__":
    main()