#!/usr/bin/env python3
"""
Sample Bloomberg Data Generator
==============================

Creates realistic sample Bloomberg data that mimics the actual xbbg API output format.
This allows testing the Bloomberg integration code without requiring actual API access.

The output format matches exactly what xbbg.bdh() returns, including:
- MultiIndex columns (ticker, field)
- DatetimeIndex
- Realistic gas market data patterns
- Proper data types and structure
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def load_ticker_configuration(use4_file='use4.xlsx'):
    """Load ticker configuration from use4.xlsx to get realistic ticker list."""
    logger.info("📊 Loading ticker configuration from use4.xlsx...")
    
    try:
        ticker_config = pd.read_excel(use4_file, sheet_name='TickerList', skiprows=8)
        logger.info(f"✅ Loaded {len(ticker_config)} rows from TickerList")
        
        # Extract Bloomberg tickers
        bloomberg_tickers = []
        for idx, row in ticker_config.iterrows():
            ticker = str(row.get('Ticker', '')).strip()
            if any(suffix in ticker for suffix in ['Index', 'Comdty', 'BGN', 'Equity']):
                bloomberg_tickers.append({
                    'ticker': ticker,
                    'description': str(row.get('Description', '')),
                    'category': str(row.get('Category', '')),
                    'region_from': str(row.get('Region from', '')),
                    'region_to': str(row.get('Region to', '')),
                    'normalization': float(row.get('Normalization factor', 1.0)),
                    'units': str(row.get('Units', ''))
                })
        
        logger.info(f"🎯 Found {len(bloomberg_tickers)} Bloomberg tickers")
        return bloomberg_tickers
        
    except Exception as e:
        logger.error(f"❌ Error loading ticker configuration: {str(e)}")
        return []


def generate_realistic_gas_data(ticker_info, dates):
    """
    Generate realistic gas market data patterns based on ticker category.
    
    This mimics actual European gas market behavior:
    - Seasonal patterns (higher in winter)
    - Daily volatility
    - Category-specific ranges
    - Realistic zero periods for some series
    """
    
    category = ticker_info['category'].lower()
    region_from = ticker_info['region_from'].lower()
    region_to = ticker_info['region_to'].lower()
    
    # Base parameters by category
    if 'demand' in category or 'ldz' in category or 'residential' in category:
        # Demand patterns: higher in winter, seasonal variation
        base_level = np.random.uniform(50, 200)
        seasonal_amplitude = base_level * 0.4
        daily_volatility = 0.1
        zero_probability = 0.02  # Occasional maintenance/outages
        
    elif 'industrial' in category:
        # Industrial: more stable, less seasonal
        base_level = np.random.uniform(30, 150)
        seasonal_amplitude = base_level * 0.2
        daily_volatility = 0.08
        zero_probability = 0.01
        
    elif 'power' in category or 'generation' in category:
        # Power generation: highly variable, weather dependent
        base_level = np.random.uniform(20, 120)
        seasonal_amplitude = base_level * 0.3
        daily_volatility = 0.15
        zero_probability = 0.05
        
    elif 'import' in category or 'pipeline' in category or 'flow' in category:
        # Pipeline flows: can be zero for extended periods, high variability
        base_level = np.random.uniform(0, 250)
        seasonal_amplitude = base_level * 0.5
        daily_volatility = 0.12
        zero_probability = 0.1
        
    elif 'production' in category:
        # Production: relatively stable, some seasonal variation
        base_level = np.random.uniform(20, 180)
        seasonal_amplitude = base_level * 0.3
        daily_volatility = 0.06
        zero_probability = 0.03
        
    elif 'lng' in category:
        # LNG: highly variable, depends on global markets
        base_level = np.random.uniform(0, 100)
        seasonal_amplitude = base_level * 0.6
        daily_volatility = 0.2
        zero_probability = 0.15
        
    else:
        # Default pattern
        base_level = np.random.uniform(10, 100)
        seasonal_amplitude = base_level * 0.3
        daily_volatility = 0.1
        zero_probability = 0.05
    
    # Generate time series
    num_days = len(dates)
    values = []
    
    for i, date in enumerate(dates):
        # Seasonal component (winter higher)
        day_of_year = date.timetuple().tm_yday
        seasonal_factor = 1 + seasonal_amplitude * np.cos(2 * np.pi * (day_of_year - 365/4) / 365)
        
        # Trend component (slight decline over time for some categories)
        if 'production' in category and 'netherlands' in region_from:
            # Dutch production declining
            trend_factor = 1 - (i / num_days) * 0.3
        else:
            trend_factor = 1.0
        
        # Daily volatility
        daily_factor = 1 + np.random.normal(0, daily_volatility)
        
        # Calculate value
        value = base_level * seasonal_factor * trend_factor * daily_factor
        
        # Apply zero probability for maintenance/outages
        if np.random.random() < zero_probability:
            value = 0.0
        
        # Ensure non-negative
        value = max(0, value)
        
        values.append(value)
    
    return np.array(values)


def create_sample_bloomberg_data(start_date='2016-01-01', end_date='2017-12-31', max_tickers=100):
    """
    Create sample Bloomberg data in the exact format returned by xbbg.bdh().
    
    Returns DataFrame with:
    - DatetimeIndex
    - MultiIndex columns: (ticker, field)
    - Field is always 'PX_LAST' (last price)
    - Realistic gas market data patterns
    """
    
    logger.info("🚀 Creating sample Bloomberg data...")
    logger.info(f"📅 Date range: {start_date} to {end_date}")
    
    # Load ticker configuration
    ticker_infos = load_ticker_configuration()
    
    if not ticker_infos:
        logger.error("❌ No tickers found, creating minimal sample")
        ticker_infos = [
            {'ticker': 'SAMPLE1 Index', 'category': 'Demand', 'region_from': 'France', 'region_to': 'France'},
            {'ticker': 'SAMPLE2 Index', 'category': 'Industrial', 'region_from': 'Germany', 'region_to': 'Germany'},
            {'ticker': 'SAMPLE3 Index', 'category': 'LNG', 'region_from': 'LNG', 'region_to': 'Europe'}
        ]
    
    # Limit number of tickers for manageable sample size
    if len(ticker_infos) > max_tickers:
        ticker_infos = ticker_infos[:max_tickers]
        logger.info(f"🔄 Limited to first {max_tickers} tickers for sample")
    
    # Create date range
    dates = pd.date_range(start=start_date, end=end_date, freq='D')
    logger.info(f"📊 Generating {len(dates)} days × {len(ticker_infos)} tickers")
    
    # Create MultiIndex columns (ticker, field) - matches xbbg format
    multi_columns = []
    for ticker_info in ticker_infos:
        multi_columns.append((ticker_info['ticker'], 'PX_LAST'))
    
    multi_index = pd.MultiIndex.from_tuples(multi_columns, names=['ticker', 'field'])
    
    # Initialize DataFrame with MultiIndex columns
    bloomberg_data = pd.DataFrame(index=dates, columns=multi_index)
    
    # Generate realistic data for each ticker
    for i, ticker_info in enumerate(ticker_infos):
        ticker = ticker_info['ticker']
        logger.info(f"  Generating data for {ticker} ({ticker_info['category']})")
        
        # Generate realistic time series
        values = generate_realistic_gas_data(ticker_info, dates)
        
        # Store in DataFrame
        bloomberg_data[(ticker, 'PX_LAST')] = values
        
        # Progress update
        if (i + 1) % 50 == 0:
            logger.info(f"  Progress: {i + 1}/{len(ticker_infos)} tickers completed")
    
    logger.info(f"✅ Sample Bloomberg data created: {bloomberg_data.shape}")
    
    # Show sample statistics
    logger.info("\n📊 SAMPLE DATA STATISTICS:")
    logger.info(f"  Date range: {bloomberg_data.index.min().date()} to {bloomberg_data.index.max().date()}")
    logger.info(f"  Tickers: {len(ticker_infos)}")
    logger.info(f"  Total data points: {bloomberg_data.size:,}")
    
    # Show sample values
    logger.info("\n📋 Sample values (first 5 tickers, first 3 dates):")
    sample_data = bloomberg_data.iloc[:3, :5]
    print(sample_data.round(2).to_string())
    
    return bloomberg_data


def save_sample_data(bloomberg_data, output_file='sample_bloomberg_data.csv'):
    """Save sample Bloomberg data to CSV file."""
    logger.info(f"💾 Saving sample data to {output_file}...")
    
    # Flatten MultiIndex columns for CSV compatibility
    bloomberg_data.columns = [f"{ticker}_{field}" for ticker, field in bloomberg_data.columns]
    
    # Save to CSV
    bloomberg_data.to_csv(output_file)
    logger.info(f"✅ Sample data saved: {output_file}")
    logger.info(f"📊 File size: {bloomberg_data.shape[0]} rows × {bloomberg_data.shape[1]} columns")
    
    return output_file


def create_fallback_file(bloomberg_data, fallback_file='bloomberg_raw_data.csv'):
    """Create fallback file that matches the expected format."""
    logger.info(f"🔄 Creating fallback file: {fallback_file}...")
    
    # Flatten columns if MultiIndex
    if isinstance(bloomberg_data.columns, pd.MultiIndex):
        bloomberg_data.columns = [f"{ticker}_{field}" for ticker, field in bloomberg_data.columns]
    
    # Save as fallback
    bloomberg_data.to_csv(fallback_file)
    logger.info(f"✅ Fallback file created: {fallback_file}")
    
    return fallback_file


def validate_sample_data(bloomberg_data):
    """Validate that sample data looks realistic."""
    logger.info("🔍 Validating sample data quality...")
    
    # Flatten columns for easier analysis
    if isinstance(bloomberg_data.columns, pd.MultiIndex):
        flat_data = bloomberg_data.copy()
        flat_data.columns = [f"{ticker}_{field}" for ticker, field in flat_data.columns]
    else:
        flat_data = bloomberg_data
    
    # Basic statistics
    total_values = flat_data.size
    non_null_values = flat_data.count().sum()
    zero_values = (flat_data == 0).sum().sum()
    
    logger.info(f"  Total values: {total_values:,}")
    logger.info(f"  Non-null values: {non_null_values:,} ({non_null_values/total_values*100:.1f}%)")
    logger.info(f"  Zero values: {zero_values:,} ({zero_values/total_values*100:.1f}%)")
    
    # Range checks
    min_val = flat_data.min().min()
    max_val = flat_data.max().max()
    mean_val = flat_data.mean().mean()
    
    logger.info(f"  Value range: {min_val:.2f} to {max_val:.2f}")
    logger.info(f"  Overall mean: {mean_val:.2f}")
    
    # Quality checks
    quality_issues = []
    
    if min_val < 0:
        quality_issues.append("Negative values found")
    
    if max_val > 1000:
        quality_issues.append("Very high values found (>1000)")
        
    if mean_val == 0:
        quality_issues.append("All values are zero")
    
    if len(quality_issues) == 0:
        logger.info("✅ Sample data quality: GOOD")
    else:
        logger.warning("⚠️ Sample data quality issues:")
        for issue in quality_issues:
            logger.warning(f"    {issue}")
    
    return len(quality_issues) == 0


def main():
    """Create sample Bloomberg data for testing."""
    
    logger.info("=" * 80)
    logger.info("🚀 SAMPLE BLOOMBERG DATA GENERATOR")
    logger.info("=" * 80)
    logger.info(f"Purpose: Create realistic Bloomberg API sample data for testing")
    logger.info(f"Generation time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    try:
        # Create sample Bloomberg data
        bloomberg_data = create_sample_bloomberg_data(
            start_date='2016-01-01',
            end_date='2017-12-31',
            max_tickers=100  # Limit for manageable file size
        )
        
        # Validate data quality
        is_valid = validate_sample_data(bloomberg_data)
        
        # Save files
        sample_file = save_sample_data(bloomberg_data)
        fallback_file = create_fallback_file(bloomberg_data)
        
        # Final summary
        logger.info("\n" + "=" * 80)
        logger.info("✅ SAMPLE DATA GENERATION COMPLETED!")
        logger.info("=" * 80)
        logger.info("📁 Files created:")
        logger.info(f"  • {sample_file} - Full sample with MultiIndex columns")
        logger.info(f"  • {fallback_file} - Fallback file for Bloomberg integration")
        logger.info("\n🧪 Test with Bloomberg integration:")
        logger.info("  python gas_market_bloomberg_chunked.py")
        logger.info("  python run_with_bloomberg_data.py")
        logger.info("\n🎯 Sample data ready for testing Bloomberg integration!")
        logger.info("=" * 80)
        
        return bloomberg_data
        
    except Exception as e:
        logger.error(f"❌ Error creating sample data: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    main()