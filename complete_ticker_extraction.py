#!/usr/bin/env python3
"""
Complete Bloomberg ticker extraction from TickerList sheet in use4.xlsx.
Extract all tickers with full metadata for MultiTicker tab creation.
"""

import pandas as pd
import numpy as np
import logging
import re
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def extract_complete_ticker_list(file_path='use4.xlsx', sheet_name='TickerList'):
    """
    Extract complete Bloomberg ticker list from TickerList sheet with full metadata.
    """
    logger.info(f"Extracting complete ticker list from {sheet_name} in {file_path}")
    
    # Load TickerList sheet - skip first 8 rows as indicated in CLAUDE.md
    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=8)
    
    logger.info(f"Loaded TickerList with shape: {df.shape}")
    logger.info(f"Columns: {list(df.columns)}")
    
    # Display first few rows
    logger.info("First 5 rows of TickerList:")
    print(df.head().to_string())
    
    return df


def process_ticker_metadata(df):
    """
    Process the ticker metadata from the TickerList DataFrame.
    """
    logger.info("Processing ticker metadata")
    
    # Clean up column names
    df.columns = [str(col).strip() for col in df.columns]
    
    # Key columns we expect
    expected_cols = ['Ticker', 'Description', 'Category', 'Region from', 'Region to', 
                    'Units', 'Normalization factor', 'Positive/Negative']
    
    logger.info("Available columns:")
    for i, col in enumerate(df.columns):
        logger.info(f"  {i+1}. {col}")
    
    # Filter out rows with empty tickers
    original_count = len(df)
    df = df.dropna(subset=['Ticker'])
    df = df[df['Ticker'].astype(str).str.strip() != '']
    
    logger.info(f"Filtered from {original_count} to {len(df)} rows with valid tickers")
    
    # Validate Bloomberg tickers
    bloomberg_suffixes = ['Index', 'Comdty', 'BGN', 'Curncy', 'Govt', 'Corp', 'Equity']
    
    def is_bloomberg_ticker(ticker):
        if pd.isna(ticker):
            return False
        ticker_str = str(ticker).strip()
        return any(suffix in ticker_str for suffix in bloomberg_suffixes)
    
    df['Is_Bloomberg_Ticker'] = df['Ticker'].apply(is_bloomberg_ticker)
    bloomberg_df = df[df['Is_Bloomberg_Ticker']].copy()
    
    logger.info(f"Found {len(bloomberg_df)} Bloomberg tickers out of {len(df)} total tickers")
    
    return bloomberg_df


def create_multiticker_structure(ticker_df):
    """
    Create the structure for MultiTicker tab based on ticker metadata.
    """
    logger.info("Creating MultiTicker structure")
    
    # Extract key metadata for MultiTicker format
    multiticker_data = []
    
    for idx, row in ticker_df.iterrows():
        ticker_info = {
            'ticker': str(row['Ticker']).strip(),
            'description': str(row.get('Description', '')).strip() if pd.notna(row.get('Description')) else '',
            'category': str(row.get('Category', '')).strip() if pd.notna(row.get('Category')) else '',
            'region_from': str(row.get('Region from', '')).strip() if pd.notna(row.get('Region from')) else '',
            'region_to': str(row.get('Region to', '')).strip() if pd.notna(row.get('Region to')) else '',
            'units': str(row.get('Units', 'GWh')).strip() if pd.notna(row.get('Units')) else 'GWh',
            'normalization_factor': row.get('Normalization factor', 1.0) if pd.notna(row.get('Normalization factor')) else 1.0,
            'positive_negative': str(row.get('Positive/Negative', '')).strip() if pd.notna(row.get('Positive/Negative')) else '',
            'start_date': str(row.get('Start date', '')).strip() if pd.notna(row.get('Start date')) else '',
            'other_notes': str(row.get('Other notes or comments', '')).strip() if pd.notna(row.get('Other notes or comments')) else ''
        }
        
        multiticker_data.append(ticker_info)
    
    logger.info(f"Created MultiTicker structure for {len(multiticker_data)} tickers")
    
    # Show sample metadata
    logger.info("\nSample ticker metadata:")
    for i, ticker_info in enumerate(multiticker_data[:5]):
        logger.info(f"  {i+1}. {ticker_info['ticker']}: {ticker_info['description']}")
        logger.info(f"     Category: {ticker_info['category']}, Region: {ticker_info['region_from']} -> {ticker_info['region_to']}")
    
    return multiticker_data


def export_complete_ticker_list(multiticker_data, output_file='complete_ticker_list.csv'):
    """
    Export complete ticker list with all metadata.
    """
    logger.info(f"Exporting complete ticker list to {output_file}")
    
    # Convert to DataFrame
    df_export = pd.DataFrame(multiticker_data)
    
    # Round normalization factors
    df_export['normalization_factor'] = pd.to_numeric(df_export['normalization_factor'], errors='coerce').fillna(1.0).round(6)
    
    # Export to CSV
    df_export.to_csv(output_file, index=False)
    
    logger.info(f"Exported {len(df_export)} tickers with complete metadata")
    
    # Show summary statistics
    logger.info("\n📊 TICKER SUMMARY:")
    logger.info(f"  Total tickers: {len(df_export)}")
    
    # Category breakdown
    if 'category' in df_export.columns:
        category_counts = df_export['category'].value_counts()
        logger.info("  Category breakdown:")
        for category, count in category_counts.head(10).items():
            logger.info(f"    {category}: {count}")
    
    # Region breakdown
    if 'region_from' in df_export.columns:
        region_counts = df_export['region_from'].value_counts()
        logger.info("  Region breakdown (top 10):")
        for region, count in region_counts.head(10).items():
            logger.info(f"    {region}: {count}")
    
    return df_export


def main():
    """
    Main execution function for complete ticker extraction.
    """
    try:
        logger.info("🚀 Starting COMPLETE Bloomberg ticker extraction")
        logger.info("=" * 70)
        
        # Step 1: Extract complete ticker list from TickerList sheet
        ticker_df = extract_complete_ticker_list('use4.xlsx', 'TickerList')
        
        # Step 2: Process ticker metadata
        bloomberg_df = process_ticker_metadata(ticker_df)
        
        # Step 3: Create MultiTicker structure
        multiticker_data = create_multiticker_structure(bloomberg_df)
        
        # Step 4: Export complete ticker list
        complete_df = export_complete_ticker_list(multiticker_data)
        
        logger.info("=" * 70)
        logger.info("✅ COMPLETE ticker extraction finished successfully!")
        logger.info(f"📄 Output files:")
        logger.info(f"  - complete_ticker_list.csv: {len(complete_df)} tickers with full metadata")
        
        return complete_df, multiticker_data
        
    except Exception as e:
        logger.error(f"❌ Error in complete ticker extraction: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    result = main()