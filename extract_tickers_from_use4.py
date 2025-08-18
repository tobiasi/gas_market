#!/usr/bin/env python3
"""
Extract Bloomberg ticker list from use4.xlsx file for MultiTicker tab creation.
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


def examine_use4_sheets(file_path='use4.xlsx'):
    """
    Examine all sheets in use4.xlsx to understand structure.
    """
    logger.info(f"Examining sheets in {file_path}")
    
    try:
        # Get all sheet names
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        logger.info(f"Found {len(sheet_names)} sheets:")
        for i, sheet in enumerate(sheet_names):
            logger.info(f"  {i+1}. {sheet}")
        
        return sheet_names
    except Exception as e:
        logger.error(f"Error examining sheets: {str(e)}")
        raise


def extract_tickers_from_use4(file_path='use4.xlsx'):
    """
    Extract Bloomberg ticker list from the 'use4' sheet.
    """
    logger.info("Extracting Bloomberg tickers from use4 sheet")
    
    # First examine available sheets
    sheet_names = examine_use4_sheets(file_path)
    
    # Try to find the sheet with ticker data
    potential_sheets = ['use4', 'TickerList', 'Tickers', 'Bloomberg', 'Data']
    target_sheet = None
    
    for sheet in potential_sheets:
        if sheet in sheet_names:
            target_sheet = sheet
            break
    
    if not target_sheet:
        logger.warning("No obvious ticker sheet found. Trying first few sheets...")
        target_sheet = sheet_names[0]  # Default to first sheet
    
    logger.info(f"Using sheet: {target_sheet}")
    
    # Load the sheet
    df = pd.read_excel(file_path, sheet_name=target_sheet, header=None)
    logger.info(f"Sheet {target_sheet} has shape: {df.shape}")
    
    # Display first few rows to understand structure
    logger.info(f"First 10 rows of {target_sheet}:")
    print(df.head(10).to_string())
    
    return df, target_sheet


def identify_bloomberg_tickers(df):
    """
    Identify Bloomberg tickers in the dataframe.
    Bloomberg tickers typically contain: Index, Comdty, BGN, Curncy, etc.
    """
    logger.info("Identifying Bloomberg ticker patterns")
    
    bloomberg_suffixes = ['Index', 'Comdty', 'BGN', 'Curncy', 'Govt', 'Corp', 'Equity']
    bloomberg_tickers = []
    
    # Search through all cells for Bloomberg tickers
    for row_idx in range(min(50, df.shape[0])):  # Check first 50 rows
        for col_idx in range(min(50, df.shape[1])):  # Check first 50 columns
            cell_value = df.iloc[row_idx, col_idx]
            
            if pd.isna(cell_value):
                continue
            
            cell_str = str(cell_value).strip()
            
            # Check if cell contains Bloomberg ticker pattern
            for suffix in bloomberg_suffixes:
                if suffix in cell_str and len(cell_str) > 3:
                    # Additional validation: should have proper Bloomberg format
                    if ' ' in cell_str and cell_str.endswith(suffix):
                        bloomberg_tickers.append({
                            'ticker': cell_str,
                            'row': row_idx + 1,  # 1-indexed
                            'col': col_idx + 1,  # 1-indexed
                            'suffix': suffix
                        })
                        logger.info(f"Found ticker: {cell_str} at row {row_idx+1}, col {col_idx+1}")
    
    logger.info(f"Found {len(bloomberg_tickers)} Bloomberg tickers")
    
    return bloomberg_tickers


def extract_ticker_metadata(df, ticker_info):
    """
    Extract metadata associated with tickers (descriptions, categories, etc.)
    """
    logger.info("Extracting ticker metadata")
    
    metadata = []
    
    for ticker_data in ticker_info:
        row = ticker_data['row'] - 1  # Convert to 0-indexed
        col = ticker_data['col'] - 1
        
        # Look for metadata in surrounding cells
        metadata_entry = {
            'ticker': ticker_data['ticker'],
            'row': ticker_data['row'],
            'col': ticker_data['col']
        }
        
        # Check adjacent cells for descriptions
        # Above
        if row > 0:
            above = df.iloc[row-1, col]
            if pd.notna(above) and str(above).strip():
                metadata_entry['description_above'] = str(above).strip()
        
        # Below
        if row < df.shape[0] - 1:
            below = df.iloc[row+1, col]
            if pd.notna(below) and str(below).strip():
                metadata_entry['description_below'] = str(below).strip()
        
        # Left
        if col > 0:
            left = df.iloc[row, col-1]
            if pd.notna(left) and str(left).strip():
                metadata_entry['description_left'] = str(left).strip()
        
        # Right
        if col < df.shape[1] - 1:
            right = df.iloc[row, col+1]
            if pd.notna(right) and str(right).strip():
                metadata_entry['description_right'] = str(right).strip()
        
        metadata.append(metadata_entry)
    
    return metadata


def clean_and_validate_tickers(ticker_info):
    """
    Clean and validate Bloomberg tickers.
    """
    logger.info("Cleaning and validating tickers")
    
    cleaned_tickers = []
    seen_tickers = set()
    
    for ticker_data in ticker_info:
        ticker = ticker_data['ticker']
        
        # Basic cleaning
        ticker = ticker.strip()
        ticker = re.sub(r'\s+', ' ', ticker)  # Normalize whitespace
        
        # Skip duplicates
        if ticker in seen_tickers:
            continue
        
        # Basic validation
        if len(ticker) < 5:  # Bloomberg tickers are typically longer
            logger.warning(f"Skipping short ticker: {ticker}")
            continue
        
        # Must contain a space (Bloomberg format: "SYMBOL Type")
        if ' ' not in ticker:
            logger.warning(f"Skipping ticker without space: {ticker}")
            continue
        
        seen_tickers.add(ticker)
        ticker_data['ticker'] = ticker  # Update with cleaned version
        cleaned_tickers.append(ticker_data)
    
    logger.info(f"Cleaned ticker list: {len(cleaned_tickers)} valid tickers")
    return cleaned_tickers


def export_ticker_list(ticker_info, metadata, output_file='ticker_list.csv'):
    """
    Export ticker list and metadata to CSV.
    """
    logger.info(f"Exporting ticker list to {output_file}")
    
    # Create DataFrame for export
    export_data = []
    
    for i, ticker_data in enumerate(ticker_info):
        ticker = ticker_data['ticker']
        
        # Find corresponding metadata
        ticker_metadata = None
        for meta in metadata:
            if meta['ticker'] == ticker:
                ticker_metadata = meta
                break
        
        export_row = {
            'ticker': ticker,
            'suffix': ticker_data['suffix'],
            'row_in_source': ticker_data['row'],
            'col_in_source': ticker_data['col']
        }
        
        # Add metadata if available
        if ticker_metadata:
            for key, value in ticker_metadata.items():
                if key not in ['ticker', 'row', 'col']:
                    export_row[key] = value
        
        export_data.append(export_row)
    
    # Create DataFrame and export
    df_export = pd.DataFrame(export_data)
    df_export.to_csv(output_file, index=False)
    
    logger.info(f"Exported {len(export_data)} tickers to {output_file}")
    return df_export


def main():
    """
    Main execution function to extract Bloomberg tickers from use4.xlsx.
    """
    try:
        logger.info("🚀 Starting Bloomberg ticker extraction from use4.xlsx")
        
        # Step 1: Extract data from use4 sheet
        df, sheet_name = extract_tickers_from_use4('use4.xlsx')
        
        # Step 2: Identify Bloomberg tickers
        ticker_info = identify_bloomberg_tickers(df)
        
        if not ticker_info:
            logger.error("❌ No Bloomberg tickers found in the file")
            return None
        
        # Step 3: Clean and validate tickers
        cleaned_tickers = clean_and_validate_tickers(ticker_info)
        
        # Step 4: Extract metadata
        metadata = extract_ticker_metadata(df, cleaned_tickers)
        
        # Step 5: Export ticker list
        ticker_df = export_ticker_list(cleaned_tickers, metadata)
        
        # Summary
        logger.info("=" * 60)
        logger.info("📊 TICKER EXTRACTION SUMMARY:")
        logger.info(f"  Source sheet: {sheet_name}")
        logger.info(f"  Total tickers found: {len(cleaned_tickers)}")
        
        # Show ticker breakdown by suffix
        suffixes = {}
        for ticker_data in cleaned_tickers:
            suffix = ticker_data['suffix']
            suffixes[suffix] = suffixes.get(suffix, 0) + 1
        
        logger.info("  Ticker breakdown by type:")
        for suffix, count in sorted(suffixes.items()):
            logger.info(f"    {suffix}: {count} tickers")
        
        logger.info("✅ Ticker extraction completed successfully!")
        
        return ticker_df, cleaned_tickers
        
    except Exception as e:
        logger.error(f"❌ Error in ticker extraction: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    result = main()