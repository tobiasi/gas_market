#!/usr/bin/env python3
"""
Gas-to-Power demand with NORMALIZATION FACTORS from TickerList.
This should fix the over-counting by applying Excel's normalization factors.
"""

import pandas as pd
import numpy as np
from industrial_gas_demand_exact import load_multiticker_with_full_metadata
import logging
import warnings
warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def load_ticker_normalization_mapping(file_path='use4.xlsx'):
    """
    Load TickerList sheet to get normalization factors for each ticker.
    """
    logger.info("Loading TickerList normalization factors")
    
    ticker_list = pd.read_excel(file_path, sheet_name='TickerList', skiprows=8)
    
    # Create mapping: ticker -> normalization factor
    normalization_mapping = {}
    
    for idx, row in ticker_list.iterrows():
        ticker = row.get('Ticker', '')
        norm_factor = row.get('Normalization factor', 1.0)
        category = row.get('Category', '')
        region_from = row.get('Region from', '')
        region_to = row.get('Region to', '')
        
        if pd.notna(norm_factor) and ticker:
            normalization_mapping[ticker] = {
                'factor': float(norm_factor),
                'category': category,
                'region_from': region_from,
                'region_to': region_to
            }
    
    logger.info(f"Loaded normalization factors for {len(normalization_mapping)} tickers")
    
    # Show key normalization factors
    key_factors = {}
    for ticker, info in normalization_mapping.items():
        factor = info['factor']
        if factor not in key_factors:
            key_factors[factor] = []
        key_factors[factor].append(ticker)
    
    logger.info("Key normalization factors found:")
    for factor, tickers in sorted(key_factors.items()):
        logger.info(f"  Factor {factor}: {len(tickers)} tickers")
    
    return normalization_mapping


def map_multiticker_to_tickerlist(metadata, normalization_mapping):
    """
    Map MultiTicker column names to TickerList ticker names.
    This is needed to apply the correct normalization factors.
    """
    logger.info("Mapping MultiTicker columns to TickerList entries")
    
    # Create reverse mapping: MultiTicker column -> TickerList ticker
    column_to_ticker_mapping = {}
    
    # The mapping might be based on position or other logic
    # For now, let's try to match based on patterns
    
    # Debug: show some MultiTicker metadata
    logger.info("Sample MultiTicker metadata:")
    for i, (col, info) in enumerate(list(metadata.items())[:10]):
        logger.info(f"  {col}: {info['category']}/{info['region']}/{info['subcategory']}")
    
    # For Gas-to-Power tickers, we need to find the specific ones causing issues
    gtp_columns = []
    for col, info in metadata.items():
        if (info['category'] == 'Demand' and 
            info['subcategory'] == 'Gas-to-Power'):
            gtp_columns.append((col, info))
    
    logger.info(f"Found {len(gtp_columns)} Gas-to-Power columns in MultiTicker")
    
    # Try to map based on region
    for col, info in gtp_columns:
        region = info['region']
        logger.info(f"  {col}: Region={region}")
        
        # Look for matching TickerList entries
        matching_tickers = []
        for ticker, ticker_info in normalization_mapping.items():
            if (region.lower() in ticker_info['region_from'].lower() or 
                region.lower() in ticker_info['region_to'].lower()):
                if ticker_info['category'] == 'Demand':
                    matching_tickers.append(ticker)
        
        logger.info(f"    Potential matches: {matching_tickers}")
        
        # For now, use the most likely mapping based on the patterns we found
        if region == 'Italy' and matching_tickers:
            # Use SNAMGPGE (Gas Power Generation) as it's most likely for Gas-to-Power
            power_tickers = [t for t in matching_tickers if 'gpge' in t.lower() or 'power' in t.lower()]
            if power_tickers:
                column_to_ticker_mapping[col] = power_tickers[0]
            else:
                column_to_ticker_mapping[col] = matching_tickers[0]  # Use first match
        
        elif region == 'France' and matching_tickers:
            # Use power-related ticker if available
            power_tickers = [t for t in matching_tickers if 'pira' in t.lower() or 'power' in t.lower()]
            if power_tickers:
                column_to_ticker_mapping[col] = power_tickers[0]
            else:
                column_to_ticker_mapping[col] = matching_tickers[0]
    
    logger.info("Column to ticker mapping:")
    for col, ticker in column_to_ticker_mapping.items():
        factor = normalization_mapping[ticker]['factor']
        logger.info(f"  {col} -> {ticker} (factor: {factor})")
    
    return column_to_ticker_mapping


def apply_normalization_factors(data_df, metadata, normalization_mapping, column_mapping):
    """
    Apply normalization factors to the data.
    """
    logger.info("Applying normalization factors to Gas-to-Power data")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Process each Gas-to-Power country with normalization
    countries_data = {}
    
    for col, info in metadata.items():
        if (info['category'] == 'Demand' and 
            info['subcategory'] == 'Gas-to-Power'):
            
            region = info['region']
            raw_data = data_df[col]
            
            # Apply normalization factor if available
            if col in column_mapping:
                ticker = column_mapping[col]
                factor = normalization_mapping[ticker]['factor']
                normalized_data = raw_data * factor
                
                logger.info(f"{region}: Raw mean={raw_data.mean():.2f}, Factor={factor}, Normalized mean={normalized_data.mean():.2f}")
                
                countries_data[f'{region}_GtP'] = normalized_data
            else:
                logger.warning(f"No normalization mapping found for {col} ({region})")
                countries_data[f'{region}_GtP'] = raw_data
    
    # Add all country data to result
    for country_col, data in countries_data.items():
        result[country_col] = data
    
    # Calculate total
    gtp_columns = [col for col in result.columns if col.endswith('_GtP')]
    result['Total_Gas_to_Power_Demand'] = result[gtp_columns].sum(axis=1)
    
    return result


def validate_normalized_results(normalized_df):
    """
    Validate the normalized results against Excel targets.
    """
    logger.info("Validating normalized Gas-to-Power results")
    
    # Target values from Excel Gas-to-power sheet
    targets = {
        'Italy_GtP': 50.91,
        'France_GtP': 12.37,
        'Total_Gas_to_Power_Demand': 141.35
    }
    
    sample = normalized_df[normalized_df['Date'] == '2016-10-03']
    if sample.empty:
        logger.warning("2016-10-03 not found in results")
        return
    
    logger.info("Normalized validation results:")
    for col, target in targets.items():
        if col in sample.columns:
            actual = sample[col].iloc[0]
            diff = abs(actual - target)
            
            if diff < 2.0:  # Allow 2.0 tolerance
                logger.info(f"  ✅ {col}: {actual:.2f} (target {target:.2f}, diff: {diff:.2f})")
            else:
                logger.warning(f"  ⚠️  {col}: {actual:.2f} (target {target:.2f}, diff: {diff:.2f})")
        else:
            logger.warning(f"  ❌ {col}: Column not found")


def main():
    """
    Main execution function with normalization.
    """
    try:
        # Load MultiTicker data
        data_df, metadata = load_multiticker_with_full_metadata('use4.xlsx', 'MultiTicker')
        
        # Load normalization factors
        normalization_mapping = load_ticker_normalization_mapping('use4.xlsx')
        
        # Map MultiTicker columns to TickerList tickers
        column_mapping = map_multiticker_to_tickerlist(metadata, normalization_mapping)
        
        # Apply normalization factors
        normalized_gtp = apply_normalization_factors(data_df, metadata, normalization_mapping, column_mapping)
        
        # Validate results
        validate_normalized_results(normalized_gtp)
        
        # Display sample results
        logger.info(f"\nNormalized Gas-to-Power for 2016-10-03:")
        sample = normalized_gtp[normalized_gtp['Date'] == '2016-10-03']
        if not sample.empty:
            for col in sample.columns:
                if col != 'Date':
                    logger.info(f"  {col}: {sample[col].iloc[0]:.2f}")
        
        # Export results
        export_df = normalized_gtp.copy()
        export_df['Date'] = export_df['Date'].dt.strftime('%Y-%m-%d')
        export_df.round(2).to_csv('gas_to_power_normalized.csv', index=False)
        logger.info("Exported normalized Gas-to-Power data to gas_to_power_normalized.csv")
        
        # Export just the total for Daily historic data
        total_only = export_df[['Date', 'Total_Gas_to_Power_Demand']].copy()
        total_only.to_csv('total_gas_to_power_normalized.csv', index=False)
        logger.info("Exported normalized total to total_gas_to_power_normalized.csv")
        
        return normalized_gtp
        
    except Exception as e:
        logger.error(f"Error in normalized Gas-to-Power execution: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    normalized_data = main()