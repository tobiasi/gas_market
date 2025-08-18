#!/usr/bin/env python3
"""
Transform MultiTicker sheet data into Daily historic data by category format.
Replicates Excel SUMIFS logic for aggregating Bloomberg time series data.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import logging
import warnings
warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def load_multiticker_from_excel(file_path='use4.xlsx', sheet_name='MultiTicker'):
    """
    Load MultiTicker data directly from Excel file.
    Now also includes subcategory data from row 16 for proper Industrial/LDZ/GtP calculations.
    
    Returns:
        data_df: DataFrame with dates and ticker values (from row 21 onwards)
        metadata: Dictionary with category, region, and subcategory for each ticker
    """
    logger.info(f"Loading MultiTicker sheet from {file_path}")
    
    try:
        # Try to use the enhanced loader with subcategories
        from multiticker_industrial_demand import load_multiticker_with_subcategories
        data_df, metadata = load_multiticker_with_subcategories(file_path, sheet_name)
        logger.info("Successfully loaded with subcategory support")
        return data_df, metadata
        
    except Exception as e:
        logger.warning(f"Could not load with subcategories: {e}, using basic loader")
        
        # Fallback to basic loader
        # Read the full sheet to get metadata rows
        df_full = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # Extract metadata from specific rows (0-indexed: row 14 = index 13, row 15 = index 14)
        # Row 14 contains categories
        categories = df_full.iloc[13, 2:].fillna('')  # Starting from column C (index 2)
        
        # Row 15 contains regions
        regions = df_full.iloc[14, 2:].fillna('')  # Starting from column C (index 2)
        
        # Data starts from row 21 (index 20)
        # Column B contains dates, columns C onwards contain data
        data_rows = df_full.iloc[20:, 1:].copy()  # Column B is index 1
        
        # Set column names
        data_rows.columns = ['Date'] + [f'Col_{i}' for i in range(len(data_rows.columns)-1)]
        
        # Convert first column to datetime
        data_rows['Date'] = pd.to_datetime(data_rows['Date'], errors='coerce')
        
        # Remove rows with invalid dates
        data_rows = data_rows.dropna(subset=['Date'])
        
        # Convert data columns to numeric
        for col in data_rows.columns[1:]:
            data_rows[col] = pd.to_numeric(data_rows[col], errors='coerce')
        
        # Create metadata dictionary (basic version without subcategories)
        metadata = {}
        for i, col in enumerate(data_rows.columns[1:]):
            metadata[col] = {
                'category': str(categories.iloc[i]) if i < len(categories) else '',
                'region': str(regions.iloc[i]) if i < len(regions) else '',
                'subcategory': ''  # Empty for fallback mode
            }
        
        logger.info(f"Loaded {len(data_rows)} dates with {len(data_rows.columns)-1} tickers (basic mode)")
        
        return data_rows, metadata


def aggregate_by_category_region(data_df, metadata, target_category, target_region):
    """
    Replicate Excel SUMIFS logic to aggregate data by category and region.
    
    Args:
        data_df: DataFrame with Date column and ticker columns
        metadata: Dictionary with category and region for each ticker column
        target_category: Category to filter (e.g., 'Demand')
        target_region: Region to filter (e.g., 'France')
    
    Returns:
        Series with aggregated values for each date
    """
    # Find columns matching both category and region
    matching_cols = []
    for col, info in metadata.items():
        if info['category'] == target_category and info['region'] == target_region:
            matching_cols.append(col)
    
    if not matching_cols:
        logger.debug(f"No columns found for {target_category}/{target_region}")
        return pd.Series(0, index=data_df.index)
    
    logger.debug(f"Found {len(matching_cols)} columns for {target_category}/{target_region}")
    
    # Sum across matching columns, handling NaN values
    result = data_df[matching_cols].sum(axis=1, skipna=True)
    
    return result


def calculate_subcategory_totals(data_df, metadata):
    """
    Calculate Industrial, LDZ, and Gas-to-Power totals using 3-criteria SUMIFS logic.
    """
    # Check if we have subcategory support
    has_subcategories = any(m.get('subcategory', '') != '' for m in metadata.values())
    
    if has_subcategories:
        logger.info("Calculating subcategories using 3-criteria SUMIFS")
        
        try:
            from multiticker_industrial_demand import create_subcategory_totals
            subcategory_df, _ = create_subcategory_totals(data_df, metadata)
            
            totals = {
                'Industrial': subcategory_df['Industrial'],
                'LDZ': subcategory_df['LDZ'],
                'Gas_to_Power': subcategory_df['Gas_to_Power']
            }
            
            return totals
            
        except Exception as e:
            logger.warning(f"3-criteria calculation failed: {e}")
    
    # Fallback: try to load from pre-calculated files
    try:
        logger.info("Loading subcategories from pre-calculated file")
        subcategory_df = pd.read_csv('subcategory_totals.csv')
        subcategory_df['Date'] = pd.to_datetime(subcategory_df['Date'])
        
        # Merge with our data on dates
        merged = pd.merge(data_df[['Date']], subcategory_df, on='Date', how='left')
        
        totals = {
            'Industrial': merged['Industrial'].fillna(0),
            'LDZ': merged['LDZ'].fillna(0),
            'Gas_to_Power': merged['Gas_to_Power'].fillna(0)
        }
        
        return totals
        
    except Exception as e:
        logger.warning(f"Could not load pre-calculated subcategories: {e}")
        
        # Final fallback - zeros
        logger.warning("Using zero values for subcategories - run multiticker_industrial_demand.py first")
        totals = {
            'Industrial': pd.Series(0, index=data_df.index, dtype=float),
            'LDZ': pd.Series(0, index=data_df.index, dtype=float),
            'Gas_to_Power': pd.Series(0, index=data_df.index, dtype=float)
        }
        
        return totals


def create_daily_historic_data(data_df, metadata):
    """
    Create the complete daily historic data table matching Excel output.
    
    Columns match Excel exactly:
    - Date
    - Country demand columns (France, Belgium, Italy, etc.)
    - Total (sum of country demands)
    - Industrial, LDZ, Gas-to-Power subcategory totals
    """
    logger.info("Creating daily historic data aggregation")
    
    # Define countries in exact Excel order
    countries = [
        'France', 'Belgium', 'Italy', 'Netherlands', 'GB', 
        'Austria', 'Germany', 'Switzerland', 'Luxembourg', 
        'Island of Ireland'
    ]
    
    # Alternative names that might appear in the data
    # Note: Region names must match exactly as they appear in row 15 of MultiTicker
    country_aliases = {
        'GB': ['GB', 'Great Britain', 'United Kingdom', 'UK'],
        'Island of Ireland': ['Island of Ireland', 'Ireland', 'IE'],
        'Netherlands': ['Netherlands', 'NL', 'Holland']
    }
    
    # Create result dataframe
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Calculate demand for each country
    country_totals = []
    for country in countries:
        # Check for aliases
        aliases = country_aliases.get(country, [country])
        
        # Try each alias until we find data
        country_data = pd.Series(0, index=data_df.index, dtype=float)
        for alias in aliases:
            # Try different demand category variations
            demand_categories = ['Demand', 'Demand (Net)']
            for demand_cat in demand_categories:
                temp_data = aggregate_by_category_region(data_df, metadata, demand_cat, alias)
                if temp_data.sum() > 0:
                    country_data = temp_data
                    logger.info(f"Found data for {country} using alias '{alias}' with category '{demand_cat}'")
                    break
            if country_data.sum() > 0:
                break
        
        result[country] = country_data
        country_totals.append(country_data)
    
    # Calculate total demand
    result['Total'] = sum(country_totals)
    
    # Calculate subcategory totals
    subcategory_totals = calculate_subcategory_totals(data_df, metadata)
    result['Industrial'] = subcategory_totals['Industrial']
    result['LDZ'] = subcategory_totals['LDZ']
    result['Gas_to_Power'] = subcategory_totals['Gas_to_Power']
    
    # Sort by date
    result = result.sort_values('Date').reset_index(drop=True)
    
    logger.info(f"Created daily historic data with {len(result)} rows")
    
    return result


def validate_results(daily_data, validation_dates=None):
    """
    Validate results against known Excel values.
    
    Default validation for known dates:
    - 2016-10-01: France = 68.46, Total = 613.56
    - 2016-10-03: France = 90.13, Belgium = 38.16, Total = 715.22
    """
    if validation_dates is None:
        validation_dates = [
            {'date': '2016-10-01', 'France': 68.46, 'Total': 613.56},
            {'date': '2016-10-03', 'France': 90.13, 'Belgium': 38.16, 'Total': 715.22}
        ]
    
    logger.info("Validating results against known values")
    
    for val in validation_dates:
        date_str = val['date']
        date = pd.to_datetime(date_str)
        
        # Find row for this date
        row = daily_data[daily_data['Date'] == date]
        
        if row.empty:
            logger.warning(f"Date {date_str} not found in results")
            continue
        
        # Check each value
        for col, expected in val.items():
            if col == 'date':
                continue
            
            actual = row[col].iloc[0]
            diff = abs(actual - expected)
            
            if diff < 0.01:
                logger.info(f"✓ {date_str} {col}: {actual:.2f} (expected {expected:.2f})")
            else:
                logger.warning(f"✗ {date_str} {col}: {actual:.2f} (expected {expected:.2f}, diff: {diff:.2f})")


def export_to_csv(daily_data, output_file='daily_historic_data_output.csv'):
    """
    Export daily historic data to CSV file.
    """
    # Format date column
    daily_data['Date'] = daily_data['Date'].dt.strftime('%Y-%m-%d')
    
    # Round numeric columns to 2 decimal places
    numeric_cols = daily_data.select_dtypes(include=[np.number]).columns
    daily_data[numeric_cols] = daily_data[numeric_cols].round(2)
    
    # Export to CSV
    daily_data.to_csv(output_file, index=False)
    logger.info(f"Exported results to {output_file}")
    
    return output_file


def main():
    """
    Main execution function.
    """
    try:
        # Load MultiTicker data from Excel
        data_df, metadata = load_multiticker_from_excel('use4.xlsx', 'MultiTicker')
        
        # Create daily historic data aggregation
        daily_data = create_daily_historic_data(data_df, metadata)
        
        # Validate against known values
        validate_results(daily_data)
        
        # Export to CSV
        output_file = export_to_csv(daily_data, 'daily_historic_data_by_category_output.csv')
        
        # Display sample results
        logger.info("\nSample output (first 5 rows):")
        print(daily_data.head())
        
        logger.info("\nSample output (dates around 2016-10-01):")
        sample_dates = daily_data[
            (daily_data['Date'] >= '2016-10-01') & 
            (daily_data['Date'] <= '2016-10-05')
        ]
        if not sample_dates.empty:
            print(sample_dates[['Date', 'France', 'Belgium', 'Italy', 'Total']])
        
        return daily_data
        
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        raise


if __name__ == "__main__":
    daily_data = main()