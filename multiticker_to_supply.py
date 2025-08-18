#!/usr/bin/env python3
"""
Transform MultiTicker sheet data into Supply data aggregation.
Focuses on imports, production, storage, and other supply-side categories.
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
    
    Returns:
        data_df: DataFrame with dates and ticker values (from row 21 onwards)
        metadata: Dictionary with category and region for each ticker column
    """
    logger.info(f"Loading MultiTicker sheet from {file_path}")
    
    # Read the full sheet to get metadata rows
    df_full = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Extract metadata from specific rows (0-indexed: row 14 = index 13, row 15 = index 14)
    categories = df_full.iloc[13, 2:].fillna('')  # Starting from column C (index 2)
    regions = df_full.iloc[14, 2:].fillna('')  # Starting from column C (index 2)
    
    # Data starts from row 21 (index 20)
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
    
    # Create metadata dictionary
    metadata = {}
    for i, col in enumerate(data_rows.columns[1:]):
        metadata[col] = {
            'category': str(categories.iloc[i]) if i < len(categories) else '',
            'region': str(regions.iloc[i]) if i < len(regions) else ''
        }
    
    logger.info(f"Loaded {len(data_rows)} dates with {len(data_rows.columns)-1} tickers")
    
    return data_rows, metadata


def identify_supply_flows(metadata):
    """
    Categorize tickers by supply type.
    
    Returns dictionary with categorized supply flows:
    - pipeline_imports: Imports from pipeline sources
    - lng_imports: LNG imports
    - production: Domestic production
    - storage: Storage operations (Net Withdrawals, Inventory)
    - exports: Export flows
    """
    supply_flows = {
        'pipeline_imports': [],
        'lng_imports': [],
        'production': [],
        'storage_withdrawals': [],
        'storage_inventory': [],
        'exports': [],
        'other_imports': []
    }
    
    # Major pipeline sources
    pipeline_sources = ['Norway', 'Russia', 'Algeria', 'Libya', 'Azerbaijan', 
                       'Russia (Nord Stream)', 'MAB', 'Slovenia']
    
    for col, info in metadata.items():
        category = info['category']
        region = info['region']
        
        if category == 'Import':
            if region == 'LNG':
                supply_flows['lng_imports'].append(col)
            elif any(source in region for source in pipeline_sources):
                supply_flows['pipeline_imports'].append(col)
            else:
                supply_flows['other_imports'].append(col)
        
        elif category == 'Production':
            supply_flows['production'].append(col)
        
        elif category == 'Net Withdrawals':
            supply_flows['storage_withdrawals'].append(col)
        
        elif category == 'Inventory':
            supply_flows['storage_inventory'].append(col)
        
        elif category == 'Export':
            supply_flows['exports'].append(col)
    
    # Log summary
    for flow_type, cols in supply_flows.items():
        if cols:
            logger.info(f"{flow_type}: {len(cols)} columns")
    
    return supply_flows


def aggregate_pipeline_imports(data_df, metadata, source_country=None):
    """
    Aggregate pipeline imports from specific source or all sources.
    
    Args:
        source_country: Specific source like 'Norway', 'Russia', or None for all
    
    Returns:
        Series with aggregated pipeline imports for each date
    """
    matching_cols = []
    
    for col, info in metadata.items():
        if info['category'] == 'Import':
            if source_country:
                # Look for specific source
                if source_country.lower() in info['region'].lower():
                    matching_cols.append(col)
            else:
                # Include all non-LNG imports
                if info['region'] != 'LNG':
                    matching_cols.append(col)
    
    if not matching_cols:
        logger.debug(f"No pipeline import columns found for {source_country or 'all sources'}")
        return pd.Series(0, index=data_df.index)
    
    logger.debug(f"Found {len(matching_cols)} pipeline import columns for {source_country or 'all sources'}")
    
    # Sum across matching columns
    result = data_df[matching_cols].sum(axis=1, skipna=True)
    
    return result


def aggregate_lng_imports(data_df, metadata):
    """
    Aggregate total LNG imports.
    
    Returns:
        Series with aggregated LNG imports for each date
    """
    matching_cols = []
    
    for col, info in metadata.items():
        if info['category'] == 'Import' and info['region'] == 'LNG':
            matching_cols.append(col)
    
    if not matching_cols:
        logger.debug("No LNG import columns found")
        return pd.Series(0, index=data_df.index)
    
    logger.info(f"Found {len(matching_cols)} LNG import columns")
    
    # Sum across matching columns
    result = data_df[matching_cols].sum(axis=1, skipna=True)
    
    return result


def aggregate_domestic_production(data_df, metadata, country=None):
    """
    Aggregate domestic gas production by country or total.
    
    Args:
        country: Specific country or None for total European production
    
    Returns:
        Series with aggregated production for each date
    """
    matching_cols = []
    
    for col, info in metadata.items():
        if info['category'] == 'Production':
            if country:
                if country.lower() in info['region'].lower():
                    matching_cols.append(col)
            else:
                matching_cols.append(col)
    
    if not matching_cols:
        logger.debug(f"No production columns found for {country or 'total'}")
        return pd.Series(0, index=data_df.index)
    
    logger.debug(f"Found {len(matching_cols)} production columns for {country or 'total'}")
    
    # Sum across matching columns
    result = data_df[matching_cols].sum(axis=1, skipna=True)
    
    return result


def aggregate_storage_flows(data_df, metadata, flow_type='net_withdrawals', country=None):
    """
    Aggregate storage operations.
    
    Args:
        flow_type: 'net_withdrawals' or 'inventory'
        country: Specific country or None for total
    
    Returns:
        Series with aggregated storage flows for each date
    """
    category_map = {
        'net_withdrawals': 'Net Withdrawals',
        'inventory': 'Inventory'
    }
    
    target_category = category_map.get(flow_type, 'Net Withdrawals')
    matching_cols = []
    
    for col, info in metadata.items():
        if info['category'] == target_category:
            if country:
                if country.lower() in info['region'].lower():
                    matching_cols.append(col)
            else:
                matching_cols.append(col)
    
    if not matching_cols:
        logger.debug(f"No {flow_type} columns found for {country or 'total'}")
        return pd.Series(0, index=data_df.index)
    
    logger.debug(f"Found {len(matching_cols)} {flow_type} columns for {country or 'total'}")
    
    # Sum across matching columns
    result = data_df[matching_cols].sum(axis=1, skipna=True)
    
    return result


def calculate_total_supply_by_source(data_df, metadata):
    """
    Create supply breakdown by major sources.
    
    Returns DataFrame with columns:
    - Date
    - Norway_Pipeline
    - Russia_Pipeline  
    - Algeria_Pipeline
    - LNG_Total
    - Domestic_Production
    - Storage_Net_Withdrawals
    - Other_Imports
    - Total_Supply
    """
    logger.info("Creating supply breakdown by major sources")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Pipeline imports by source
    result['Norway_Pipeline'] = aggregate_pipeline_imports(data_df, metadata, 'Norway')
    result['Russia_Pipeline'] = aggregate_pipeline_imports(data_df, metadata, 'Russia')
    result['Algeria_Pipeline'] = aggregate_pipeline_imports(data_df, metadata, 'Algeria')
    
    # LNG imports
    result['LNG_Total'] = aggregate_lng_imports(data_df, metadata)
    
    # Domestic production
    result['Domestic_Production'] = aggregate_domestic_production(data_df, metadata)
    
    # Storage flows
    result['Storage_Net_Withdrawals'] = aggregate_storage_flows(data_df, metadata, 'net_withdrawals')
    
    # Other imports (not covered above)
    all_pipeline = aggregate_pipeline_imports(data_df, metadata)
    major_pipeline = result['Norway_Pipeline'] + result['Russia_Pipeline'] + result['Algeria_Pipeline']
    result['Other_Imports'] = all_pipeline - major_pipeline
    
    # Total supply
    supply_cols = ['Norway_Pipeline', 'Russia_Pipeline', 'Algeria_Pipeline', 
                   'LNG_Total', 'Domestic_Production', 'Storage_Net_Withdrawals', 'Other_Imports']
    result['Total_Supply'] = result[supply_cols].sum(axis=1)
    
    # Sort by date
    result = result.sort_values('Date').reset_index(drop=True)
    
    logger.info(f"Created supply breakdown with {len(result)} rows")
    
    return result


def calculate_country_supply_balance(data_df, metadata, country):
    """
    Calculate complete supply balance for specific country.
    
    Returns DataFrame with country-specific supply breakdown
    """
    logger.info(f"Creating supply balance for {country}")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    result['Country'] = country
    
    # Country-specific production
    result['Domestic_Production'] = aggregate_domestic_production(data_df, metadata, country)
    
    # Country-specific storage
    result['Storage_Withdrawals'] = aggregate_storage_flows(data_df, metadata, 'net_withdrawals', country)
    
    # Imports to this country (need to check region patterns)
    import_cols = []
    for col, info in metadata.items():
        if info['category'] == 'Import':
            # Check if this import is destined for the country
            # This requires understanding the region naming convention
            if country.lower() in info['region'].lower():
                import_cols.append(col)
    
    if import_cols:
        result['Total_Imports'] = data_df[import_cols].sum(axis=1, skipna=True)
    else:
        result['Total_Imports'] = 0
    
    # Total supply for country
    result['Total_Supply'] = (result['Domestic_Production'] + 
                              result['Storage_Withdrawals'] + 
                              result['Total_Imports'])
    
    # Sort by date
    result = result.sort_values('Date').reset_index(drop=True)
    
    return result


def validate_supply_demand_balance(supply_data, demand_data):
    """
    Verify that total supply approximately equals total demand.
    
    Args:
        supply_data: DataFrame with supply totals
        demand_data: DataFrame with demand totals (from Task 1)
    
    Returns:
        DataFrame with balance comparison
    """
    # Merge on date
    balance = pd.merge(
        supply_data[['Date', 'Total_Supply']],
        demand_data[['Date', 'Total']],
        on='Date',
        how='inner',
        suffixes=('_supply', '_demand')
    )
    
    balance.rename(columns={'Total': 'Total_Demand'}, inplace=True)
    
    # Calculate imbalance
    balance['Imbalance'] = balance['Total_Supply'] - balance['Total_Demand']
    balance['Imbalance_Pct'] = (balance['Imbalance'] / balance['Total_Demand']) * 100
    
    # Log summary statistics
    logger.info(f"Supply-Demand Balance Statistics:")
    logger.info(f"  Mean imbalance: {balance['Imbalance'].mean():.2f}")
    logger.info(f"  Mean imbalance %: {balance['Imbalance_Pct'].mean():.2f}%")
    logger.info(f"  Max absolute imbalance: {balance['Imbalance'].abs().max():.2f}")
    
    return balance


def create_complete_gas_balance(supply_data, demand_data):
    """
    Merge supply and demand data for complete European gas balance.
    
    Returns DataFrame with complete gas balance
    """
    # Merge supply and demand data
    balance = pd.merge(
        supply_data,
        demand_data,
        on='Date',
        how='outer',
        suffixes=('_supply', '_demand')
    )
    
    # Calculate net balance
    if 'Total_Supply' in balance.columns and 'Total' in balance.columns:
        balance['Net_Balance'] = balance['Total_Supply'] - balance['Total']
        balance['Storage_Implied'] = -balance['Net_Balance']  # Negative balance implies storage injection
    
    return balance


def export_supply_data(supply_by_source, output_file='european_supply_balance.csv'):
    """
    Export supply data to CSV file.
    """
    # Format date column
    supply_by_source['Date'] = supply_by_source['Date'].dt.strftime('%Y-%m-%d')
    
    # Round numeric columns to 2 decimal places
    numeric_cols = supply_by_source.select_dtypes(include=[np.number]).columns
    supply_by_source[numeric_cols] = supply_by_source[numeric_cols].round(2)
    
    # Export to CSV
    supply_by_source.to_csv(output_file, index=False)
    logger.info(f"Exported supply balance to {output_file}")
    
    return output_file


def main():
    """
    Main execution function.
    """
    try:
        # Load MultiTicker data from Excel
        data_df, metadata = load_multiticker_from_excel('use4.xlsx', 'MultiTicker')
        
        # Identify supply flows
        supply_flows = identify_supply_flows(metadata)
        
        # Calculate total supply by source
        supply_by_source = calculate_total_supply_by_source(data_df, metadata)
        
        # Display sample results
        logger.info("\nSample supply data (first 5 rows):")
        print(supply_by_source.head())
        
        # Display specific date for validation
        logger.info("\nSupply data for 2016-10-03:")
        sample = supply_by_source[supply_by_source['Date'] == '2016-10-03']
        if not sample.empty:
            for col in supply_by_source.columns:
                if col != 'Date':
                    print(f"  {col}: {sample[col].iloc[0]:.2f}")
        
        # Export to CSV
        output_file = export_supply_data(supply_by_source, 'european_supply_balance.csv')
        
        # Also calculate country-specific supply (example: Germany)
        germany_supply = calculate_country_supply_balance(data_df, metadata, 'Germany')
        logger.info("\nGermany supply balance (first 5 rows):")
        print(germany_supply.head())
        
        return supply_by_source
        
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        raise


if __name__ == "__main__":
    supply_data = main()