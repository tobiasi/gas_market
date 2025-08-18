#!/usr/bin/env python3
"""
Complete European Gas Balance: Integrate supply and demand data from MultiTicker.
Combines demand aggregation (Task 1) and supply aggregation (Task 3).
"""

import pandas as pd
import numpy as np
from datetime import datetime
import logging
import warnings
warnings.filterwarnings('ignore')

# Import the demand and supply modules
from multiticker_to_daily_historic import (
    load_multiticker_from_excel as load_demand_data,
    create_daily_historic_data
)
from multiticker_to_supply import (
    calculate_total_supply_by_source,
    identify_supply_flows
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def create_complete_gas_balance(file_path='use4.xlsx'):
    """
    Create complete European gas balance combining supply and demand.
    
    Returns:
        DataFrame with complete gas balance including:
        - Demand by country
        - Supply by source
        - Net balance
        - Storage implications
    """
    logger.info("Creating complete European gas balance")
    
    # Load MultiTicker data
    data_df, metadata = load_demand_data(file_path, 'MultiTicker')
    
    # Create demand aggregation
    logger.info("Processing demand data...")
    demand_data = create_daily_historic_data(data_df, metadata)
    
    # Create supply aggregation
    logger.info("Processing supply data...")
    supply_data = calculate_total_supply_by_source(data_df, metadata)
    
    # Merge supply and demand on date
    balance = pd.merge(
        demand_data,
        supply_data,
        on='Date',
        how='inner',
        suffixes=('_demand', '_supply')
    )
    
    # Calculate key balance metrics
    # The demand total column is named 'Total' not 'Total_demand'
    balance['Net_Balance'] = balance['Total_Supply'] - balance['Total']
    balance['Balance_Pct'] = (balance['Net_Balance'] / balance['Total']) * 100
    
    # Storage implications
    # Positive net balance = excess supply (storage injection expected)
    # Negative net balance = supply deficit (storage withdrawal expected)
    balance['Storage_Implied'] = -balance['Net_Balance']
    
    logger.info(f"Created complete balance for {len(balance)} dates")
    
    return balance


def analyze_balance_statistics(balance_df):
    """
    Analyze and report key statistics from the gas balance.
    
    Args:
        balance_df: Complete gas balance DataFrame
    
    Returns:
        Dictionary with key statistics
    """
    stats = {}
    
    # Demand statistics
    demand_cols = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 
                   'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    for col in demand_cols:
        if col in balance_df.columns:
            stats[f'{col}_demand_mean'] = balance_df[col].mean()
            stats[f'{col}_demand_max'] = balance_df[col].max()
    
    # Supply statistics
    supply_sources = ['Norway_Pipeline', 'Russia_Pipeline', 'Algeria_Pipeline', 
                     'LNG_Total', 'Domestic_Production', 'Storage_Net_Withdrawals']
    
    for col in supply_sources:
        if col in balance_df.columns:
            stats[f'{col}_mean'] = balance_df[col].mean()
            stats[f'{col}_max'] = balance_df[col].max()
    
    # Balance statistics
    stats['net_balance_mean'] = balance_df['Net_Balance'].mean()
    stats['net_balance_std'] = balance_df['Net_Balance'].std()
    stats['balance_pct_mean'] = balance_df['Balance_Pct'].mean()
    stats['balance_pct_abs_max'] = balance_df['Balance_Pct'].abs().max()
    
    return stats


def create_seasonal_analysis(balance_df):
    """
    Create seasonal analysis of supply and demand patterns.
    
    Args:
        balance_df: Complete gas balance DataFrame
    
    Returns:
        DataFrame with seasonal aggregations
    """
    # Add month and season columns
    balance_df['Month'] = pd.to_datetime(balance_df['Date']).dt.month
    balance_df['Year'] = pd.to_datetime(balance_df['Date']).dt.year
    
    # Define seasons
    def get_season(month):
        if month in [12, 1, 2]:
            return 'Winter'
        elif month in [3, 4, 5]:
            return 'Spring'
        elif month in [6, 7, 8]:
            return 'Summer'
        else:
            return 'Autumn'
    
    balance_df['Season'] = balance_df['Month'].apply(get_season)
    
    # Aggregate by season
    seasonal = balance_df.groupby(['Year', 'Season']).agg({
        'Total': 'mean',  # Demand total column
        'Total_Supply': 'mean',
        'Norway_Pipeline': 'mean',
        'Russia_Pipeline': 'mean',
        'LNG_Total': 'mean',
        'Storage_Net_Withdrawals': 'mean',
        'Net_Balance': 'mean'
    }).round(2)
    
    return seasonal


def export_complete_balance(balance_df, output_file='european_gas_complete_balance.csv'):
    """
    Export complete balance to CSV file.
    """
    # Format date column
    balance_df = balance_df.copy()
    balance_df['Date'] = pd.to_datetime(balance_df['Date']).dt.strftime('%Y-%m-%d')
    
    # Select key columns for export
    export_cols = [
        'Date',
        # Demand columns
        'France', 'Belgium', 'Italy', 'Netherlands', 'GB', 
        'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland',
        'Total',  # Demand total column
        # Supply columns
        'Norway_Pipeline', 'Russia_Pipeline', 'Algeria_Pipeline',
        'LNG_Total', 'Domestic_Production', 'Storage_Net_Withdrawals',
        'Total_Supply',
        # Balance columns
        'Net_Balance', 'Balance_Pct', 'Storage_Implied'
    ]
    
    # Filter to available columns
    export_cols = [col for col in export_cols if col in balance_df.columns]
    
    # Round numeric columns
    numeric_cols = balance_df[export_cols].select_dtypes(include=[np.number]).columns
    balance_df[numeric_cols] = balance_df[numeric_cols].round(2)
    
    # Export
    balance_df[export_cols].to_csv(output_file, index=False)
    logger.info(f"Exported complete balance to {output_file}")
    
    return output_file


def create_balance_summary(balance_df, date='2016-10-03'):
    """
    Create a detailed summary for a specific date.
    
    Args:
        balance_df: Complete gas balance DataFrame
        date: Date to summarize
    
    Returns:
        Dictionary with detailed breakdown for the date
    """
    # Convert date to datetime
    target_date = pd.to_datetime(date)
    
    # Find row for this date
    row = balance_df[pd.to_datetime(balance_df['Date']) == target_date]
    
    if row.empty:
        logger.warning(f"Date {date} not found in balance data")
        return None
    
    row = row.iloc[0]
    
    summary = {
        'date': date,
        'demand': {
            'total': row.get('Total', 0),
            'countries': {
                'France': row.get('France', 0),
                'Belgium': row.get('Belgium', 0),
                'Italy': row.get('Italy', 0),
                'Netherlands': row.get('Netherlands', 0),
                'GB': row.get('GB', 0),
                'Austria': row.get('Austria', 0),
                'Germany': row.get('Germany', 0),
                'Switzerland': row.get('Switzerland', 0),
                'Luxembourg': row.get('Luxembourg', 0),
                'Ireland': row.get('Island of Ireland', 0)
            }
        },
        'supply': {
            'total': row.get('Total_Supply', 0),
            'sources': {
                'Norway_Pipeline': row.get('Norway_Pipeline', 0),
                'Russia_Pipeline': row.get('Russia_Pipeline', 0),
                'Algeria_Pipeline': row.get('Algeria_Pipeline', 0),
                'LNG_Total': row.get('LNG_Total', 0),
                'Domestic_Production': row.get('Domestic_Production', 0),
                'Storage_Net_Withdrawals': row.get('Storage_Net_Withdrawals', 0),
                'Other_Imports': row.get('Other_Imports', 0)
            }
        },
        'balance': {
            'net_balance': row.get('Net_Balance', 0),
            'balance_pct': row.get('Balance_Pct', 0),
            'storage_implied': row.get('Storage_Implied', 0)
        }
    }
    
    return summary


def main():
    """
    Main execution function.
    """
    try:
        # Create complete gas balance
        balance = create_complete_gas_balance('use4.xlsx')
        
        # Analyze statistics
        stats = analyze_balance_statistics(balance)
        
        # Display key statistics
        logger.info("\n=== European Gas Balance Statistics ===")
        logger.info(f"Average daily demand: {stats.get('net_balance_mean', 0):.2f}")
        logger.info(f"Average balance %: {stats.get('balance_pct_mean', 0):.2f}%")
        logger.info(f"Max absolute balance %: {stats.get('balance_pct_abs_max', 0):.2f}%")
        
        # Create summary for specific date
        summary = create_balance_summary(balance, '2016-10-03')
        if summary:
            logger.info("\n=== Balance Summary for 2016-10-03 ===")
            logger.info(f"Total Demand: {summary['demand']['total']:.2f}")
            logger.info(f"Total Supply: {summary['supply']['total']:.2f}")
            logger.info(f"Net Balance: {summary['balance']['net_balance']:.2f}")
            logger.info(f"Balance %: {summary['balance']['balance_pct']:.2f}%")
            
            logger.info("\nTop demand countries:")
            top_demand = sorted(summary['demand']['countries'].items(), 
                              key=lambda x: x[1], reverse=True)[:3]
            for country, value in top_demand:
                logger.info(f"  {country}: {value:.2f}")
            
            logger.info("\nTop supply sources:")
            top_supply = sorted(summary['supply']['sources'].items(), 
                              key=lambda x: x[1], reverse=True)[:3]
            for source, value in top_supply:
                logger.info(f"  {source}: {value:.2f}")
        
        # Create seasonal analysis
        seasonal = create_seasonal_analysis(balance)
        logger.info("\n=== Seasonal Analysis (Sample) ===")
        print(seasonal.head(8))
        
        # Export complete balance
        output_file = export_complete_balance(balance)
        
        # Also export seasonal analysis
        seasonal.to_csv('european_gas_seasonal_analysis.csv')
        logger.info("Exported seasonal analysis to european_gas_seasonal_analysis.csv")
        
        return balance
        
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        raise


if __name__ == "__main__":
    complete_balance = main()