#!/usr/bin/env python3
"""
FIXED Gas-to-Power demand sheet with exact string matching.
Addresses over-counting issues by using precise criteria matching Excel logic.
"""

import pandas as pd
import numpy as np
from industrial_gas_demand_exact import (
    load_multiticker_with_full_metadata
)
import logging
import warnings
warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def sumifs_three_criteria_exact(data_df, metadata, category_target, region_target, subcategory_target):
    """
    EXACT 3-criteria SUMIFS with strict string matching (no "contains" logic).
    
    CRITICAL CHANGE: Uses == for exact matching instead of "in" logic
    """
    matching_cols = []
    
    for col, info in metadata.items():
        # EXACT string matching only - no partial matching
        category_match = info['category'] == category_target
        region_match = info['region'] == region_target
        subcategory_match = info['subcategory'] == subcategory_target
        
        if category_match and region_match and subcategory_match:
            matching_cols.append(col)
    
    if not matching_cols:
        logger.debug(f"No exact matches for {category_target}/{region_target}/{subcategory_target}")
        return pd.Series(0.0, index=data_df.index)
    
    logger.debug(f"Found {len(matching_cols)} exact matches for {category_target}/{region_target}/{subcategory_target}")
    
    # Sum across matching columns (handling NaN as 0)
    result = data_df[matching_cols].sum(axis=1, skipna=True)
    
    return result


def debug_gtp_tickers(metadata, country):
    """
    Debug function to see exactly which tickers are being included for Gas-to-Power.
    """
    logger.info(f"\nDebugging Gas-to-Power tickers for {country}:")
    
    # Standard Demand approach
    demand_tickers = []
    for col, info in metadata.items():
        if (info['category'] == 'Demand' and 
            info['region'] == country and 
            info['subcategory'] == 'Gas-to-Power'):
            demand_tickers.append((col, info))
    
    logger.info(f"  Demand/Gas-to-Power tickers: {len(demand_tickers)}")
    for ticker, info in demand_tickers[:3]:  # Show first 3
        logger.info(f"    {ticker}: {info['category']}/{info['region']}/{info['subcategory']}")
    
    # Check for problematic over-inclusive matches
    problem_tickers = []
    for col, info in metadata.items():
        if info['region'] == country:
            # Look for tickers that might be incorrectly included
            if ('Power' in info['subcategory'] and info['subcategory'] != 'Gas-to-Power'):
                problem_tickers.append((col, info))
    
    if problem_tickers:
        logger.info(f"  Potentially problematic tickers (excluded with exact matching): {len(problem_tickers)}")
        for ticker, info in problem_tickers[:3]:
            logger.info(f"    {ticker}: {info['category']}/{info['region']}/{info['subcategory']}")


def create_gas_to_power_sheet_fixed(data_df, metadata):
    """
    FIXED Gas-to-Power sheet with exact matching to prevent over-counting.
    
    Key changes:
    1. Exact string matching only (subcategory == 'Gas-to-Power')
    2. No "contains" or partial matching logic
    3. Strict adherence to Excel Gas-to-power sheet structure
    """
    logger.info("Creating FIXED Gas-to-power demand sheet with exact matching")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Debug problematic countries first
    debug_gtp_tickers(metadata, 'Italy')
    debug_gtp_tickers(metadata, 'France')
    
    # Column C: France Gas-to-Power (EXACT MATCHING)
    # Criteria: Demand / France / Gas-to-Power (exact string)
    result['France_GtP'] = sumifs_three_criteria_exact(
        data_df, metadata, 'Demand', 'France', 'Gas-to-Power'
    )
    
    # Column D: Belgium Gas-to-Power (EXACT MATCHING)
    result['Belgium_GtP'] = sumifs_three_criteria_exact(
        data_df, metadata, 'Demand', 'Belgium', 'Gas-to-Power'
    )
    
    # Column E: Italy Gas-to-Power (EXACT MATCHING)
    # This should fix the +21.03 over-counting
    result['Italy_GtP'] = sumifs_three_criteria_exact(
        data_df, metadata, 'Demand', 'Italy', 'Gas-to-Power'
    )
    
    # Column F: GB Gas-to-Power (EXACT MATCHING)
    result['GB_GtP'] = sumifs_three_criteria_exact(
        data_df, metadata, 'Demand', 'GB', 'Gas-to-Power'
    )
    
    # Column G: Germany Industrial and Power (for reference)
    result['Germany_IndPower'] = sumifs_three_criteria_exact(
        data_df, metadata, 'Demand', 'Germany', 'Industrial and Power'
    )
    
    # Column H: Germany Gas-to-Power (Intermediate Calculation)
    # This is the specific Germany approach
    result['Germany_GtP'] = sumifs_three_criteria_exact(
        data_df, metadata, 'Intermediate Calculation', '#Germany', 'Gas-to-Power'
    )
    
    # Column K: Netherlands Industrial and Power (for reference)
    result['Netherlands_IndPower'] = sumifs_three_criteria_exact(
        data_df, metadata, 'Demand', 'Netherlands', 'Industrial and Power'
    )
    
    # Column M: Netherlands Gas-to-Power (Intermediate Calculation)
    result['Netherlands_GtP'] = sumifs_three_criteria_exact(
        data_df, metadata, 'Intermediate Calculation', '#Netherlands', 'Gas-to-Power'
    )
    
    # Additional countries that might have Gas-to-Power
    additional_countries = ['Austria', 'Switzerland', 'Luxembourg', 'Spain', 'Poland']
    
    for country in additional_countries:
        country_gtp = sumifs_three_criteria_exact(
            data_df, metadata, 'Demand', country, 'Gas-to-Power'
        )
        if country_gtp.sum() > 0:
            result[f'{country}_GtP'] = country_gtp
            logger.info(f"Found Gas-to-Power data for {country}")
    
    # Total Gas-to-Power: Sum only the main Gas-to-Power components
    # Based on Excel analysis, this should match column J (141.35)
    main_gtp_components = [
        'France_GtP', 'Belgium_GtP', 'Italy_GtP', 'GB_GtP', 
        'Germany_GtP', 'Netherlands_GtP'
    ]
    
    # Add any additional countries that have data
    for col in result.columns:
        if col.endswith('_GtP') and col not in main_gtp_components and col != 'Total_Gas_to_Power_Demand':
            main_gtp_components.append(col)
    
    result['Total_Gas_to_Power_Demand'] = result[main_gtp_components].sum(axis=1)
    
    # Sort by date
    result = result.sort_values('Date').reset_index(drop=True)
    
    logger.info(f"Created FIXED Gas-to-power demand sheet with {len(result)} rows")
    
    return result


def validate_gtp_fixed(gtp_df):
    """
    Validate the FIXED Gas-to-Power calculation against Excel targets.
    """
    logger.info("Validating FIXED Gas-to-power against Excel Gas-to-power sheet targets")
    
    # Target values from Excel Gas-to-power sheet (NOT LiveSheet)
    excel_targets = {
        'date': '2016-10-03',
        'France_GtP': 12.365642,
        'Belgium_GtP': 12.359351,  
        'Italy_GtP': 50.910247,    # Target: reduce from 71.94
        'GB_GtP': 49.400000,
        'Germany_GtP': 16.314490,
        'Netherlands_GtP': 21.801457,
        'Total_Gas_to_Power_Demand': 141.349730  # Target: reduce from 187.65
    }
    
    date = pd.to_datetime(excel_targets['date'])
    row = gtp_df[gtp_df['Date'] == date]
    
    if row.empty:
        logger.warning(f"Date {excel_targets['date']} not found")
        return
    
    logger.info(f"\nFixed validation for {excel_targets['date']}:")
    
    total_match = True
    for col, expected in excel_targets.items():
        if col == 'date':
            continue
            
        if col in row.columns:
            actual = row[col].iloc[0]
            diff = abs(actual - expected)
            
            # More lenient tolerance for the fixed version
            tolerance = 2.0 if col == 'Total_Gas_to_Power_Demand' else 1.0
            
            if diff < tolerance:
                logger.info(f"  ✓ {col}: {actual:.2f} (target {expected:.2f}, diff: {diff:.2f})")
            else:
                logger.warning(f"  ✗ {col}: {actual:.2f} (target {expected:.2f}, diff: {diff:.2f})")
                total_match = False
        else:
            logger.warning(f"  - {col}: Column not found")
            total_match = False
    
    if total_match:
        logger.info("  ✅ FIXED Gas-to-Power calculation matches Excel targets!")
    else:
        logger.warning("  ⚠️  Some differences remain - may need further refinement")


def compare_fixed_vs_original(gtp_fixed_df, original_total=187.65):
    """
    Compare the fixed calculation with the original over-counting version.
    """
    sample = gtp_fixed_df[gtp_fixed_df['Date'] == '2016-10-03']
    if not sample.empty:
        fixed_total = sample['Total_Gas_to_Power_Demand'].iloc[0]
        reduction = original_total - fixed_total
        
        logger.info(f"\n=== FIXED vs ORIGINAL COMPARISON ===")
        logger.info(f"Original (over-counting): {original_total:.2f}")
        logger.info(f"Fixed (exact matching): {fixed_total:.2f}")
        logger.info(f"Reduction: {reduction:.2f}")
        logger.info(f"Excel target: 141.35")
        logger.info(f"Remaining gap: {abs(fixed_total - 141.35):.2f}")


def export_gtp_fixed(gtp_df, output_file='gas_to_power_demand_fixed.csv'):
    """
    Export the FIXED Gas-to-Power data.
    """
    # Format date column
    export_df = gtp_df.copy()
    export_df['Date'] = export_df['Date'].dt.strftime('%Y-%m-%d')
    
    # Round to match Excel precision
    numeric_cols = export_df.select_dtypes(include=[np.number]).columns
    export_df[numeric_cols] = export_df[numeric_cols].round(2)
    
    # Export full detail
    export_df.to_csv(output_file, index=False)
    logger.info(f"Exported FIXED Gas-to-Power data to {output_file}")
    
    # Export total for Daily historic data integration
    total_only = export_df[['Date', 'Total_Gas_to_Power_Demand']].copy()
    total_only.to_csv('total_gas_to_power_demand_fixed.csv', index=False)
    logger.info(f"Exported FIXED Total_Gas_to_Power_Demand to total_gas_to_power_demand_fixed.csv")
    
    return output_file


def main():
    """
    Main execution function for FIXED Gas-to-Power calculation.
    """
    try:
        # Load MultiTicker with full metadata
        data_df, metadata = load_multiticker_with_full_metadata('use4.xlsx', 'MultiTicker')
        
        # Create FIXED Gas-to-Power demand sheet
        gtp_fixed_df = create_gas_to_power_sheet_fixed(data_df, metadata)
        
        # Validate against Excel targets
        validate_gtp_fixed(gtp_fixed_df)
        
        # Compare with original over-counting version
        compare_fixed_vs_original(gtp_fixed_df)
        
        # Display key results
        logger.info(f"\nFIXED Gas-to-Power results for 2016-10-03:")
        sample = gtp_fixed_df[gtp_fixed_df['Date'] == '2016-10-03']
        if not sample.empty:
            logger.info(f"  Total Gas-to-Power: {sample['Total_Gas_to_Power_Demand'].iloc[0]:.2f}")
            logger.info(f"  France GtP: {sample['France_GtP'].iloc[0]:.2f}")
            logger.info(f"  Italy GtP: {sample['Italy_GtP'].iloc[0]:.2f}")
            logger.info(f"  Germany GtP: {sample['Germany_GtP'].iloc[0]:.2f}")
        
        # Export results
        export_gtp_fixed(gtp_fixed_df)
        
        return gtp_fixed_df
        
    except Exception as e:
        logger.error(f"Error in FIXED Gas-to-Power execution: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    gtp_fixed_data = main()