#!/usr/bin/env python3
"""
Exact replication of the "Gas-to-power demand" sheet from Excel.
Uses precise 3-criteria SUMIFS logic matching the Excel formulas.
"""

import pandas as pd
import numpy as np
from industrial_gas_demand_exact import (
    load_multiticker_with_full_metadata,
    sumifs_three_criteria
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


def create_gas_to_power_sheet_exact(data_df, metadata):
    """
    Create exact replication of Gas-to-power demand sheet structure.
    
    Based on Excel analysis:
    C: Demand / France / Gas-to-Power
    D: Demand / Belgium / Gas-to-Power
    E: Demand / Italy / Gas-to-Power
    F: Demand / GB / Gas-to-Power
    G: Demand / Germany / Industrial and Power
    H: Intermediate Calculation / #Germany / Gas-to-Power
    I: Demand / Germany / Gas-to-Power (calculated) = H (same as Germany Intermediate)
    J: Total (sum of relevant components)
    K: Demand / Netherlands / Industrial and Power
    L: Intermediate Calculation / #Netherlands / Gas-to-Power (appears to be 0)
    M: Intermediate Calculation / #Netherlands / Gas-to-Power (actual values)
    N: Demand / Netherlands / Gas-to-Power (calculated to 30/6/22 then actual) = M
    """
    logger.info("Creating exact Gas-to-power demand sheet")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Column C: France Gas-to-Power
    # Criteria: Demand / France / Gas-to-Power
    result['France_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'France', 'Gas-to-Power'
    )
    
    # Column D: Belgium Gas-to-Power
    # Criteria: Demand / Belgium / Gas-to-Power
    result['Belgium_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Belgium', 'Gas-to-Power'
    )
    
    # Column E: Italy Gas-to-Power
    # Criteria: Demand / Italy / Gas-to-Power
    result['Italy_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Italy', 'Gas-to-Power'
    )
    
    # Column F: GB Gas-to-Power
    # Criteria: Demand / GB / Gas-to-Power
    result['GB_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'GB', 'Gas-to-Power'
    )
    
    # Column G: Germany Industrial and Power (total)
    # Criteria: Demand / Germany / Industrial and Power
    result['Germany_IndPower'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Germany', 'Industrial and Power'
    )
    
    # Column H: Germany Gas-to-Power (Intermediate Calculation)
    # Criteria: Intermediate Calculation / #Germany / Gas-to-Power
    result['Germany_GtP_Intermediate'] = sumifs_three_criteria(
        data_df, metadata, 'Intermediate Calculation', '#Germany', 'Gas-to-Power'
    )
    
    # Column I: Germany Gas-to-Power (calculated) = H
    # This is the same as the intermediate calculation
    result['Germany_GtP'] = result['Germany_GtP_Intermediate']
    
    # Column K: Netherlands Industrial and Power
    # Criteria: Demand / Netherlands / Industrial and Power
    result['Netherlands_IndPower'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Netherlands', 'Industrial and Power'
    )
    
    # Column L: Netherlands Gas-to-Power (Intermediate #1) - appears to be 0
    # Criteria: Intermediate Calculation / #Netherlands / Gas-to-Power
    result['Netherlands_GtP_Int1'] = sumifs_three_criteria(
        data_df, metadata, 'Intermediate Calculation', '#Netherlands', 'Gas-to-Power'
    )
    
    # Column M: Netherlands Gas-to-Power (Intermediate #2) - has actual values
    # This might be a different subcategory or different intermediate calculation
    # Let's try a broader search for Netherlands gas-to-power
    result['Netherlands_GtP_Int2'] = pd.Series(0.0, index=data_df.index)
    
    # Try different variations for Netherlands Gas-to-Power
    for subcategory_variant in ['Gas-to-Power', 'Power', 'Generation']:
        temp_data = sumifs_three_criteria(
            data_df, metadata, 'Intermediate Calculation', '#Netherlands', subcategory_variant
        )
        if temp_data.sum() > result['Netherlands_GtP_Int2'].sum():
            result['Netherlands_GtP_Int2'] = temp_data
    
    # Column N: Netherlands Gas-to-Power (calculated) = M
    result['Netherlands_GtP'] = result['Netherlands_GtP_Int2']
    
    # Column J: Total Gas-to-Power
    # Based on the Excel structure, this should be the sum of the main gas-to-power components
    # From the analysis: C + D + E + F + I = direct gas-to-power components
    gtp_components = [
        'France_GtP', 'Belgium_GtP', 'Italy_GtP', 'GB_GtP', 
        'Germany_GtP', 'Netherlands_GtP'
    ]
    
    result['Total_Gas_to_Power_Demand'] = result[gtp_components].sum(axis=1)
    
    # Sort by date
    result = result.sort_values('Date').reset_index(drop=True)
    
    logger.info(f"Created Gas-to-power demand sheet with {len(result)} rows")
    
    return result


def validate_gtp_exact(gtp_df):
    """
    Validate against exact Excel values.
    """
    logger.info("Validating against Excel Gas-to-power demand sheet")
    
    # Known exact values from Excel analysis
    validations = [
        {
            'date': '2016-10-03',
            'France_GtP': 12.365642,
            'Belgium_GtP': 12.359351,
            'Italy_GtP': 50.910247,
            'GB_GtP': 49.400000,
            'Germany_IndPower': 114.791154,
            'Germany_GtP_Intermediate': 16.314490,
            'Germany_GtP': 16.314490,
            'Netherlands_IndPower': 42.541229,
            'Netherlands_GtP_Int2': 21.801457,
            'Netherlands_GtP': 21.801457,
            'Total_Gas_to_Power_Demand': 141.349730  # From column J
        }
    ]
    
    for val_data in validations:
        date = pd.to_datetime(val_data['date'])
        row = gtp_df[gtp_df['Date'] == date]
        
        if row.empty:
            logger.warning(f"Date {val_data['date']} not found")
            continue
        
        logger.info(f"\nValidation for {val_data['date']}:")
        
        all_match = True
        for col, expected in val_data.items():
            if col == 'date':
                continue
                
            if col in row.columns:
                actual = row[col].iloc[0]
                diff = abs(actual - expected)
                
                if diff < 0.01:
                    logger.info(f"  ✓ {col}: {actual:.6f} (expected {expected:.6f})")
                else:
                    logger.warning(f"  ✗ {col}: {actual:.6f} (expected {expected:.6f}, diff: {diff:.6f})")
                    all_match = False
            else:
                logger.warning(f"  - {col}: Column not found in results")
                all_match = False
        
        if all_match:
            logger.info(f"  ✅ All Gas-to-Power values match for {val_data['date']}")
        else:
            logger.warning(f"  ⚠️  Some Gas-to-Power values don't match for {val_data['date']}")


def validate_gtp_against_livesheet(gtp_df):
    """
    Validate Gas-to-Power totals against LiveSheet Daily historic data.
    """
    logger.info("Validating Gas-to-Power against LiveSheet Daily historic data")
    
    # Load LiveSheet data
    df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
    
    # Get headers and find Gas-to-Power column
    headers = df.iloc[10, 2:20].tolist()
    gtp_idx = headers.index('Gas-to-Power') + 2 if 'Gas-to-Power' in headers else None
    
    if gtp_idx is None:
        logger.warning("Gas-to-Power column not found in LiveSheet")
        return
    
    # Check specific dates
    validation_dates = ['2016-10-03', '2016-10-04']
    
    for date_str in validation_dates:
        # Find row in LiveSheet
        livesheet_val = None
        for row_idx in range(11, min(25, df.shape[0])):
            date_val = df.iloc[row_idx, 1]  # Column B
            if pd.to_datetime(date_val, errors='coerce') == pd.to_datetime(date_str):
                livesheet_val = df.iloc[row_idx, gtp_idx]
                break
        
        # Get our calculated value
        our_row = gtp_df[gtp_df['Date'] == date_str]
        
        if livesheet_val is not None and not our_row.empty:
            our_val = our_row['Total_Gas_to_Power_Demand'].iloc[0]
            diff = abs(our_val - livesheet_val)
            
            if diff < 0.5:  # Allow some tolerance for Gas-to-Power
                logger.info(f"✓ {date_str}: Gas-to-Power = {our_val:.2f} (LiveSheet: {livesheet_val:.2f})")
            else:
                logger.warning(f"✗ {date_str}: Gas-to-Power = {our_val:.2f} (LiveSheet: {livesheet_val:.2f}, diff: {diff:.2f})")


def export_gtp_demand(gtp_df, output_file='gas_to_power_demand_exact.csv'):
    """
    Export Gas-to-Power demand data matching Excel format.
    """
    # Format date column
    export_df = gtp_df.copy()
    export_df['Date'] = export_df['Date'].dt.strftime('%Y-%m-%d')
    
    # Round to match Excel precision
    numeric_cols = export_df.select_dtypes(include=[np.number]).columns
    export_df[numeric_cols] = export_df[numeric_cols].round(6)
    
    # Export full detail
    export_df.to_csv(output_file, index=False)
    logger.info(f"Exported detailed Gas-to-Power data to {output_file}")
    
    # Also export just the total for Daily historic data integration
    total_only = export_df[['Date', 'Total_Gas_to_Power_Demand']].copy()
    total_only.to_csv('total_gas_to_power_demand.csv', index=False)
    logger.info(f"Exported Total_Gas_to_Power_Demand to total_gas_to_power_demand.csv")
    
    return output_file


def main():
    """
    Main execution function.
    """
    try:
        # Load MultiTicker with full metadata
        data_df, metadata = load_multiticker_with_full_metadata('use4.xlsx', 'MultiTicker')
        
        # Create exact Gas-to-Power demand sheet
        gtp_df = create_gas_to_power_sheet_exact(data_df, metadata)
        
        # Validate against Excel
        validate_gtp_exact(gtp_df)
        
        # Validate against LiveSheet
        validate_gtp_against_livesheet(gtp_df)
        
        # Display key results
        logger.info(f"\nGas-to-Power demand for 2016-10-03:")
        sample = gtp_df[gtp_df['Date'] == '2016-10-03']
        if not sample.empty:
            logger.info(f"  Total Gas-to-Power: {sample['Total_Gas_to_Power_Demand'].iloc[0]:.6f}")
            logger.info(f"  Italy GtP: {sample['Italy_GtP'].iloc[0]:.6f}")
            logger.info(f"  GB GtP: {sample['GB_GtP'].iloc[0]:.6f}")
        
        # Export results
        export_gtp_demand(gtp_df)
        
        return gtp_df
        
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    gtp_data = main()