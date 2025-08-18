#!/usr/bin/env python3
"""
Exact replication of the "LDZ demand" sheet from Excel.
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


def create_ldz_demand_sheet_exact(data_df, metadata):
    """
    Create exact replication of LDZ demand sheet structure.
    
    Based on Excel analysis:
    C: Demand / France / LDZ
    D: Demand / Belgium / LDZ
    E: Demand / Italy / LDZ
    F: Demand / Italy / Other
    G: Demand / Netherlands / LDZ
    H: Demand / GB / LDZ
    I: Demand / Austria / Austria
    J: Demand / Germany / LDZ
    K: Demand / Switzerland / Switzerland
    L: Demand / Luxembourg / Luxembourg
    M: Demand (Net) / Island of Ireland / Island of Ireland
    N: Total (sum of all components)
    """
    logger.info("Creating exact LDZ demand sheet")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Column C: France LDZ
    # Criteria: Demand / France / LDZ
    result['France_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'France', 'LDZ'
    )
    
    # Column D: Belgium LDZ
    # Criteria: Demand / Belgium / LDZ
    result['Belgium_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Belgium', 'LDZ'
    )
    
    # Column E: Italy LDZ
    # Criteria: Demand / Italy / LDZ
    result['Italy_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Italy', 'LDZ'
    )
    
    # Column F: Italy Other
    # Criteria: Demand / Italy / Other
    result['Italy_Other'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Italy', 'Other'
    )
    
    # Column G: Netherlands LDZ
    # Criteria: Demand / Netherlands / LDZ
    result['Netherlands_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Netherlands', 'LDZ'
    )
    
    # Column H: GB LDZ
    # Criteria: Demand / GB / LDZ
    result['GB_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'GB', 'LDZ'
    )
    
    # Column I: Austria Austria (special case - subcategory matches region)
    # Criteria: Demand / Austria / Austria
    result['Austria_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Austria', 'Austria'
    )
    
    # Column J: Germany LDZ
    # Criteria: Demand / Germany / LDZ
    result['Germany_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Germany', 'LDZ'
    )
    
    # Column K: Switzerland Switzerland (special case - subcategory matches region)
    # Criteria: Demand / Switzerland / Switzerland
    result['Switzerland_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Switzerland', 'Switzerland'
    )
    
    # Column L: Luxembourg Luxembourg (special case - subcategory matches region)
    # Criteria: Demand / Luxembourg / Luxembourg
    result['Luxembourg_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Luxembourg', 'Luxembourg'
    )
    
    # Column M: Island of Ireland (Net demand)
    # Criteria: Demand (Net) / Island of Ireland / Island of Ireland
    result['Ireland_LDZ'] = sumifs_three_criteria(
        data_df, metadata, 'Demand (Net)', 'Island of Ireland', 'Island of Ireland'
    )
    
    # Column N: Total LDZ = SUM(C:M)
    ldz_components = [
        'France_LDZ', 'Belgium_LDZ', 'Italy_LDZ', 'Italy_Other',
        'Netherlands_LDZ', 'GB_LDZ', 'Austria_LDZ', 'Germany_LDZ',
        'Switzerland_LDZ', 'Luxembourg_LDZ', 'Ireland_LDZ'
    ]
    
    result['Total_LDZ_Demand'] = result[ldz_components].sum(axis=1)
    
    # Sort by date
    result = result.sort_values('Date').reset_index(drop=True)
    
    logger.info(f"Created LDZ demand sheet with {len(result)} rows")
    
    return result


def validate_ldz_exact(ldz_df):
    """
    Validate against exact Excel values.
    """
    logger.info("Validating against Excel LDZ demand sheet")
    
    # Known exact values from Excel analysis
    validations = [
        {
            'date': '2016-10-03',
            'France_LDZ': 24.193529,
            'Belgium_LDZ': 10.485989,
            'Italy_LDZ': 30.106114,
            'Italy_Other': 1.117208,
            'Netherlands_LDZ': 22.324613,
            'GB_LDZ': 93.410000,
            'Austria_LDZ': 13.203623,
            'Germany_LDZ': 36.030145,
            'Switzerland_LDZ': 4.583913,
            'Luxembourg_LDZ': 0.793636,
            'Ireland_LDZ': 7.160000,
            'Total_LDZ_Demand': 243.408770
        }
    ]
    
    for val_data in validations:
        date = pd.to_datetime(val_data['date'])
        row = ldz_df[ldz_df['Date'] == date]
        
        if row.empty:
            logger.warning(f"Date {val_data['date']} not found")
            continue
        
        logger.info(f"\nValidation for {val_data['date']}:")
        
        all_match = True
        for col, expected in val_data.items():
            if col == 'date':
                continue
                
            actual = row[col].iloc[0]
            diff = abs(actual - expected)
            
            if diff < 0.01:
                logger.info(f"  ✓ {col}: {actual:.6f} (expected {expected:.6f})")
            else:
                logger.warning(f"  ✗ {col}: {actual:.6f} (expected {expected:.6f}, diff: {diff:.6f})")
                all_match = False
        
        if all_match:
            logger.info(f"  ✅ All LDZ values match for {val_data['date']}")
        else:
            logger.warning(f"  ⚠️  Some LDZ values don't match for {val_data['date']}")


def validate_ldz_against_livesheet(ldz_df):
    """
    Validate LDZ totals against LiveSheet Daily historic data.
    """
    logger.info("Validating LDZ against LiveSheet Daily historic data")
    
    # Load LiveSheet data
    df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
    
    # Get headers and find LDZ column
    headers = df.iloc[10, 2:20].tolist()
    ldz_idx = headers.index('LDZ') + 2 if 'LDZ' in headers else None
    
    if ldz_idx is None:
        logger.warning("LDZ column not found in LiveSheet")
        return
    
    # Check specific dates
    validation_dates = ['2016-10-03', '2016-10-04']
    
    for date_str in validation_dates:
        # Find row in LiveSheet
        livesheet_val = None
        for row_idx in range(11, min(25, df.shape[0])):
            date_val = df.iloc[row_idx, 1]  # Column B
            if pd.to_datetime(date_val, errors='coerce') == pd.to_datetime(date_str):
                livesheet_val = df.iloc[row_idx, ldz_idx]
                break
        
        # Get our calculated value
        our_row = ldz_df[ldz_df['Date'] == date_str]
        
        if livesheet_val is not None and not our_row.empty:
            our_val = our_row['Total_LDZ_Demand'].iloc[0]
            diff = abs(our_val - livesheet_val)
            
            if diff < 0.01:
                logger.info(f"✓ {date_str}: LDZ = {our_val:.2f} (LiveSheet: {livesheet_val:.2f})")
            else:
                logger.warning(f"✗ {date_str}: LDZ = {our_val:.2f} (LiveSheet: {livesheet_val:.2f}, diff: {diff:.2f})")


def export_ldz_demand(ldz_df, output_file='ldz_demand_exact.csv'):
    """
    Export LDZ demand data matching Excel format.
    """
    # Format date column
    export_df = ldz_df.copy()
    export_df['Date'] = export_df['Date'].dt.strftime('%Y-%m-%d')
    
    # Round to match Excel precision
    numeric_cols = export_df.select_dtypes(include=[np.number]).columns
    export_df[numeric_cols] = export_df[numeric_cols].round(6)
    
    # Export full detail
    export_df.to_csv(output_file, index=False)
    logger.info(f"Exported detailed LDZ data to {output_file}")
    
    # Also export just the total for Daily historic data integration
    total_only = export_df[['Date', 'Total_LDZ_Demand']].copy()
    total_only.to_csv('total_ldz_demand.csv', index=False)
    logger.info(f"Exported Total_LDZ_Demand to total_ldz_demand.csv")
    
    return output_file


def main():
    """
    Main execution function.
    """
    try:
        # Load MultiTicker with full metadata
        data_df, metadata = load_multiticker_with_full_metadata('use4.xlsx', 'MultiTicker')
        
        # Create exact LDZ demand sheet
        ldz_df = create_ldz_demand_sheet_exact(data_df, metadata)
        
        # Validate against Excel
        validate_ldz_exact(ldz_df)
        
        # Validate against LiveSheet
        validate_ldz_against_livesheet(ldz_df)
        
        # Display key results
        logger.info(f"\nLDZ demand for 2016-10-03:")
        sample = ldz_df[ldz_df['Date'] == '2016-10-03']
        if not sample.empty:
            logger.info(f"  Total LDZ: {sample['Total_LDZ_Demand'].iloc[0]:.6f}")
            logger.info(f"  France LDZ: {sample['France_LDZ'].iloc[0]:.6f}")
            logger.info(f"  GB LDZ: {sample['GB_LDZ'].iloc[0]:.6f}")
        
        # Export results
        export_ldz_demand(ldz_df)
        
        return ldz_df
        
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    ldz_data = main()