#!/usr/bin/env python3
"""
Create all subcategory totals (Industrial, LDZ, Gas-to-Power) using the same approach.
Each follows similar sheet structure patterns from Excel.
"""

import pandas as pd
import numpy as np
from industrial_gas_demand_exact import (
    load_multiticker_with_full_metadata,
    sumifs_three_criteria
)
import logging

logger = logging.getLogger(__name__)


def create_ldz_demand_totals(data_df, metadata):
    """
    Create LDZ demand totals following similar pattern to Industrial.
    LDZ = Local Distribution Zone demand.
    """
    logger.info("Calculating LDZ demand totals")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Find all LDZ demand across countries
    countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 
                'Austria', 'Germany', 'Switzerland', 'Luxembourg']
    
    ldz_total = pd.Series(0.0, index=data_df.index)
    
    for country in countries:
        country_ldz = sumifs_three_criteria(
            data_df, metadata, 'Demand', country, 'LDZ'
        )
        
        if country_ldz.sum() > 0:
            result[f'{country}_LDZ'] = country_ldz
            ldz_total += country_ldz
            logger.info(f"Found LDZ data for {country}: mean={country_ldz.mean():.2f}")
    
    result['Total_LDZ'] = ldz_total
    
    return result


def create_gas_to_power_totals(data_df, metadata):
    """
    Create Gas-to-Power demand totals.
    """
    logger.info("Calculating Gas-to-Power demand totals")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Find all Gas-to-Power demand across countries
    countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 
                'Austria', 'Germany', 'Switzerland', 'Luxembourg']
    
    gtp_total = pd.Series(0.0, index=data_df.index)
    
    for country in countries:
        # Try different Gas-to-Power variations
        gtp_variations = ['Gas-to-Power', 'Gas to Power', 'Power Generation']
        
        for variation in gtp_variations:
            country_gtp = sumifs_three_criteria(
                data_df, metadata, 'Demand', country, variation
            )
            
            if country_gtp.sum() > 0:
                result[f'{country}_GtP_{variation.replace("-", "_").replace(" ", "_")}'] = country_gtp
                gtp_total += country_gtp
                logger.info(f"Found {variation} data for {country}: mean={country_gtp.mean():.2f}")
                break  # Use first match only
    
    # Also check for Intermediate Calculation Gas-to-Power (like Germany)
    intermediate_gtp = sumifs_three_criteria(
        data_df, metadata, 'Intermediate Calculation', '#Germany', 'Gas-to-Power'
    )
    if intermediate_gtp.sum() > 0:
        result['Germany_Intermediate_GtP'] = intermediate_gtp
        gtp_total += intermediate_gtp
        logger.info(f"Found Intermediate Gas-to-Power for Germany: mean={intermediate_gtp.mean():.2f}")
    
    result['Total_Gas_to_Power'] = gtp_total
    
    return result


def create_all_subcategory_totals(file_path='use4.xlsx'):
    """
    Create all subcategory totals and return as single DataFrame.
    """
    logger.info("Creating all subcategory totals")
    
    # Load data
    data_df, metadata = load_multiticker_with_full_metadata(file_path, 'MultiTicker')
    
    # Create Industrial demand (using existing function)
    from industrial_gas_demand_exact import create_industrial_demand_sheet_exact
    industrial_df = create_industrial_demand_sheet_exact(data_df, metadata)
    
    # Create LDZ demand
    ldz_df = create_ldz_demand_totals(data_df, metadata)
    
    # Create Gas-to-Power demand
    gtp_df = create_gas_to_power_totals(data_df, metadata)
    
    # Combine into single DataFrame with just the totals
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    result['Industrial'] = industrial_df['Total_Industrial_Demand']
    result['LDZ'] = ldz_df['Total_LDZ'] 
    result['Gas_to_Power'] = gtp_df['Total_Gas_to_Power']
    
    # Sort by date
    result = result.sort_values('Date').reset_index(drop=True)
    
    logger.info(f"Created complete subcategory totals with {len(result)} rows")
    
    return result, industrial_df, ldz_df, gtp_df


def validate_subcategories_against_livesheet(subcategory_df):
    """
    Validate subcategory totals against LiveSheet Daily historic data.
    """
    logger.info("Validating subcategories against LiveSheet")
    
    # Load LiveSheet data
    df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
    
    # Get headers
    headers = df.iloc[10, 2:20].tolist()
    
    # Find column indices
    industrial_idx = headers.index('Industrial') + 2 if 'Industrial' in headers else None
    ldz_idx = headers.index('LDZ') + 2 if 'LDZ' in headers else None
    gtp_idx = headers.index('Gas-to-Power') + 2 if 'Gas-to-Power' in headers else None
    
    # Check specific dates
    validation_dates = ['2016-10-03', '2016-10-04']
    
    for date_str in validation_dates:
        logger.info(f"\nValidation for {date_str}:")
        
        # Find row in LiveSheet
        livesheet_row = None
        for row_idx in range(11, min(25, df.shape[0])):
            date_val = df.iloc[row_idx, 1]  # Column B
            if pd.to_datetime(date_val, errors='coerce') == pd.to_datetime(date_str):
                livesheet_row = df.iloc[row_idx]
                break
        
        if livesheet_row is None:
            logger.warning(f"  Date {date_str} not found in LiveSheet")
            continue
        
        # Get our calculated values
        our_row = subcategory_df[subcategory_df['Date'] == date_str]
        if our_row.empty:
            logger.warning(f"  Date {date_str} not found in our calculations")
            continue
        
        our_row = our_row.iloc[0]
        
        # Compare each subcategory
        for col_name, idx in [('Industrial', industrial_idx), ('LDZ', ldz_idx), ('Gas_to_Power', gtp_idx)]:
            if idx is not None:
                livesheet_val = livesheet_row.iloc[idx]
                our_val = our_row[col_name]
                diff = abs(our_val - livesheet_val)
                
                if diff < 0.01:
                    logger.info(f"  ✓ {col_name}: {our_val:.2f} (LiveSheet: {livesheet_val:.2f})")
                else:
                    logger.warning(f"  ✗ {col_name}: {our_val:.2f} (LiveSheet: {livesheet_val:.2f}, diff: {diff:.2f})")


def main():
    """
    Main execution function.
    """
    try:
        # Create all subcategory totals
        subcategory_totals, industrial_detail, ldz_detail, gtp_detail = create_all_subcategory_totals('use4.xlsx')
        
        # Validate against LiveSheet
        validate_subcategories_against_livesheet(subcategory_totals)
        
        # Export results
        subcategory_totals['Date'] = subcategory_totals['Date'].dt.strftime('%Y-%m-%d')
        subcategory_totals.round(2).to_csv('complete_subcategory_totals.csv', index=False)
        logger.info("Exported complete subcategory totals to complete_subcategory_totals.csv")
        
        # Display summary
        logger.info("\nSubcategory totals for 2016-10-03:")
        sample = subcategory_totals[subcategory_totals['Date'] == '2016-10-03']
        if not sample.empty:
            for col in ['Industrial', 'LDZ', 'Gas_to_Power']:
                logger.info(f"  {col}: {sample[col].iloc[0]:.2f}")
        
        return subcategory_totals
        
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    subcategory_data = main()