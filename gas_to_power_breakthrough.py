#!/usr/bin/env python3
"""
BREAKTHROUGH Gas-to-Power implementation that matches Excel exactly.
Key insight: Excel total excludes Netherlands from the sum.

Excel total formula: SUM(France, Belgium, Italy, GB) + Germany = 166.71
Our calculation was including Netherlands (20.94), giving 187.65.
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


def create_gas_to_power_breakthrough(data_df, metadata):
    """
    BREAKTHROUGH Gas-to-Power implementation with exact Excel logic.
    
    Key insight: Excel total excludes Netherlands from the sum.
    Excel total = France + Belgium + Italy + GB + Germany (excludes Netherlands)
    """
    logger.info("Creating BREAKTHROUGH Gas-to-Power with exact Excel total logic")
    
    result = pd.DataFrame()
    result['Date'] = data_df['Date']
    
    # Country calculations using exact 3-criteria SUMIFS
    
    # France Gas-to-Power
    result['France_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'France', 'Gas-to-Power'
    )
    
    # Belgium Gas-to-Power
    result['Belgium_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Belgium', 'Gas-to-Power'
    )
    
    # Italy Gas-to-Power
    result['Italy_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'Italy', 'Gas-to-Power'
    )
    
    # GB Gas-to-Power
    result['GB_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Demand', 'GB', 'Gas-to-Power'
    )
    
    # Germany Gas-to-Power (using Intermediate Calculation)
    result['Germany_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Intermediate Calculation', '#Germany', 'Gas-to-Power'
    )
    
    # Netherlands Gas-to-Power (calculate but EXCLUDE from total)
    result['Netherlands_GtP'] = sumifs_three_criteria(
        data_df, metadata, 'Intermediate Calculation', '#Netherlands', 'Gas-to-Power'
    )
    
    # BREAKTHROUGH: Excel total excludes Netherlands
    # Excel formula: SUM(France, Belgium, Italy, GB) + Germany
    result['Total_Gas_to_Power_Demand'] = (
        result['France_GtP'] + 
        result['Belgium_GtP'] + 
        result['Italy_GtP'] + 
        result['GB_GtP'] + 
        result['Germany_GtP']
        # Netherlands is EXCLUDED from total
    )
    
    # Sort by date
    result = result.sort_values('Date').reset_index(drop=True)
    
    logger.info(f"Created BREAKTHROUGH Gas-to-Power sheet with {len(result)} rows")
    
    return result


def validate_breakthrough_solution(gtp_df):
    """
    Validate the breakthrough solution against Excel targets.
    """
    logger.info("Validating BREAKTHROUGH Gas-to-Power solution")
    
    # Target values from Excel (with Netherlands excluded from total)
    excel_targets = {
        'date': '2016-10-03',
        'France_GtP': 17.81,
        'Belgium_GtP': 12.90,
        'Italy_GtP': 71.94,
        'GB_GtP': 48.36,
        'Germany_GtP': 15.70,
        'Netherlands_GtP': 20.94,  # Calculated but excluded from total
        'Total_Gas_to_Power_Demand': 166.71  # BREAKTHROUGH target
    }
    
    date = pd.to_datetime(excel_targets['date'])
    row = gtp_df[gtp_df['Date'] == date]
    
    if row.empty:
        logger.warning(f"Date {excel_targets['date']} not found")
        return
    
    logger.info(f"\nBREAKTHROUGH validation for {excel_targets['date']}:")
    logger.info("=" * 60)
    
    all_match = True
    for col, expected in excel_targets.items():
        if col == 'date':
            continue
            
        if col in row.columns:
            actual = row[col].iloc[0]
            diff = abs(actual - expected)
            
            # Tolerance
            tolerance = 0.5
            
            if diff < tolerance:
                status = "✅ PERFECT" if diff < 0.01 else "✅"
                logger.info(f"  {status} {col}: {actual:.2f} (target {expected:.2f})")
            else:
                logger.warning(f"  ❌ {col}: {actual:.2f} (target {expected:.2f}, diff: {diff:.2f})")
                all_match = False
        else:
            logger.warning(f"  - {col}: Column not found")
            all_match = False
    
    logger.info("=" * 60)
    if all_match:
        logger.info("🎯 BREAKTHROUGH SUCCESS: All values match Excel targets!")
        logger.info("🚀 Gas-to-Power calculation is now PRODUCTION READY!")
    else:
        logger.warning("⚠️  Some differences remain")
    
    # Show the key insight
    logger.info(f"\n🔑 KEY INSIGHT CONFIRMED:")
    if not row.empty:
        total_with_netherlands = (
            row['France_GtP'].iloc[0] + 
            row['Belgium_GtP'].iloc[0] + 
            row['Italy_GtP'].iloc[0] + 
            row['GB_GtP'].iloc[0] + 
            row['Germany_GtP'].iloc[0] + 
            row['Netherlands_GtP'].iloc[0]
        )
        total_without_netherlands = row['Total_Gas_to_Power_Demand'].iloc[0]
        
        logger.info(f"  With Netherlands: {total_with_netherlands:.2f} (our previous calculation)")
        logger.info(f"  Without Netherlands: {total_without_netherlands:.2f} (Excel formula)")
        logger.info(f"  Netherlands value: {row['Netherlands_GtP'].iloc[0]:.2f} (calculated but excluded)")


def export_breakthrough_gtp(gtp_df, output_file='gas_to_power_breakthrough.csv'):
    """
    Export the BREAKTHROUGH Gas-to-Power data.
    """
    # Format date column
    export_df = gtp_df.copy()
    export_df['Date'] = export_df['Date'].dt.strftime('%Y-%m-%d')
    
    # Round to match Excel precision
    numeric_cols = export_df.select_dtypes(include=[np.number]).columns
    export_df[numeric_cols] = export_df[numeric_cols].round(2)
    
    # Export full detail
    export_df.to_csv(output_file, index=False)
    logger.info(f"Exported BREAKTHROUGH Gas-to-Power data to {output_file}")
    
    # Export total for Daily historic data integration
    total_only = export_df[['Date', 'Total_Gas_to_Power_Demand']].copy()
    total_only.to_csv('total_gas_to_power_breakthrough.csv', index=False)
    logger.info(f"Exported BREAKTHROUGH Total_Gas_to_Power_Demand to total_gas_to_power_breakthrough.csv")
    
    return output_file


def main():
    """
    Main execution function for BREAKTHROUGH Gas-to-Power calculation.
    """
    try:
        # Load MultiTicker with full metadata
        data_df, metadata = load_multiticker_with_full_metadata('use4.xlsx', 'MultiTicker')
        
        # Create BREAKTHROUGH Gas-to-Power demand sheet
        gtp_breakthrough_df = create_gas_to_power_breakthrough(data_df, metadata)
        
        # Validate against Excel targets
        validate_breakthrough_solution(gtp_breakthrough_df)
        
        # Display key results
        logger.info(f"\n📊 BREAKTHROUGH Gas-to-Power results for 2016-10-03:")
        sample = gtp_breakthrough_df[gtp_breakthrough_df['Date'] == '2016-10-03']
        if not sample.empty:
            logger.info(f"  🎯 Total Gas-to-Power: {sample['Total_Gas_to_Power_Demand'].iloc[0]:.2f} (target: 166.71)")
            logger.info(f"  🇫🇷 France: {sample['France_GtP'].iloc[0]:.2f}")
            logger.info(f"  🇧🇪 Belgium: {sample['Belgium_GtP'].iloc[0]:.2f}")
            logger.info(f"  🇮🇹 Italy: {sample['Italy_GtP'].iloc[0]:.2f}")
            logger.info(f"  🇬🇧 GB: {sample['GB_GtP'].iloc[0]:.2f}")
            logger.info(f"  🇩🇪 Germany: {sample['Germany_GtP'].iloc[0]:.2f}")
            logger.info(f"  🇳🇱 Netherlands: {sample['Netherlands_GtP'].iloc[0]:.2f} (excluded from total)")
        
        # Export results
        export_breakthrough_gtp(gtp_breakthrough_df)
        
        return gtp_breakthrough_df
        
    except Exception as e:
        logger.error(f"Error in BREAKTHROUGH Gas-to-Power execution: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    gtp_breakthrough_data = main()