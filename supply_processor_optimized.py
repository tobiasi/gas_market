#!/usr/bin/env python3
"""
OPTIMIZED Supply Processor - Processes in chunks for better performance.
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd
import numpy as np
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def run_optimized_supply_processing():
    """Run optimized supply processing with chunking."""
    logger.info("🚀 OPTIMIZED SUPPLY PROCESSING")
    logger.info("=" * 60)
    
    try:
        # Load and extract route criteria (PROVEN approach)
        daily_df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
        
        supply_columns = ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 
                         'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']
        
        routes = []
        for col_letter in supply_columns:
            # Convert column letter to index
            if len(col_letter) == 1:
                col_idx = ord(col_letter) - ord('A')
            else:  # AA, AB, etc.
                col_idx = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
            
            category = str(daily_df.iloc[9, col_idx]).strip()
            subcategory = str(daily_df.iloc[10, col_idx]).strip()  
            third_level = str(daily_df.iloc[11, col_idx]).strip()
            
            if category and subcategory:  # Valid route
                route_name = f"{subcategory.replace(' ', '_').replace('(', '').replace(')', '')}"
                if category.lower() == 'import':
                    if third_level not in ['*', 'ALL', '']:
                        route_name += f"_to_{third_level.replace(' ', '_')}"
                    else:
                        route_name += "_Import"
                elif category.lower() == 'production':
                    route_name += "_Production"
                elif category.lower() == 'export':
                    route_name += f"_to_{third_level.replace(' ', '_')}_Export"
                
                routes.append({
                    'col': col_letter,
                    'route_name': route_name,
                    'category': category,
                    'subcategory': subcategory,
                    'third_level': third_level,
                    'wildcard': third_level in ['*', 'ALL', '']
                })
        
        logger.info(f"✅ Extracted {len(routes)} supply routes")
        for route in routes:
            logger.info(f"   {route['col']}: {route['route_name']}")
        
        # Load MultiTicker data 
        logger.info("\n📊 Loading MultiTicker data...")
        multiticker_df = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
        
        # Extract criteria headers (WIDE SCAN - up to column 400)
        max_col = min(400, len(multiticker_df.columns))
        criteria_row1 = multiticker_df.iloc[13, 2:max_col].values  # Category
        criteria_row2 = multiticker_df.iloc[14, 2:max_col].values  # Subcategory  
        criteria_row3 = multiticker_df.iloc[15, 2:max_col].values  # Third level
        
        # Get dates
        dates = multiticker_df.iloc[19:, 1].dropna().values
        dates_df = pd.to_datetime(dates)
        
        logger.info(f"   Loaded {len(dates_df)} dates from {dates_df.min()} to {dates_df.max()}")
        
        # Process in chunks for better performance
        chunk_size = 500
        total_chunks = (len(dates_df) + chunk_size - 1) // chunk_size
        
        # Initialize results DataFrame
        results_df = pd.DataFrame()
        results_df['Date'] = dates_df
        for route in routes:
            results_df[route['route_name']] = 0.0
        
        logger.info(f"\n🛠️ Processing {len(routes)} routes in {total_chunks} chunks of {chunk_size} dates")
        
        # Process each chunk
        for chunk_idx in range(total_chunks):
            start_idx = chunk_idx * chunk_size
            end_idx = min(start_idx + chunk_size, len(dates_df))
            
            logger.info(f"   Processing chunk {chunk_idx + 1}/{total_chunks} (dates {start_idx + 1}-{end_idx})")
            
            # Process each date in chunk
            for date_idx in range(start_idx, end_idx):
                data_row_idx = 19 + date_idx  # Row 20 + offset
                data_row = multiticker_df.iloc[data_row_idx, 2:max_col].values
                
                # Process each route for this date
                for route in routes:
                    matching_sum = 0.0
                    
                    for i in range(len(criteria_row1)):
                        if i >= len(data_row):
                            continue
                            
                        c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
                        c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
                        c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
                        
                        # EXACT matching (PROVEN approach)
                        c1_match = c1 == route['category']
                        c2_match = c2 == route['subcategory']
                        c3_match = route['wildcard'] or (c3 == route['third_level'])
                        
                        if c1_match and c2_match and c3_match:
                            value = data_row[i]
                            if pd.notna(value) and isinstance(value, (int, float)):
                                matching_sum += float(value)
                    
                    results_df.iloc[date_idx, results_df.columns.get_loc(route['route_name'])] = matching_sum
        
        # Calculate Total_Supply
        logger.info("\n📊 Calculating Total_Supply...")
        supply_cols = [col for col in results_df.columns if col != 'Date']
        results_df['Total_Supply'] = results_df[supply_cols].sum(axis=1, skipna=True)
        
        # Validate first date
        logger.info("\n🔍 Validation - First Date Results:")
        first_row = results_df.iloc[0]
        logger.info(f"   Date: {first_row['Date'].strftime('%Y-%m-%d')}")
        
        # Expected values from your verification
        validation_targets = {
            'Slovakia_to_Austria': 108.87,
            'Russia_Nord_Stream_to_Germany': 121.08, 
            'Norway_to_Europe': 206.67,
            'Netherlands_Production': 79.69,
            'GB_Production': 94.01
        }
        
        for route_name, expected in validation_targets.items():
            actual = first_row.get(route_name, 0)
            diff = abs(actual - expected) if actual else expected
            status = "✅" if diff < 1.0 else "❌"
            logger.info(f"   {route_name}: {actual:.2f} (expected {expected:.2f}) {status}")
        
        total_actual = first_row['Total_Supply']
        logger.info(f"   Total_Supply: {total_actual:.2f}")
        
        # Export results
        logger.info("\n💾 Exporting results...")
        export_df = results_df.copy()
        export_df['Date'] = export_df['Date'].dt.strftime('%Y-%m-%d')
        export_df = export_df.round(2)
        
        export_df.to_csv('supply_results.csv', index=False)
        
        logger.info(f"✅ Exported {len(export_df)} rows × {len(export_df.columns)} columns to supply_results.csv")
        
        # Show sample
        logger.info("\n📋 Sample results:")
        sample_cols = ['Date'] + list(export_df.columns)[1:6] + ['Total_Supply']
        logger.info(export_df[sample_cols].head(3).to_string(index=False))
        
        logger.info("\n🎉 OPTIMIZED PROCESSING COMPLETE!")
        return results_df
        
    except Exception as e:
        logger.error(f"❌ Processing failed: {str(e)}")
        import traceback
        traceback.print_exc()
        raise

if __name__ == "__main__":
    result = run_optimized_supply_processing()