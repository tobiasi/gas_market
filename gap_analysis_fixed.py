#!/usr/bin/env python3
"""
STRATEGIC GAP ANALYSIS - Fixed version to identify missing routes
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd
import numpy as np
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def strategic_gap_analysis():
    """Identify exactly what's causing the 50-70 unit gap."""
    logger.info("🎯 STRATEGIC GAP ANALYSIS")
    logger.info("=" * 70)
    
    # Test dates with known Excel totals
    test_cases = [
        ('2016-10-02', 754.38, 'BENCHMARK'),
        ('2016-10-03', 797.86, 'PROBLEM DATE'),
        ('2016-10-04', 840.47, 'PROBLEM DATE'),
        ('2016-10-10', 955.85, 'PROBLEM DATE')
    ]
    
    try:
        # Load MultiTicker data
        multiticker_df = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
        
        # Get dates from MultiTicker
        dates_raw = multiticker_df.iloc[19:, 1].dropna()
        dates_series = pd.to_datetime(dates_raw.values)
        
        logger.info(f"📊 MultiTicker date range: {dates_series.min()} to {dates_series.max()}")
        
        # Analyze each test case
        for target_date, excel_total, status in test_cases:
            logger.info(f"\n🔍 ANALYZING {target_date} ({status})")
            logger.info("-" * 50)
            
            # Find date index
            target_dt = pd.to_datetime(target_date)
            matching_dates = dates_series[dates_series.date == target_dt.date()]
            
            if len(matching_dates) == 0:
                logger.warning(f"Date {target_date} not found in MultiTicker")
                continue
            
            # Get the first matching date's position
            date_position = np.where(dates_series.date == target_dt.date())[0][0]
            data_row_idx = 19 + date_position
            
            logger.info(f"   Date found at MultiTicker row {data_row_idx}")
            
            # Test different column range strategies
            column_strategies = [
                (2, 50, "Conservative (C-AX)"),
                (2, 100, "Standard (C-CV)"), 
                (2, 200, "Extended (C-GR)"),
                (2, 400, "Wide (C-NL)")
            ]
            
            strategy_results = {}
            
            for start_col, end_col, strategy_name in column_strategies:
                max_col = min(end_col, len(multiticker_df.columns))
                
                # Get criteria and data for this range
                criteria_row1 = multiticker_df.iloc[13, start_col:max_col].values
                criteria_row2 = multiticker_df.iloc[14, start_col:max_col].values  
                criteria_row3 = multiticker_df.iloc[15, start_col:max_col].values
                data_row = multiticker_df.iloc[data_row_idx, start_col:max_col].values
                
                # Calculate total using SUMIFS for all supply routes
                total_supply = 0.0
                route_count = 0
                
                for i in range(len(criteria_row1)):
                    if i >= len(data_row):
                        continue
                        
                    c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
                    c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
                    c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
                    
                    # Include all supply-related routes
                    if c1 in ['Import', 'Production', 'Export'] and c2 and c2 != 'nan':
                        value = data_row[i] if pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
                        value = float(value)
                        
                        if abs(value) > 0.1:  # Only significant values
                            total_supply += value
                            route_count += 1
                
                strategy_results[strategy_name] = {
                    'total': total_supply,
                    'routes': route_count,
                    'gap': excel_total - total_supply
                }
                
                gap = excel_total - total_supply
                gap_pct = (gap / excel_total) * 100 if excel_total != 0 else 0
                
                logger.info(f"   {strategy_name:<20}: {total_supply:>8.2f} (gap: {gap:>6.2f}, {gap_pct:>5.1f}%, {route_count} routes)")
            
            # Find best strategy
            best_strategy = min(strategy_results.items(), key=lambda x: abs(x[1]['gap']))
            
            logger.info(f"   {'Excel Total':<20}: {excel_total:>8.2f}")
            logger.info(f"   {'BEST STRATEGY':<20}: {best_strategy[0]} (gap: {best_strategy[1]['gap']:.2f})")
            
            # If still significant gap, identify specific missing routes
            if abs(best_strategy[1]['gap']) > 10.0:
                logger.info(f"\n   🔍 INVESTIGATING MISSING ROUTES (gap: {best_strategy[1]['gap']:.2f}):")
                
                # Use the widest range to find all possible routes
                max_col = min(400, len(multiticker_df.columns))
                criteria_row1 = multiticker_df.iloc[13, 2:max_col].values
                criteria_row2 = multiticker_df.iloc[14, 2:max_col].values  
                criteria_row3 = multiticker_df.iloc[15, 2:max_col].values
                data_row = multiticker_df.iloc[data_row_idx, 2:max_col].values
                
                # Find high-value routes we might be missing
                high_value_routes = []
                
                for i in range(len(criteria_row1)):
                    if i >= len(data_row):
                        continue
                        
                    c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
                    c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
                    c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
                    
                    if c1 in ['Import', 'Production', 'Export'] and c2 and c2 != 'nan':
                        value = data_row[i] if pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
                        value = float(value)
                        
                        if abs(value) > 5.0:  # High-value routes
                            excel_col = chr(65 + i + 2) if i + 2 < 26 else f"A{chr(65 + i + 2 - 26)}"
                            high_value_routes.append((excel_col, c1, c2, c3, value))
                
                # Sort by value and show top contributors
                high_value_routes.sort(key=lambda x: abs(x[4]), reverse=True)
                
                logger.info(f"      Top high-value routes on {target_date}:")
                for excel_col, c1, c2, c3, value in high_value_routes[:10]:
                    logger.info(f"      {excel_col}: {c1}|{c2}|{c3} = {value:.2f}")
        
        logger.info(f"\n🎯 STRATEGIC CONCLUSIONS:")
        logger.info("=" * 50)
        logger.info("1. Check if Wide range (C-NL) gives better accuracy")
        logger.info("2. Identify high-value routes missing from 18-route system")
        logger.info("3. Consider dynamic route inclusion based on significance")
        logger.info("4. Investigate wildcard route expansions")
        
    except Exception as e:
        logger.error(f"❌ Gap analysis failed: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    strategic_gap_analysis()