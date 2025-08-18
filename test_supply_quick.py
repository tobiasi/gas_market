#!/usr/bin/env python3
"""
Quick test of supply processor - first 10 dates only for validation.
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd
import numpy as np
import openpyxl
import logging
from typing import Dict, List, Tuple, Optional, Any

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def quick_supply_test():
    """Quick test with first 10 dates to validate the approach."""
    logger.info("🔍 QUICK SUPPLY TEST - First 10 dates")
    logger.info("=" * 60)
    
    try:
        # Load Daily sheet to get criteria
        daily_df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
        
        supply_columns = ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 
                         'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']
        
        routes = []
        for col_letter in supply_columns[:5]:  # Test first 5 routes only
            col_idx = ord(col_letter) - ord('A') if len(col_letter) == 1 else (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
            
            category = str(daily_df.iloc[9, col_idx]).strip()
            subcategory = str(daily_df.iloc[10, col_idx]).strip()  
            third_level = str(daily_df.iloc[11, col_idx]).strip()
            
            routes.append({
                'col': col_letter,
                'category': category,
                'subcategory': subcategory,
                'third_level': third_level,
                'wildcard': third_level in ['*', 'ALL', '']
            })
            
            logger.info(f"   {col_letter}: {category} | {subcategory} | {third_level}")
        
        # Load MultiTicker with wide scan
        logger.info("\n📊 Loading MultiTicker data...")
        multiticker_df = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
        
        # Get criteria headers (wide scan)
        criteria_row1 = multiticker_df.iloc[13, 2:400].values  # Category
        criteria_row2 = multiticker_df.iloc[14, 2:400].values  # Subcategory
        criteria_row3 = multiticker_df.iloc[15, 2:400].values  # Third level
        
        # Get first 10 data rows
        dates = multiticker_df.iloc[19:29, 1].values  # First 10 dates
        logger.info(f"   Testing {len(dates)} dates starting from {dates[0]}")
        
        # Test SUMIFS for first date
        logger.info(f"\n🔍 Testing SUMIFS for {dates[0]}:")
        data_row = multiticker_df.iloc[19, 2:400].values  # First data row
        
        for route in routes:
            matching_sum = 0.0
            matches = 0
            
            for i in range(len(criteria_row1)):
                if i >= len(data_row):
                    continue
                    
                c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
                c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
                c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
                
                c1_match = c1 == route['category']
                c2_match = c2 == route['subcategory']
                c3_match = route['wildcard'] or (c3 == route['third_level'])
                
                if c1_match and c2_match and c3_match:
                    value = data_row[i]
                    if pd.notna(value) and isinstance(value, (int, float)):
                        matching_sum += float(value)
                        matches += 1
            
            logger.info(f"   {route['col']} ({route['subcategory']}): {matches} matches = {matching_sum:.2f}")
        
        logger.info("\n✅ QUICK TEST COMPLETED!")
        logger.info("If you see non-zero values above, the approach is working!")
        
    except Exception as e:
        logger.error(f"❌ Quick test failed: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    quick_supply_test()