#!/usr/bin/env python3
"""
FINAL CORRECTED Supply Processor - Restricts SUMIFS to Daily sheet column range.
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd
import numpy as np
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def create_final_corrected_processor():
    """Create final corrected supply processor with restricted column range."""
    logger.info("🚀 FINAL CORRECTED SUPPLY PROCESSOR")
    logger.info("=" * 60)
    
    try:
        # Load Daily sheet to get exact column range
        daily_df = pd.read_excel('use4.xlsx', sheet_name='Daily historic data by category', header=None)
        
        supply_columns = ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 
                         'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']
        
        # Convert to column indices
        daily_col_indices = []
        for col_letter in supply_columns:
            if len(col_letter) == 1:
                col_idx = ord(col_letter) - ord('A')
            else:
                col_idx = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
            daily_col_indices.append(col_idx)
        
        min_col = min(daily_col_indices)  # Column R = 17
        max_col = max(daily_col_indices)  # Column AI = 34
        
        logger.info(f"✅ Daily sheet column range: {min_col} to {max_col} (Excel {chr(65+min_col)} to AI)")
        
        # Load MultiTicker with RESTRICTED range
        multiticker_df = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
        
        # CRITICAL FIX: Only scan the Daily sheet column range in MultiTicker
        start_scan = min_col - 2  # Adjust for MultiTicker offset (column A, B)
        end_scan = max_col - 2 + 1
        
        logger.info(f"📊 MultiTicker scan range: columns {start_scan} to {end_scan}")
        
        # Extract criteria headers (RESTRICTED SCAN)
        criteria_row1 = multiticker_df.iloc[13, start_scan:end_scan].values  # Category
        criteria_row2 = multiticker_df.iloc[14, start_scan:end_scan].values  # Subcategory  
        criteria_row3 = multiticker_df.iloc[15, start_scan:end_scan].values  # Third level
        
        # Get first data row
        data_row = multiticker_df.iloc[19, start_scan:end_scan].values
        
        logger.info(f"✅ Scanning {len(criteria_row1)} columns (restricted to Daily sheet range)")
        
        # Process Czech_and_Poland with restricted scan
        logger.info(f"\n🔍 Testing Czech_and_Poland with restricted scan:")
        
        czech_matches = []
        for i in range(len(criteria_row1)):
            if i >= len(data_row):
                continue
                
            c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
            c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
            c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
            
            if c1 == "Import" and c2 == "Czech and Poland" and c3 == "Germany":
                value = data_row[i] if pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
                value = float(value)
                
                # Calculate actual Excel column
                actual_col_idx = start_scan + i + 2  # Add back the offset
                excel_col = chr(65 + actual_col_idx) if actual_col_idx < 26 else f"A{chr(65 + actual_col_idx - 26)}"
                
                czech_matches.append((excel_col, value))
        
        logger.info(f"   Found {len(czech_matches)} Czech_and_Poland matches in restricted range:")
        
        czech_total = 0.0
        for excel_col, value in czech_matches:
            czech_total += value
            logger.info(f"      {excel_col}: {value:.2f}")
        
        logger.info(f"   TOTAL: {czech_total:.2f}")
        logger.info(f"   TARGET: 58.41")
        logger.info(f"   DIFFERENCE: {abs(czech_total - 58.41):.2f}")
        
        if abs(czech_total - 58.41) < 1.0:
            logger.info("   ✅ CZECH FIXED - Within 1.0 of target!")
        else:
            logger.warning(f"   ⚠️  Still {abs(czech_total - 58.41):.2f} away from target")
        
        # Test all other major routes with restricted scan
        logger.info(f"\n🔍 Testing other routes with restricted scan:")
        
        test_routes = [
            ('Slovakia', 'Austria', 108.87),
            ('Russia (Nord Stream)', 'Germany', 121.08),
            ('Norway', 'Europe', 206.67),
            ('Netherlands', 'Netherlands', 79.69),
            ('GB', 'GB', 94.01)
        ]
        
        restricted_total = czech_total
        
        for subcategory, third_level, expected in test_routes:
            route_value = 0.0
            
            for i in range(len(criteria_row1)):
                if i >= len(data_row):
                    continue
                    
                c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
                c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
                c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
                
                # Match logic
                c1_match = c1 in ["Import", "Production"]
                c2_match = c2 == subcategory
                c3_match = c3 == third_level
                
                if c1_match and c2_match and c3_match:
                    value = data_row[i] if pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
                    route_value += float(value)
            
            diff = abs(route_value - expected)
            status = "✅" if diff < 1.0 else "❌"
            
            logger.info(f"   {subcategory:<20}: {route_value:8.2f} (expected {expected:8.2f}) {status}")
            restricted_total += route_value
        
        logger.info(f"\n📊 RESTRICTED SCAN RESULTS:")
        logger.info(f"   Partial Total: {restricted_total:.2f}")
        logger.info(f"   Should be much closer to 754.38 benchmark!")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ Processing failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = create_final_corrected_processor()