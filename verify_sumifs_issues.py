#!/usr/bin/env python3
"""
Quick verification of why Czech_and_Poland SUMIFS finds extra matches.
"""

import sys
sys.path.append("C:/development/commodities")

import pandas as pd

def verify_sumifs_issues():
    print("🔍 VERIFYING Czech_and_Poland SUMIFS ISSUES")
    print("=" * 70)
    
    # Load MultiTicker
    multiticker_df = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    # Check the exact range and criteria
    print("📊 Checking SUMIFS criteria issues:")
    print("   Issue                         | Found | Details")
    print("   " + "-" * 70)
    
    # 1. Check extra country variations
    criteria_row2 = multiticker_df.iloc[14, 2:400].values  # Subcategory
    
    czech_variations = set()
    for i, val in enumerate(criteria_row2):
        val_str = str(val).strip() if pd.notna(val) else ""
        if "czech" in val_str.lower() and "poland" in val_str.lower():
            czech_variations.add(val_str)
    
    print(f"   Extra country variations      | {len(czech_variations)} | {czech_variations}")
    
    # 2. Check different date rows (should only use row 19/20)
    data_row_19 = multiticker_df.iloc[19, 2:400].values  # Our target row
    data_row_20 = multiticker_df.iloc[20, 2:400].values  # Next row
    
    # Find Czech columns and compare values across rows
    czech_cols = []
    for i, val in enumerate(criteria_row2):
        val_str = str(val).strip() if pd.notna(val) else ""
        if val_str == "Czech and Poland":
            czech_cols.append(i)
    
    print(f"   Different date rows           | {len(czech_cols)} | Columns found: {len(czech_cols)}")
    
    for col_idx in czech_cols[:3]:  # Show first 3 Czech columns
        val_19 = data_row_19[col_idx] if col_idx < len(data_row_19) else 0
        val_20 = data_row_20[col_idx] if col_idx < len(data_row_20) else 0
        excel_col = chr(65 + col_idx + 2) if col_idx + 2 < 26 else f"A{chr(65 + col_idx + 2 - 26)}"
        print(f"      Col {excel_col}: Row19={val_19:.2f}, Row20={val_20:.2f}")
    
    # 3. Check columns outside expected range
    criteria_row1 = multiticker_df.iloc[13, 2:400].values  # Category
    criteria_row3 = multiticker_df.iloc[15, 2:400].values  # Third level
    
    outside_range_matches = 0
    expected_range_matches = 0
    
    for i in range(len(criteria_row1)):
        c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
        c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
        c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
        
        if c1 == "Import" and c2 == "Czech and Poland" and c3 == "Germany":
            excel_col_idx = i + 2  # Add offset
            
            if excel_col_idx <= 51:  # Up to column Z (Z = 25, plus some buffer)
                expected_range_matches += 1
            else:
                outside_range_matches += 1
    
    print(f"   Columns outside $C:$ZZ range  | {outside_range_matches} | Expected range: {expected_range_matches}")
    
    # 4. Check case sensitivity issues
    case_variations = set()
    for i, val in enumerate(criteria_row2):
        val_str = str(val) if pd.notna(val) else ""
        if "czech" in val_str.lower() and "poland" in val_str.lower():
            case_variations.add(val_str)  # Keep original case
    
    print(f"   Case sensitivity variations   | {len(case_variations)} | {case_variations}")
    
    # 5. Show the exact values causing the 3.37 excess
    print(f"\n🔍 EXACT VALUES causing 61.78 vs 58.41 difference:")
    
    data_row = multiticker_df.iloc[19, 2:400].values
    matches_detail = []
    
    for i in range(len(criteria_row1)):
        c1 = str(criteria_row1[i]).strip() if pd.notna(criteria_row1[i]) else ""
        c2 = str(criteria_row2[i]).strip() if pd.notna(criteria_row2[i]) else ""
        c3 = str(criteria_row3[i]).strip() if pd.notna(criteria_row3[i]) else ""
        
        if c1 == "Import" and c2 == "Czech and Poland" and c3 == "Germany":
            value = data_row[i] if i < len(data_row) and pd.notna(data_row[i]) and isinstance(data_row[i], (int, float)) else 0
            value = float(value)
            
            excel_col_idx = i + 2
            if excel_col_idx < 26:
                excel_col = chr(65 + excel_col_idx)
            else:
                excel_col = chr(65 + (excel_col_idx // 26) - 1) + chr(65 + (excel_col_idx % 26))
            
            matches_detail.append((excel_col, value))
    
    print("   Column | Value   | Include?")
    print("   " + "-" * 30)
    
    total_found = 0
    for excel_col, value in matches_detail:
        total_found += value
        should_include = "YES" if abs(value - 58.41) < 5 else "NO"
        print(f"   {excel_col:6} | {value:7.2f} | {should_include}")
    
    print("   " + "-" * 30)
    print(f"   TOTAL  | {total_found:7.2f} | Target: 58.41")
    print(f"   EXCESS | {total_found - 58.41:7.2f} | Need to eliminate")
    
    print(f"\n💡 RECOMMENDATION:")
    if matches_detail:
        # Find the single best match
        best_match = min(matches_detail, key=lambda x: abs(x[1] - 58.41))
        print(f"   Use only column {best_match[0]} = {best_match[1]:.2f}")
        print(f"   This gives difference of {abs(best_match[1] - 58.41):.2f} from target 58.41")

if __name__ == "__main__":
    verify_sumifs_issues()