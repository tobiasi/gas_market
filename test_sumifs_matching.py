#!/usr/bin/env python3
"""
Test SUMIFS matching to debug why we're getting zeros.
"""

import pandas as pd
import sys
sys.path.append("C:/development/commodities")

def test_sumifs_matching():
    print("🔍 TESTING SUMIFS MATCHING LOGIC")
    print("=" * 80)
    
    # Load Excel file
    df_full = pd.read_excel('use4.xlsx', sheet_name='MultiTicker', header=None)
    
    # Check structure
    category_row = 14 - 1      # Row 14 (0-based = 13)  
    subcategory_row = 15 - 1   # Row 15 (0-based = 14)
    third_level_row = 16 - 1   # Row 16 (0-based = 15)
    data_start = 19 - 1        # Row 19 (0-based = 18)
    
    # Extract headers for first 100 columns
    headers = {}
    for col_idx in range(2, min(100, len(df_full.columns))):  # Start from column C
        col_name = f'Col_{col_idx-1}'  # Col_1, Col_2, etc.
        
        category = df_full.iloc[category_row, col_idx]
        subcategory = df_full.iloc[subcategory_row, col_idx]
        third_level = df_full.iloc[third_level_row, col_idx]
        
        headers[col_name] = {
            'category': str(category) if pd.notna(category) else '',
            'subcategory': str(subcategory) if pd.notna(subcategory) else '',
            'third_level': str(third_level) if pd.notna(third_level) else '',
            'excel_col_index': col_idx,
        }
    
    print(f"✅ Extracted {len(headers)} headers")
    
    # Test specific matching criteria
    test_criteria = [
        {
            'name': 'MAB_to_Austria',
            'category': ['Import', 'Imports'],
            'subcategory': ['MAB'],
            'third_level': ['Austria']
        },
        {
            'name': 'Norway_to_Europe', 
            'category': ['Import', 'Imports'],
            'subcategory': ['Norway'],
            'third_level': ['Europe', '#Europe']
        },
        {
            'name': 'Austria_Production',
            'category': ['Production'],
            'subcategory': ['Austria'],
            'third_level': ['Austria']
        },
        {
            'name': 'Germany_Production',
            'category': ['Production'],
            'subcategory': ['Germany'],
            'third_level': ['Germany']
        }
    ]
    
    def flexible_match(actual_value, criteria_list):
        if not actual_value or pd.isna(actual_value):
            return False
        
        actual_clean = str(actual_value).strip().lower()
        
        for criterion in criteria_list:
            if not criterion:
                continue
            criterion_clean = str(criterion).strip().lower()
            
            if actual_clean == criterion_clean:
                return True
            if criterion_clean in actual_clean or actual_clean in criterion_clean:
                return True
        
        return False
    
    # Test each criteria
    for criteria in test_criteria:
        print(f"\n🎯 Testing {criteria['name']}:")
        print(f"   Looking for: Category={criteria['category']}, Sub={criteria['subcategory']}, Third={criteria['third_level']}")
        
        matches = []
        for col_name, col_info in headers.items():
            # Test each criteria separately
            cat_match = flexible_match(col_info['category'], criteria['category'])
            sub_match = flexible_match(col_info['subcategory'], criteria['subcategory'])  
            third_match = flexible_match(col_info['third_level'], criteria['third_level'])
            
            if cat_match and sub_match and third_match:
                matches.append({
                    'column': col_name,
                    'excel_idx': col_info['excel_col_index'],
                    'category': col_info['category'],
                    'subcategory': col_info['subcategory'],
                    'third_level': col_info['third_level']
                })
        
        if matches:
            print(f"   ✅ Found {len(matches)} matches:")
            for match in matches:
                excel_col = chr(65 + match['excel_idx']) if match['excel_idx'] < 26 else f"A{chr(65 + match['excel_idx'] - 26)}"
                print(f"      {excel_col}: {match['category']} | {match['subcategory']} | {match['third_level']}")
        else:
            print(f"   ❌ No matches found")
            
            # Debug why no matches
            print(f"   🔍 Debug: Checking all columns for partial matches...")
            for col_name, col_info in list(headers.items())[:20]:  # Check first 20
                cat_match = flexible_match(col_info['category'], criteria['category'])
                sub_match = flexible_match(col_info['subcategory'], criteria['subcategory'])  
                third_match = flexible_match(col_info['third_level'], criteria['third_level'])
                
                if cat_match or sub_match or third_match:
                    excel_col = chr(65 + col_info['excel_col_index']) if col_info['excel_col_index'] < 26 else f"A{chr(65 + col_info['excel_col_index'] - 26)}"
                    print(f"      {excel_col}: Cat={cat_match}, Sub={sub_match}, Third={third_match} | {col_info['category']} | {col_info['subcategory']} | {col_info['third_level']}")

if __name__ == "__main__":
    test_sumifs_matching()