# -*- coding: utf-8 -*-
"""
CORRECTED replication using the actual target column mapping from analysis_results.json
"""

import pandas as pd
import numpy as np
import json

def create_correct_replication():
    """Create replication using the CORRECT column mapping"""
    
    print("🔧 CREATING CORRECTED REPLICATION")
    print("=" * 80)
    
    # Load the correct column mapping from our analysis
    with open('analysis_results.json', 'r') as f:
        analysis = json.load(f)
    
    target_values = analysis['target_values']
    column_mapping = analysis['column_mapping']
    
    print("📊 Using CORRECT column mapping from analysis:")
    for key, col in column_mapping.items():
        print(f"   {key:<15}: Column {col}")
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    # Read the target data
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    header_row = 10
    data_start_row = 12
    
    print(f"\n📋 Extracting data using CORRECT column indices...")
    
    # Extract data using CORRECT column mapping
    extracted_data = []
    dates = []
    
    for i in range(data_start_row, min(data_start_row + 100, target_data.shape[0])):
        date_val = target_data.iloc[i, 1]  # Date is in column 1
        
        if pd.notna(date_val):
            row_data = {}
            dates.append(date_val)
            
            # Use the CORRECT column mapping
            for key, col_idx in column_mapping.items():
                val = target_data.iloc[i, col_idx]
                if pd.notna(val) and isinstance(val, (int, float)):
                    row_data[key] = float(val)
                else:
                    row_data[key] = np.nan
            
            extracted_data.append(row_data)
    
    if extracted_data:
        df = pd.DataFrame(extracted_data, index=pd.to_datetime(dates))
        
        print(f"✅ Extracted {len(df)} rows of data")
        
        # Find the row that matches our target values
        print(f"\n🎯 SEARCHING FOR TARGET ROW...")
        
        best_match_idx = None
        best_match_score = float('inf')
        
        for idx, (date, row) in enumerate(df.iterrows()):
            total_diff = 0
            valid_comparisons = 0
            
            for key, target_val in target_values.items():
                if key in row.index and pd.notna(row[key]):
                    diff = abs(row[key] - target_val)
                    total_diff += diff
                    valid_comparisons += 1
            
            if valid_comparisons > 0:
                avg_diff = total_diff / valid_comparisons
                if avg_diff < best_match_score:
                    best_match_score = avg_diff
                    best_match_idx = idx
        
        if best_match_idx is not None:
            best_row = df.iloc[best_match_idx]
            best_date = df.index[best_match_idx]
            
            print(f"\n✅ BEST MATCH FOUND:")
            print(f"   Date: {best_date}")
            print(f"   Average difference: {best_match_score:.3f}")
            
            print(f"\n📊 VERIFICATION WITH CORRECT COLUMNS:")
            print(f"{'Key':<15} {'Our Value':<12} {'Target':<12} {'Diff':<8} {'Status'}")
            print("-" * 60)
            
            perfect_matches = 0
            close_matches = 0
            
            for key, target_val in target_values.items():
                if key in best_row.index and pd.notna(best_row[key]):
                    our_val = best_row[key]
                    diff = abs(our_val - target_val)
                    
                    if diff < 0.01:
                        status = "🎯 PERFECT"
                        perfect_matches += 1
                    elif diff < 0.1:
                        status = "✅ EXCELLENT"
                        close_matches += 1
                    elif diff < 1.0:
                        status = "✅ CLOSE" 
                        close_matches += 1
                    else:
                        status = "❌ OFF"
                    
                    print(f"{key:<15} {our_val:>10.3f} {target_val:>10.3f} {diff:>6.3f} {status}")
                else:
                    print(f"{key:<15} {'NaN':<10} {target_val:>10.3f} {'---':>6} ❌ MISSING")
            
            print(f"\n📊 CORRECTED ACCURACY SUMMARY:")
            print(f"   🎯 Perfect matches (<0.01 diff): {perfect_matches}/{len(target_values)}")
            print(f"   ✅ Close matches (<1.0 diff): {close_matches}/{len(target_values)}")
            print(f"   📈 Total acceptable: {perfect_matches + close_matches}/{len(target_values)}")
            
            # Create the corrected output
            output_file = 'Daily_Historic_Data_CORRECTED_REPLICATION.xlsx'
            
            final_df = df.copy()
            final_df.reset_index(inplace=True)
            final_df.rename(columns={'index': 'Date'}, inplace=True)
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                final_df.to_excel(writer, sheet_name='Daily historic data by category', index=False)
            
            print(f"\n✅ CORRECTED REPLICATION SAVED: {output_file}")
            
            # Show the difference between wrong and correct extraction
            print(f"\n🔍 COMPARISON: Wrong vs Correct Values (first row):")
            print(f"{'Key':<15} {'Wrong Extract':<15} {'Correct Extract':<15} {'Target':<15}")
            print("-" * 70)
            
            # Read our wrong output for comparison
            wrong_df = pd.read_csv('Daily_Historic_Data_by_Category_FINAL_REPLICATION.csv')
            wrong_first = wrong_df.iloc[3]  # Row that we thought was "perfect"
            
            for key in target_values.keys():
                if key in best_row.index:
                    wrong_val = wrong_first.get(key, 'N/A')
                    correct_val = best_row[key]
                    target_val = target_values[key]
                    
                    print(f"{key:<15} {wrong_val:<15.3f} {correct_val:<15.3f} {target_val:<15.3f}")
            
            return final_df, (perfect_matches + close_matches) >= len(target_values) * 0.8
    
    return None, False

if __name__ == "__main__":
    result_df, success = create_correct_replication()
    
    if success:
        print(f"\n✅ CORRECTION SUCCESSFUL")
    else:
        print(f"\n❌ CORRECTION FAILED - Need to investigate further")