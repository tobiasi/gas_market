# -*- coding: utf-8 -*-
"""
DETAILED ANALYSIS: Examine Italy data structure and find exact discrepancies
"""

import pandas as pd
import numpy as np

def detailed_italy_analysis():
    """Detailed analysis of Italy data structure and discrepancies"""
    
    print(f"\n{'='*80}")
    print("🔍 DETAILED ITALY STRUCTURE ANALYSIS")
    print(f"{'='*80}")
    
    # Read files with different approaches to understand structure
    try:
        target = pd.read_csv('The_daily_we_want_to_replicate.csv')
        our_output = pd.read_csv('our_daily.csv')
        
        print("📋 FILE STRUCTURE ANALYSIS:")
        print(f"\n🎯 TARGET FILE STRUCTURE:")
        print(f"   Shape: {target.shape}")
        print(f"   First 5 rows, first 5 columns:")
        print(target.iloc[:5, :5])
        
        print(f"\n🔧 OUR OUTPUT STRUCTURE:")
        print(f"   Shape: {our_output.shape}")
        print(f"   First 5 rows, first 5 columns:")
        print(our_output.iloc[:5, :5])
        
        # Find the Italy column in each file by looking at headers
        print(f"\n🇮🇹 ITALY COLUMN IDENTIFICATION:")
        
        # Look for Italy in target file headers
        italy_col_target = None
        for row_idx in range(min(5, target.shape[0])):
            row = target.iloc[row_idx]
            for col_idx, val in enumerate(row):
                if str(val).strip().lower() == 'italy':
                    italy_col_target = col_idx
                    print(f"   Target: Found 'Italy' at row {row_idx}, column {col_idx}")
                    break
            if italy_col_target is not None:
                break
        
        # Look for Italy in our output headers
        italy_col_ours = None
        for row_idx in range(min(5, our_output.shape[0])):
            row = our_output.iloc[row_idx]
            for col_idx, val in enumerate(row):
                if str(val).strip().lower() == 'italy':
                    italy_col_ours = col_idx
                    print(f"   Ours: Found 'Italy' at row {row_idx}, column {col_idx}")
                    break
            if italy_col_ours is not None:
                break
        
        if italy_col_target is None or italy_col_ours is None:
            print("❌ Could not find Italy column in one or both files")
            return
        
        # Extract Italy data with proper date alignment
        print(f"\n📊 ITALY DATA EXTRACTION:")
        
        # Find where actual data starts (look for first numeric row)
        target_data_start = None
        for i in range(target.shape[0]):
            try:
                val = pd.to_numeric(target.iloc[i, italy_col_target], errors='coerce')
                if not pd.isna(val):
                    target_data_start = i
                    print(f"   Target data starts at row {i}")
                    break
            except:
                continue
        
        our_data_start = None
        for i in range(our_output.shape[0]):
            try:
                val = pd.to_numeric(our_output.iloc[i, italy_col_ours], errors='coerce')
                if not pd.isna(val):
                    our_data_start = i
                    print(f"   Our data starts at row {i}")
                    break
            except:
                continue
        
        if target_data_start is None or our_data_start is None:
            print("❌ Could not find numeric data start in one or both files")
            return
        
        # Extract the actual Italy data
        target_italy = target.iloc[target_data_start:, italy_col_target]
        our_italy = our_output.iloc[our_data_start:, italy_col_ours]
        
        # Convert to numeric
        target_italy = pd.to_numeric(target_italy, errors='coerce').dropna()
        our_italy = pd.to_numeric(our_italy, errors='coerce').dropna()
        
        print(f"   Target Italy data points: {len(target_italy)}")
        print(f"   Our Italy data points: {len(our_italy)}")
        
        # Compare first 20 values with detailed output
        print(f"\n📈 DETAILED VALUE COMPARISON (first 20 points):")
        print(f"{'Index':<6} {'Target':<12} {'Ours':<12} {'Difference':<12} {'% Error':<10}")
        print("-" * 65)
        
        comparison_length = min(20, len(target_italy), len(our_italy))
        
        total_abs_diff = 0
        for i in range(comparison_length):
            target_val = target_italy.iloc[i]
            our_val = our_italy.iloc[i] 
            diff = our_val - target_val
            abs_diff = abs(diff)
            total_abs_diff += abs_diff
            pct_error = (diff / target_val * 100) if target_val != 0 else 0
            
            status = "✅" if abs_diff < 0.01 else "⚠️" if abs_diff < 1.0 else "❌"
            print(f"{i:<6} {target_val:<12.6f} {our_val:<12.6f} {diff:<12.6f} {pct_error:<10.2f}% {status}")
        
        avg_abs_diff = total_abs_diff / comparison_length
        print(f"\nAverage absolute difference in first 20 points: {avg_abs_diff:.6f}")
        
        # Look for date columns to understand time alignment
        print(f"\n📅 DATE COLUMN ANALYSIS:")
        
        # Check if first column contains dates
        target_dates = target.iloc[target_data_start:target_data_start+5, 0]
        our_dates = our_output.iloc[our_data_start:our_data_start+5, 0]
        
        print(f"   Target first column (first 5): {target_dates.values}")
        print(f"   Our first column (first 5): {our_dates.values}")
        
        # Check if these look like Excel date numbers
        try:
            target_first_date = pd.to_numeric(target_dates.iloc[0])
            our_first_date = pd.to_numeric(our_dates.iloc[0])
            
            if 40000 < target_first_date < 50000:  # Excel date range
                print(f"   Target appears to use Excel date format: {target_first_date}")
            if 40000 < our_first_date < 50000:
                print(f"   Our output appears to use Excel date format: {our_first_date}")
                
            date_diff = our_first_date - target_first_date
            print(f"   Date offset: {date_diff} days")
            
        except:
            print("   Could not parse dates as numbers")
        
        return target_italy, our_italy
        
    except Exception as e:
        print(f"❌ Error in analysis: {e}")
        import traceback
        traceback.print_exc()
        return None, None

if __name__ == "__main__":
    target_italy, our_italy = detailed_italy_analysis()