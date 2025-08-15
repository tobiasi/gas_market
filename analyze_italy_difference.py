# -*- coding: utf-8 -*-
"""
ANALYSIS: Compare Italy demand differences between Excel and Python output
"""

import pandas as pd
import numpy as np

def analyze_italy_differences():
    """Compare Italy demand values between target and our output"""
    
    print(f"\n{'='*80}")
    print("🇮🇹 ITALY DEMAND COMPARISON ANALYSIS")
    print(f"{'='*80}")
    
    # Read the CSV files
    try:
        target = pd.read_csv('The_daily_we_want_to_replicate.csv')
        our_output = pd.read_csv('our_daily.csv')
        
        print("📊 Files loaded successfully")
        print(f"   Target shape: {target.shape}")
        print(f"   Our output shape: {our_output.shape}")
        
    except Exception as e:
        print(f"❌ Error loading files: {e}")
        return
    
    # Find Italy columns (should be column index 2 based on headers)
    print(f"\n🔍 ITALY COLUMN ANALYSIS:")
    
    # Get Italy data from both files
    # Skip the header rows and get the actual data
    target_italy = target.iloc[3:, 2]  # Column 2 = Italy, skip first 3 header rows
    our_italy = our_output.iloc[4:, 2]   # Column 2 = Italy, skip first 4 header rows
    
    # Convert to numeric, handling any non-numeric values
    target_italy = pd.to_numeric(target_italy, errors='coerce')
    our_italy = pd.to_numeric(our_italy, errors='coerce')
    
    # Get common length for comparison
    min_length = min(len(target_italy), len(our_italy))
    target_italy = target_italy.iloc[:min_length]
    our_italy = our_italy.iloc[:min_length]
    
    print(f"   Comparing {min_length} data points")
    print(f"   Target Italy sample: {target_italy.iloc[:5].values}")
    print(f"   Our Italy sample: {our_italy.iloc[:5].values}")
    
    # Calculate differences
    differences = our_italy - target_italy
    abs_differences = abs(differences)
    
    print(f"\n📈 DIFFERENCE STATISTICS:")
    print(f"   Mean difference: {differences.mean():.6f}")
    print(f"   Mean absolute difference: {abs_differences.mean():.6f}")
    print(f"   Max difference: {differences.max():.6f}")
    print(f"   Min difference: {differences.min():.6f}")
    print(f"   Standard deviation: {differences.std():.6f}")
    
    # Find the biggest differences
    print(f"\n🎯 LARGEST DISCREPANCIES:")
    largest_diff_indices = abs_differences.nlargest(10).index
    
    print(f"{'Row':<5} {'Target':<12} {'Ours':<12} {'Difference':<12} {'% Error':<10}")
    print("-" * 55)
    
    for idx in largest_diff_indices:
        target_val = target_italy.iloc[idx]
        our_val = our_italy.iloc[idx]
        diff = our_val - target_val
        pct_error = (diff / target_val * 100) if target_val != 0 else 0
        
        print(f"{idx:<5} {target_val:<12.2f} {our_val:<12.2f} {diff:<12.2f} {pct_error:<10.2f}%")
    
    # Check for patterns
    print(f"\n🔍 PATTERN ANALYSIS:")
    
    # Check if we're consistently higher or lower
    positive_diffs = (differences > 0).sum()
    negative_diffs = (differences < 0).sum()
    zero_diffs = (differences == 0).sum()
    
    print(f"   We're higher than target: {positive_diffs} times")
    print(f"   We're lower than target: {negative_diffs} times") 
    print(f"   Exact matches: {zero_diffs} times")
    
    # Check if differences are consistent (systematic) or random
    if abs_differences.std() < abs_differences.mean():
        print(f"   Pattern: Differences appear SYSTEMATIC (consistent offset)")
    else:
        print(f"   Pattern: Differences appear RANDOM (varying offset)")
    
    # Show date-based analysis if we can extract dates
    print(f"\n📅 SAMPLE COMPARISON (first 10 rows):")
    print(f"{'Index':<8} {'Target':<12} {'Ours':<12} {'Diff':<10}")
    print("-" * 45)
    
    for i in range(min(10, len(target_italy))):
        target_val = target_italy.iloc[i]
        our_val = our_italy.iloc[i]
        diff = our_val - target_val
        print(f"{i:<8} {target_val:<12.2f} {our_val:<12.2f} {diff:<10.2f}")
    
    print(f"\n💡 ANALYSIS SUMMARY:")
    avg_error = abs_differences.mean()
    if avg_error < 0.1:
        print(f"   ✅ Very small differences (avg {avg_error:.3f}) - likely rounding")
    elif avg_error < 1.0:
        print(f"   ⚠️  Small differences (avg {avg_error:.3f}) - minor data issue")
    elif avg_error < 10.0:
        print(f"   ❌ Medium differences (avg {avg_error:.3f}) - missing some data")
    else:
        print(f"   🚨 Large differences (avg {avg_error:.3f}) - major calculation error")
    
    return differences

if __name__ == "__main__":
    differences = analyze_italy_differences()