"""
DEBUG SCRIPT: Investigate why Demand tab sums don't add up
"""
import warnings
import pandas as pd
warnings.filterwarnings('ignore')

# This should be run after the main script has completed
# Add this code at the end of gas_market_consolidated_use4.py to debug

def debug_demand_sums(countries, industry, ldz, gtp):
    """Debug why sums don't add up in demand calculations"""
    
    print(f"\n{'='*80}")
    print("🔍 DEBUGGING DEMAND SUMS")
    print(f"{'='*80}")
    
    # Get the data for analysis
    total_col = countries[('','','Total')]
    industrial_col = countries[('','','Industrial')]
    ldz_col = countries[('','','LDZ')]
    gtp_col = countries[('','','Gas-to-Power')]
    
    # Calculate the differences
    manual_sum = industrial_col + ldz_col + gtp_col
    difference = total_col - manual_sum
    
    # Find where differences are significant
    significant_diff = abs(difference) > 0.1  # Threshold for "significant"
    
    print(f"📊 DATA OVERVIEW:")
    print(f"   Total column range: {total_col.min():.2f} to {total_col.max():.2f}")
    print(f"   Industrial range: {industrial_col.min():.2f} to {industrial_col.max():.2f}")
    print(f"   LDZ range: {ldz_col.min():.2f} to {ldz_col.max():.2f}")
    print(f"   Gas-to-Power range: {gtp_col.min():.2f} to {gtp_col.max():.2f}")
    print(f"   Manual sum range: {manual_sum.min():.2f} to {manual_sum.max():.2f}")
    
    print(f"\n📈 DIFFERENCE ANALYSIS:")
    print(f"   Mean difference: {difference.mean():.4f}")
    print(f"   Max difference: {difference.max():.4f}")
    print(f"   Min difference: {difference.min():.4f}")
    print(f"   Std deviation: {difference.std():.4f}")
    print(f"   Days with significant diff (>{0.1}): {significant_diff.sum()}")
    
    if significant_diff.sum() > 0:
        print(f"\n⚠️  SIGNIFICANT DIFFERENCES FOUND:")
        print(f"   Sample dates with large differences:")
        
        # Show worst cases
        worst_indices = difference.abs().nlargest(5).index
        print(f"{'Date':<12} {'Total':<10} {'Ind':<8} {'LDZ':<8} {'GTP':<8} {'Sum':<10} {'Diff':<8}")
        print("-" * 70)
        
        for idx in worst_indices:
            date_str = str(idx)[:10]
            total_val = total_col[idx]
            ind_val = industrial_col[idx] 
            ldz_val = ldz_col[idx]
            gtp_val = gtp_col[idx]
            sum_val = manual_sum[idx]
            diff_val = difference[idx]
            
            print(f"{date_str:<12} {total_val:<10.2f} {ind_val:<8.2f} {ldz_val:<8.2f} {gtp_val:<8.2f} {sum_val:<10.2f} {diff_val:<8.2f}")
    
    # Check for NaN values
    print(f"\n🔍 NaN VALUE ANALYSIS:")
    print(f"   Total NaNs: {total_col.isna().sum()}")
    print(f"   Industrial NaNs: {industrial_col.isna().sum()}")
    print(f"   LDZ NaNs: {ldz_col.isna().sum()}")
    print(f"   Gas-to-Power NaNs: {gtp_col.isna().sum()}")
    
    # Check individual country sums vs category sums
    print(f"\n🏗️  CONSTRUCTION ANALYSIS:")
    print("   How Total is calculated:")
    print("   - From individual countries (France, Belgium, Italy, etc.)")
    print("   How categories are calculated:")
    print("   - Industrial: from industry DataFrame aggregation")
    print("   - LDZ: from ldz DataFrame aggregation") 
    print("   - Gas-to-Power: from gtp DataFrame aggregation")
    
    # Investigate source DataFrames
    print(f"\n🔧 SOURCE DATAFRAME ANALYSIS:")
    print(f"   Industry DataFrame shape: {industry.shape}")
    print(f"   LDZ DataFrame shape: {ldz.shape}")
    print(f"   GTP DataFrame shape: {gtp.shape}")
    
    # Check if industry + ldz + gtp totals match the sum of their constituent parts
    industry_total = industry.iloc[:,industry.columns.get_level_values(2)=='Total'].iloc[:,0]
    ldz_total = ldz.iloc[:,ldz.columns.get_level_values(2)=='Total'].iloc[:,0]
    gtp_total = gtp.iloc[:,gtp.columns.get_level_values(2)=='Total'].iloc[:,0]
    
    print(f"\n📊 CATEGORY TOTAL VERIFICATION:")
    print(f"   Industry total sample: {industry_total.iloc[0:3].values}")
    print(f"   LDZ total sample: {ldz_total.iloc[0:3].values}")
    print(f"   GTP total sample: {gtp_total.iloc[0:3].values}")
    
    return difference

# To use this, add at the end of your main script:
# debug_difference = debug_demand_sums(countries, industry, ldz, gtp)

print("🔍 Add this debug function to your main script to investigate the sum discrepancies!")