#!/usr/bin/env python3
"""
Compare LiveSheet data with code output to verify no discrepancies.
"""

import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def load_livesheet_data(file_path='use4.xlsx', sheet_name='Daily historic data by category'):
    """
    Load the LiveSheet data (Daily historic data by category).
    """
    print(f"Loading LiveSheet data from {sheet_name}...")
    
    # Load the sheet with more control
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # The structure is:
    # Row 11 (index 10): Column headers
    # Row 12+ (index 11+): Data
    
    # Get headers from row 11, starting from column C (index 2)
    headers = []
    for col_idx in range(2, min(df.shape[1], 50)):  # Check up to column AX
        val = df.iloc[10, col_idx]
        if pd.notna(val):
            headers.append(str(val))
        else:
            break
    
    print(f"Found {len(headers)} columns: {headers[:5]}...")
    
    # Extract data starting from row 12 (index 11)
    # Column B (index 1) has dates, columns C onwards have data
    data_rows = []
    for row_idx in range(11, df.shape[0]):
        date_val = df.iloc[row_idx, 1]  # Column B
        if pd.isna(date_val):
            continue
        
        row_data = {'Date': date_val}
        for i, header in enumerate(headers):
            if i + 2 < df.shape[1]:  # Column index offset
                row_data[header] = df.iloc[row_idx, i + 2]
        
        data_rows.append(row_data)
    
    # Create DataFrame
    data = pd.DataFrame(data_rows)
    
    # Convert Date column
    data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
    
    # Remove invalid date rows
    data = data.dropna(subset=['Date'])
    
    # Convert numeric columns
    for col in data.columns:
        if col != 'Date':
            data[col] = pd.to_numeric(data[col], errors='coerce')
    
    print(f"Loaded {len(data)} rows from LiveSheet")
    
    return data


def load_code_output(file_path='daily_historic_data_by_category_output.csv'):
    """
    Load the code-generated output.
    """
    print(f"Loading code output from {file_path}...")
    df = pd.read_csv(file_path)
    df['Date'] = pd.to_datetime(df['Date'])
    print(f"Loaded {len(df)} rows from code output")
    return df


def compare_dataframes(livesheet_df, code_df, tolerance=0.01):
    """
    Compare two dataframes and identify discrepancies.
    """
    print("\n" + "="*60)
    print("COMPARING LIVESHEET VS CODE OUTPUT")
    print("="*60)
    
    # Merge on date
    merged = pd.merge(
        livesheet_df,
        code_df,
        on='Date',
        how='inner',
        suffixes=('_live', '_code')
    )
    
    print(f"\nMatched {len(merged)} dates for comparison")
    
    # Identify common columns to compare
    live_cols = set(livesheet_df.columns) - {'Date'}
    code_cols = set(code_df.columns) - {'Date'}
    common_cols = live_cols.intersection(code_cols)
    
    print(f"Common columns to compare: {sorted(common_cols)}")
    
    # Track discrepancies
    discrepancies = []
    summary = {
        'total_comparisons': 0,
        'exact_matches': 0,
        'within_tolerance': 0,
        'discrepancies': 0,
        'max_diff': 0,
        'max_diff_info': None
    }
    
    # Compare each common column
    for col in sorted(common_cols):
        live_col = f"{col}_live" if f"{col}_live" in merged.columns else col
        code_col = f"{col}_code" if f"{col}_code" in merged.columns else col
        
        if live_col not in merged.columns or code_col not in merged.columns:
            continue
        
        # Calculate differences
        diff = (merged[code_col] - merged[live_col]).abs()
        
        # Remove NaN comparisons
        valid_mask = ~(merged[live_col].isna() | merged[code_col].isna())
        diff = diff[valid_mask]
        merged_valid = merged[valid_mask]
        
        if len(diff) == 0:
            continue
        
        summary['total_comparisons'] += len(diff)
        
        # Count exact matches
        exact_matches = (diff == 0).sum()
        summary['exact_matches'] += exact_matches
        
        # Count within tolerance
        within_tol = (diff <= tolerance).sum()
        summary['within_tolerance'] += within_tol - exact_matches
        
        # Find discrepancies
        disc_mask = diff > tolerance
        if disc_mask.any():
            disc_dates = merged_valid.loc[disc_mask, 'Date']
            disc_live = merged_valid.loc[disc_mask, live_col]
            disc_code = merged_valid.loc[disc_mask, code_col]
            disc_diff = diff[disc_mask]
            
            for i in range(len(disc_dates)):
                discrepancy = {
                    'date': disc_dates.iloc[i],
                    'column': col,
                    'livesheet': disc_live.iloc[i],
                    'code': disc_code.iloc[i],
                    'difference': disc_diff.iloc[i]
                }
                discrepancies.append(discrepancy)
                
                # Track max difference
                if disc_diff.iloc[i] > summary['max_diff']:
                    summary['max_diff'] = disc_diff.iloc[i]
                    summary['max_diff_info'] = discrepancy
        
        # Report column summary
        max_diff_col = diff.max()
        mean_diff_col = diff.mean()
        
        if max_diff_col > tolerance:
            print(f"\n⚠️  {col}:")
            print(f"   Max difference: {max_diff_col:.4f}")
            print(f"   Mean difference: {mean_diff_col:.4f}")
            print(f"   Discrepancies: {disc_mask.sum()}/{len(diff)}")
        else:
            print(f"✓ {col}: Perfect match (max diff: {max_diff_col:.6f})")
    
    summary['discrepancies'] = len(discrepancies)
    
    return discrepancies, summary, merged


def analyze_specific_dates(merged_df, dates_to_check):
    """
    Analyze specific dates in detail.
    """
    print("\n" + "="*60)
    print("SPECIFIC DATE ANALYSIS")
    print("="*60)
    
    for date in dates_to_check:
        date_data = merged_df[merged_df['Date'] == pd.to_datetime(date)]
        
        if date_data.empty:
            print(f"\n{date}: No data found")
            continue
        
        print(f"\n{date}:")
        
        # Check key columns
        key_cols = ['France', 'Belgium', 'Italy', 'Germany', 'Total']
        
        for col in key_cols:
            live_col = f"{col}_live" if f"{col}_live" in date_data.columns else col
            code_col = f"{col}_code" if f"{col}_code" in date_data.columns else col
            
            if live_col in date_data.columns and code_col in date_data.columns:
                live_val = date_data[live_col].iloc[0]
                code_val = date_data[code_col].iloc[0]
                diff = abs(code_val - live_val) if not pd.isna(live_val) and not pd.isna(code_val) else np.nan
                
                if pd.isna(diff):
                    print(f"  {col}: LiveSheet={live_val:.2f}, Code={code_val:.2f} (one is NaN)")
                elif diff < 0.01:
                    print(f"  {col}: ✓ Match ({code_val:.2f})")
                else:
                    print(f"  {col}: ⚠️  LiveSheet={live_val:.2f}, Code={code_val:.2f}, Diff={diff:.4f}")


def create_discrepancy_report(discrepancies, summary):
    """
    Create a detailed discrepancy report.
    """
    print("\n" + "="*60)
    print("DISCREPANCY REPORT SUMMARY")
    print("="*60)
    
    print(f"\nTotal comparisons: {summary['total_comparisons']:,}")
    print(f"Exact matches: {summary['exact_matches']:,} ({summary['exact_matches']/summary['total_comparisons']*100:.2f}%)")
    print(f"Within tolerance (±0.01): {summary['within_tolerance']:,} ({summary['within_tolerance']/summary['total_comparisons']*100:.2f}%)")
    print(f"Discrepancies: {summary['discrepancies']:,} ({summary['discrepancies']/summary['total_comparisons']*100:.2f}%)")
    
    if summary['max_diff_info']:
        print(f"\nLargest discrepancy:")
        info = summary['max_diff_info']
        print(f"  Date: {info['date'].strftime('%Y-%m-%d')}")
        print(f"  Column: {info['column']}")
        print(f"  LiveSheet: {info['livesheet']:.4f}")
        print(f"  Code: {info['code']:.4f}")
        print(f"  Difference: {info['difference']:.4f}")
    
    if discrepancies:
        # Group discrepancies by column
        by_column = {}
        for disc in discrepancies:
            col = disc['column']
            if col not in by_column:
                by_column[col] = []
            by_column[col].append(disc)
        
        print("\n" + "-"*40)
        print("DISCREPANCIES BY COLUMN:")
        print("-"*40)
        
        for col, discs in sorted(by_column.items()):
            print(f"\n{col}: {len(discs)} discrepancies")
            # Show first 3 examples
            for disc in discs[:3]:
                print(f"  {disc['date'].strftime('%Y-%m-%d')}: Live={disc['livesheet']:.2f}, Code={disc['code']:.2f}, Diff={disc['difference']:.2f}")
            if len(discs) > 3:
                print(f"  ... and {len(discs)-3} more")
    
    # Save discrepancies to CSV
    if discrepancies:
        disc_df = pd.DataFrame(discrepancies)
        disc_df.to_csv('discrepancy_report.csv', index=False)
        print(f"\nDetailed discrepancy report saved to: discrepancy_report.csv")
    
    return summary


def main():
    """
    Main comparison function.
    """
    try:
        # Load LiveSheet data
        livesheet_df = load_livesheet_data('use4.xlsx', 'Daily historic data by category')
        
        # Load code output
        code_df = load_code_output('daily_historic_data_by_category_output.csv')
        
        # Compare dataframes
        discrepancies, summary, merged = compare_dataframes(livesheet_df, code_df)
        
        # Analyze specific validation dates
        validation_dates = ['2016-10-03', '2016-10-04', '2016-10-05']
        analyze_specific_dates(merged, validation_dates)
        
        # Create discrepancy report
        create_discrepancy_report(discrepancies, summary)
        
        # Final verdict
        print("\n" + "="*60)
        print("FINAL VERDICT")
        print("="*60)
        
        accuracy = (summary['exact_matches'] + summary['within_tolerance']) / summary['total_comparisons'] * 100
        
        if accuracy >= 99.99:
            print("✅ EXCELLENT: Code output matches LiveSheet with >99.99% accuracy")
        elif accuracy >= 99.9:
            print("✅ VERY GOOD: Code output matches LiveSheet with >99.9% accuracy")
        elif accuracy >= 99:
            print("⚠️  GOOD: Code output matches LiveSheet with >99% accuracy")
        else:
            print(f"❌ NEEDS REVIEW: Code output matches LiveSheet with {accuracy:.2f}% accuracy")
        
        print(f"\nOverall accuracy: {accuracy:.4f}%")
        
        return summary
        
    except Exception as e:
        print(f"Error during comparison: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    summary = main()