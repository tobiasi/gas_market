# -*- coding: utf-8 -*-
"""
Detailed analysis of the target LiveSheet to understand the exact mapping logic
"""

import pandas as pd
import numpy as np

def analyze_target_structure_detailed():
    """Perform detailed analysis of the target structure"""
    
    print("🔬 DETAILED TARGET STRUCTURE ANALYSIS")
    print("=" * 80)
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    # Read the raw data
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    # Find header row (row 10 from previous analysis)
    header_row = 10
    data_start_row = header_row + 1
    
    print(f"📊 Target data shape: {target_data.shape}")
    print(f"📋 Header row: {header_row}")
    print(f"📅 Data starts at row: {data_start_row}")
    
    # Extract headers
    headers = []
    for j in range(target_data.shape[1]):
        val = target_data.iloc[header_row, j]
        headers.append(str(val) if pd.notna(val) else f'Col_{j}')
    
    print(f"\n📋 COLUMN STRUCTURE:")
    for i, header in enumerate(headers[:30]):
        if header not in ['nan', f'Col_{i}']:
            print(f"   Column {i:2d}: {header}")
    
    # Get sample data from multiple rows
    print(f"\n📊 SAMPLE DATA (first 5 rows):")
    sample_dates = []
    sample_data = []
    
    for i in range(data_start_row, min(data_start_row + 5, target_data.shape[0])):
        row_data = {}
        date_val = target_data.iloc[i, 0]
        sample_dates.append(date_val)
        
        # Extract key values
        key_columns = {
            'France': 2, 'Belgium': 3, 'Italy': 4, 'Netherlands': 20, 'GB': 21, 
            'Austria': 28, 'Germany': 8, 'Total': 12, 'Industrial': 13, 
            'LDZ': 14, 'Gas-to-Power': 15
        }
        
        for name, col_idx in key_columns.items():
            if col_idx < target_data.shape[1]:
                val = target_data.iloc[i, col_idx]
                if pd.notna(val) and isinstance(val, (int, float)):
                    row_data[name] = float(val)
        
        sample_data.append(row_data)
        italy_val = row_data.get('Italy', 'N/A')
        total_val = row_data.get('Total', 'N/A')
        italy_str = f"{italy_val:.1f}" if isinstance(italy_val, (int, float)) else str(italy_val)
        total_str = f"{total_val:.1f}" if isinstance(total_val, (int, float)) else str(total_val)
        print(f"   Row {i:2d} ({date_val}): Italy={italy_str}, Total={total_str}")
    
    # Analyze value ranges
    print(f"\n📈 VALUE RANGES ANALYSIS:")
    all_values = {}
    for key in ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Total', 'Industrial', 'LDZ', 'Gas-to-Power']:
        values = [row.get(key) for row in sample_data if row.get(key) is not None]
        if values:
            all_values[key] = {
                'min': min(values),
                'max': max(values),
                'avg': sum(values) / len(values),
                'first': values[0] if values else None
            }
    
    for key, stats in all_values.items():
        print(f"   {key:<12}: {stats['min']:6.1f} - {stats['max']:6.1f} (avg: {stats['avg']:6.1f})")
    
    # Check for patterns in other columns
    print(f"\n🔍 SCANNING ALL COLUMNS FOR PATTERNS:")
    interesting_cols = []
    
    for col_idx in range(target_data.shape[1]):
        header = headers[col_idx]
        if header not in ['nan', f'Col_{col_idx}'] and col_idx not in [2,3,4,8,12,13,14,15,20,21,28]:
            # Get sample values from this column
            sample_vals = []
            for row_idx in range(data_start_row, min(data_start_row + 5, target_data.shape[0])):
                val = target_data.iloc[row_idx, col_idx]
                if pd.notna(val) and isinstance(val, (int, float)) and abs(val) > 1:
                    sample_vals.append(val)
            
            if sample_vals:
                avg_val = sum(sample_vals) / len(sample_vals)
                interesting_cols.append((col_idx, header, avg_val, sample_vals[0]))
    
    # Sort by average value to find potentially important columns
    interesting_cols.sort(key=lambda x: abs(x[2]), reverse=True)
    
    print("   Top non-key columns with significant values:")
    for col_idx, header, avg_val, first_val in interesting_cols[:15]:
        print(f"   Column {col_idx:2d} ({header[:20]:20s}): avg={avg_val:8.2f}, first={first_val:8.2f}")
    
    # Save detailed analysis
    with open('target_structure_analysis.txt', 'w') as f:
        f.write("DETAILED TARGET STRUCTURE ANALYSIS\n")
        f.write("=" * 50 + "\n\n")
        
        f.write("KEY COLUMNS:\n")
        key_columns = {
            'France': 2, 'Belgium': 3, 'Italy': 4, 'Netherlands': 20, 'GB': 21, 
            'Austria': 28, 'Germany': 8, 'Total': 12, 'Industrial': 13, 
            'LDZ': 14, 'Gas-to-Power': 15
        }
        for name, col in key_columns.items():
            f.write(f"{name}: Column {col}\n")
        
        f.write(f"\nVALUE RANGES:\n")
        for key, stats in all_values.items():
            f.write(f"{key}: {stats['min']:.1f} - {stats['max']:.1f} (avg: {stats['avg']:.1f}, first: {stats['first']:.1f})\n")
        
        f.write(f"\nINTERESTING COLUMNS:\n")
        for col_idx, header, avg_val, first_val in interesting_cols[:20]:
            f.write(f"Column {col_idx}: {header} (avg={avg_val:.2f}, first={first_val:.2f})\n")
    
    print(f"\n✅ Detailed analysis saved to: target_structure_analysis.txt")
    
    return all_values, interesting_cols

if __name__ == "__main__":
    analyze_target_structure_detailed()