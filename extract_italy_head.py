# -*- coding: utf-8 -*-
"""
Extract and display the head of Italy column from both sources
"""

import pandas as pd
import json

def extract_italy_head():
    """Extract Italy column head from both sources"""
    
    print("🇮🇹 ITALY COLUMN HEAD EXTRACTION")
    print("=" * 80)
    
    # 1. Original LiveSheet Italy Column
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    
    print("📊 ORIGINAL LIVESHEET - Italy Column (Column 4):")
    print("Date         Italy Value")
    print("-" * 25)
    
    data_start_row = 12
    original_italy = []
    
    for i in range(data_start_row, min(data_start_row + 10, target_data.shape[0])):
        date_val = target_data.iloc[i, 1]
        italy_val = target_data.iloc[i, 4]
        
        if pd.notna(date_val) and pd.notna(italy_val):
            date_str = str(date_val)[:10]
            print(f"{date_str}   {italy_val:.6f}")
            original_italy.append(italy_val)
    
    # 2. Our Corrected Output
    print(f"\n📊 OUR CORRECTED OUTPUT - Italy Column:")
    
    with open('analysis_results.json', 'r') as f:
        analysis = json.load(f)
    
    column_mapping = analysis['column_mapping']
    
    print("Date         Italy Value")  
    print("-" * 25)
    
    our_italy = []
    
    for i in range(data_start_row, min(data_start_row + 10, target_data.shape[0])):
        date_val = target_data.iloc[i, 1]
        italy_val = target_data.iloc[i, column_mapping['Italy']]
        
        if pd.notna(date_val) and pd.notna(italy_val):
            date_str = str(date_val)[:10]
            print(f"{date_str}   {italy_val:.6f}")
            our_italy.append(italy_val)
    
    # 3. Side-by-side comparison
    print(f"\n📊 SIDE-BY-SIDE COMPARISON:")
    print(f"{'Date':<12} {'Original':<15} {'Our Output':<15} {'Difference':<12}")
    print("-" * 60)
    
    for i in range(min(len(original_italy), len(our_italy))):
        date_val = target_data.iloc[data_start_row + i, 1]
        date_str = str(date_val)[:10] if pd.notna(date_val) else "N/A"
        
        orig_val = original_italy[i]
        our_val = our_italy[i]
        diff = abs(orig_val - our_val)
        
        print(f"{date_str:<12} {orig_val:<15.6f} {our_val:<15.6f} {diff:<12.6f}")
    
    # 4. Statistical summary
    print(f"\n📈 STATISTICAL SUMMARY:")
    print(f"   Total rows compared: {min(len(original_italy), len(our_italy))}")
    print(f"   Perfect matches (diff < 0.000001): {sum(1 for i in range(min(len(original_italy), len(our_italy))) if abs(original_italy[i] - our_italy[i]) < 0.000001)}")
    print(f"   Maximum difference: {max(abs(original_italy[i] - our_italy[i]) for i in range(min(len(original_italy), len(our_italy)))):.10f}")
    
    return original_italy, our_italy

if __name__ == "__main__":
    orig_italy, our_italy = extract_italy_head()