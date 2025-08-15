# -*- coding: utf-8 -*-
"""
Deep analysis of LiveSheet target structure
"""

import pandas as pd
import numpy as np

def deep_analyze_livesheet():
    """Deep analysis of the target structure"""
    
    print(f"\n{'='*80}")
    print("🔬 DEEP ANALYSIS OF LIVESHEET TARGET")
    print(f"{'='*80}")
    
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    try:
        # Read without any processing
        target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
        print(f"📊 Shape: {target_data.shape}")
        
        # Search more rows for headers
        print(f"\n🔍 SEARCHING FOR HEADERS (first 20 rows):")
        
        countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany']
        categories = ['Total', 'Industrial', 'LDZ', 'Gas-to-Power']
        
        header_candidates = []
        
        for i in range(min(20, target_data.shape[0])):
            row_text = []
            for j in range(min(30, target_data.shape[1])):
                val = target_data.iloc[i, j]
                if pd.notna(val):
                    row_text.append(str(val))
                else:
                    row_text.append('')
            
            # Check if this row contains country names
            country_count = sum(1 for country in countries if any(country in cell for cell in row_text))
            category_count = sum(1 for cat in categories if any(cat in cell for cell in row_text))
            
            if country_count >= 3 or category_count >= 2:  # Likely header row
                header_candidates.append((i, country_count, category_count, row_text))
                print(f"   Row {i:2d}: {country_count} countries, {category_count} categories")
                print(f"          {row_text[:15]}")
        
        if not header_candidates:
            print("❌ No clear header rows found")
            # Let's look at more rows
            print(f"\n📋 RAW DATA SAMPLE (rows 5-15, cols 0-10):")
            for i in range(5, min(15, target_data.shape[0])):
                row_sample = []
                for j in range(min(10, target_data.shape[1])):
                    val = target_data.iloc[i, j]
                    if pd.isna(val):
                        row_sample.append('NaN')
                    elif isinstance(val, (int, float)):
                        row_sample.append(f'{val:.1f}')
                    else:
                        row_sample.append(str(val)[:8])
                print(f"   Row {i:2d}: {row_sample}")
            return None
        
        # Use the best header candidate
        best_header = max(header_candidates, key=lambda x: x[1] + x[2])
        header_row_idx, country_count, category_count, header_text = best_header
        
        print(f"\n✅ BEST HEADER ROW: {header_row_idx}")
        print(f"   Countries found: {country_count}")
        print(f"   Categories found: {category_count}")
        
        # Extract column mapping
        column_mapping = {}
        for j, cell in enumerate(header_text):
            if cell and any(keyword in cell for keyword in countries + categories):
                column_mapping[cell.strip()] = j
                
        print(f"\n📋 COLUMN MAPPING:")
        for name, col_idx in sorted(column_mapping.items()):
            print(f"   {name:<12}: Column {col_idx}")
        
        # Get sample data
        data_start_row = header_row_idx + 1
        if data_start_row < target_data.shape[0]:
            print(f"\n🎯 SAMPLE TARGET DATA:")
            sample_data = {}
            
            for i in range(data_start_row, min(data_start_row + 5, target_data.shape[0])):
                date_val = target_data.iloc[i, 0]
                row_data = {}
                
                for name, col_idx in column_mapping.items():
                    if col_idx < target_data.shape[1]:
                        val = target_data.iloc[i, col_idx]
                        if pd.notna(val) and isinstance(val, (int, float)):
                            row_data[name] = float(val)
                
                sample_data[str(date_val)] = row_data
                print(f"   {date_val}: {row_data}")
            
            # Save first row as target
            if sample_data:
                first_key = list(sample_data.keys())[0]
                target_values = sample_data[first_key]
                
                print(f"\n🎯 TARGET VALUES TO MATCH:")
                for key, val in target_values.items():
                    print(f"   {key}: {val:.6f}")
                
                return target_data, header_row_idx, column_mapping, target_values
        
        return None
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return None

def find_bloomberg_files():
    """Find any files that might contain Bloomberg data"""
    
    print(f"\n{'='*80}")
    print("🔍 SEARCHING FOR BLOOMBERG DATA FILES")
    print(f"{'='*80}")
    
    import os
    
    # Look for various patterns
    patterns = ['bloomberg', 'raw', 'data', 'ticker', 'multiticker']
    found_files = []
    
    for file in os.listdir('.'):
        if file.endswith(('.csv', '.xlsx')):
            file_lower = file.lower()
            for pattern in patterns:
                if pattern in file_lower:
                    size = os.path.getsize(file) / 1024 / 1024  # MB
                    found_files.append((file, size, pattern))
                    break
    
    if found_files:
        print("📁 Potential data files found:")
        for file, size, pattern in sorted(found_files, key=lambda x: x[1], reverse=True):
            print(f"   {file} ({size:.1f} MB) - matches '{pattern}'")
        
        # Check if we have our_multiticker.csv (might contain raw data)
        if 'our_multiticker.csv' in [f[0] for f in found_files]:
            print(f"\n📊 Found our_multiticker.csv - checking if it contains raw data...")
            try:
                data = pd.read_csv('our_multiticker.csv', nrows=5)
                print(f"   Shape: {data.shape}")
                print(f"   Sample columns: {list(data.columns[:10])}")
                return 'our_multiticker.csv'
            except Exception as e:
                print(f"   Error reading: {e}")
        
        # Return the largest file
        largest_file = max(found_files, key=lambda x: x[1])[0]
        print(f"\n✅ Recommended file: {largest_file}")
        return largest_file
    
    else:
        print("❌ No potential Bloomberg data files found")
        print("Please upload the raw Bloomberg data file")
        return None

if __name__ == "__main__":
    # Deep analyze target
    target_result = deep_analyze_livesheet()
    
    # Find Bloomberg data
    bloomberg_file = find_bloomberg_files()
    
    if target_result and bloomberg_file:
        target_data, header_row, column_mapping, target_values = target_result
        
        print(f"\n🎯 ANALYSIS COMPLETE:")
        print(f"   ✅ Target structure understood")
        print(f"   ✅ Bloomberg data file found: {bloomberg_file}")
        print(f"   📊 Ready to process and match target values")
        
        # Save results for next step
        import json
        with open('analysis_results.json', 'w') as f:
            json.dump({
                'target_values': target_values,
                'column_mapping': column_mapping,
                'header_row': header_row,
                'bloomberg_file': bloomberg_file
            }, f, indent=2)
        
        print(f"   💾 Results saved to: analysis_results.json")
        
    else:
        print(f"\n❌ ANALYSIS INCOMPLETE")
        print("   Need both target structure and Bloomberg data to proceed")