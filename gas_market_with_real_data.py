# -*- coding: utf-8 -*-
"""
Gas Market Processor - Uses Real Downloaded Bloomberg Data
This version uses the downloaded Bloomberg data files for exact accuracy
"""

import pandas as pd
import numpy as np
import warnings
from datetime import datetime
import os

warnings.filterwarnings('ignore')
pd.options.mode.chained_assignment = None

def load_real_bloomberg_data():
    """Load the downloaded Bloomberg data"""
    print("🔄 LOADING REAL BLOOMBERG DATA")
    print("=" * 60)
    
    # Check if downloaded data exists
    normalized_file = 'bloomberg_normalized_data.csv'
    raw_file = 'bloomberg_raw_data.csv'
    
    if os.path.exists(normalized_file):
        print(f"📊 Loading normalized data: {normalized_file}")
        data = pd.read_csv(normalized_file, index_col=0, parse_dates=True)
        print(f"✅ Loaded normalized data: {data.shape}")
        return data, True  # True = already normalized
        
    elif os.path.exists(raw_file):
        print(f"📊 Loading raw data: {raw_file}")
        data = pd.read_csv(raw_file, index_col=0, parse_dates=True)
        print(f"✅ Loaded raw data: {data.shape}")
        return data, False  # False = needs normalization
        
    else:
        print(f"❌ No downloaded data found!")
        print(f"   Expected files: {normalized_file} or {raw_file}")
        print(f"   Run download_bloomberg_data.py first")
        return None, None

def apply_normalization_if_needed(data, is_normalized):
    """Apply normalization if data is raw"""
    if is_normalized:
        print("✅ Data already normalized")
        return data
    
    print("🔧 Applying normalization to raw data...")
    
    # Load normalization factors
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    # Apply normalization
    data_normalized = data.copy()
    normalized_count = 0
    
    for ticker in data_normalized.columns:
        if ticker in norm_factors:
            data_normalized[ticker] = data_normalized[ticker] * norm_factors[ticker]
            normalized_count += 1
    
    print(f"✅ Applied normalization to {normalized_count}/{len(data_normalized.columns)} tickers")
    return data_normalized

def create_multiindex_structure(data):
    """Create MultiIndex structure from real data"""
    print("📋 Creating MultiIndex structure...")
    
    # Load ticker metadata
    dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
    
    # Build MultiIndex arrays
    level_0, level_1, level_2, level_3, level_4 = [], [], [], [], []
    valid_tickers = []
    
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        if pd.notna(ticker) and ticker in data.columns:
            level_0.append(row.get('Description', ''))
            level_1.append('')
            level_2.append(row.get('Category', ''))
            level_3.append(row.get('Region from', ''))
            level_4.append(row.get('Region to', ''))
            valid_tickers.append(ticker)
    
    # Create MultiIndex DataFrame
    multi_index = pd.MultiIndex.from_arrays([level_0, level_1, level_2, level_3, level_4])
    full_data = pd.DataFrame(data[valid_tickers].values, 
                           index=data.index, 
                           columns=multi_index)
    
    # Data cleaning
    full_data.index = pd.to_datetime(full_data.index)
    full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]
    full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')
    
    print(f"✅ MultiIndex data created: {full_data.shape}")
    return full_data

def create_demand_sheet_real_data(full_data):
    """Create Demand sheet using real Bloomberg data"""
    print("\n📊 CREATING DEMAND SHEET WITH REAL DATA")
    print("=" * 60)
    
    index = full_data.index
    countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 
                'Germany', 'Switzerland', 'Luxembourg']
    
    # Create DataFrame with target structure
    columns = countries + ['Island of Ireland', 'Total', 'Industrial', 'LDZ', 'Gas-to-Power']
    demand_df = pd.DataFrame(index=index, columns=columns)
    
    print("🌍 Processing countries with real data...")
    
    # Process each country
    for country in countries:
        mask = (full_data.columns.get_level_values(2) == 'Demand') & \
               (full_data.columns.get_level_values(3) == country)
        
        if mask.any():
            country_total = full_data.iloc[:, mask].sum(axis=1, skipna=True)
            demand_df[country] = country_total
            first_val = country_total.iloc[0]
            print(f"   {country:<12}: {mask.sum():2d} series → {first_val:8.2f}")
        else:
            demand_df[country] = 0.0
            print(f"   {country:<12}: No series found")
    
    # Island of Ireland
    ireland_mask = (full_data.columns.get_level_values(2) == 'Demand (Net)') & \
                   (full_data.columns.get_level_values(3) == 'Island of Ireland')
    
    if ireland_mask.any():
        demand_df['Island of Ireland'] = full_data.iloc[:, ireland_mask].sum(axis=1, skipna=True)
        print("   Ireland      : Net demand processed")
    else:
        demand_df['Island of Ireland'] = 0.0
    
    # Calculate total
    country_cols = countries + ['Island of Ireland']
    demand_df['Total'] = demand_df[country_cols].sum(axis=1, skipna=True)
    
    # Process categories with original complex logic
    print("\n📊 Processing categories with real data...")
    
    # Industrial
    industrial_mappings = [
        ('Demand', 'France', 'Industrial'),
        ('Demand', 'Belgium', 'Industrial'), 
        ('Demand', 'Italy', 'Industrial'),
        ('Demand', 'GB', 'Industrial'),
        ('Demand', 'Netherlands', 'Industrial and Power'),
        ('Demand', 'Netherlands', 'Zebra'),
        ('Demand', 'Germany', 'Industrial and Power'),
        ('Intermediate Calculation', '#Germany', 'Gas-to-Power'),
        ('Demand', 'Germany', 'Industrial (calculated)'),
        ('Demand', '#Netherlands', 'Industrial'),
        ('Intermediate Calculation', '#Netherlands', 'Gas-to-Power'),
        ('Demand', 'Netherlands', 'Industrial (calculated to 30/6/22 then actual)')
    ]
    
    industrial_total = pd.Series(0.0, index=index)
    for cat, region, subcat in industrial_mappings:
        mask = (full_data.columns.get_level_values(2) == cat) & \
               (full_data.columns.get_level_values(3) == region) & \
               (full_data.columns.get_level_values(4) == subcat)
        if mask.any():
            industrial_total += full_data.iloc[:, mask].sum(axis=1, skipna=True)
    
    demand_df['Industrial'] = industrial_total
    
    # Gas-to-Power
    gtp_mappings = [
        ('Demand', 'France', 'Gas-to-Power'),
        ('Demand', 'Belgium', 'Gas-to-Power'),
        ('Demand', 'Italy', 'Gas-to-Power'),
        ('Demand', 'GB', 'Gas-to-Power'),
        ('Demand', 'Germany', 'Industrial and Power'),
        ('Intermediate Calculation', '#Germany', 'Gas-to-Power'),
        ('Demand', 'Germany', 'Gas-to-Power (calculated)'),
        ('Demand', 'Netherlands', 'Industrial and Power'),
        ('Demand', '#Netherlands', 'Gas-to-Power'),
        ('Intermediate Calculation', '#Netherlands', 'Gas-to-Power'),
        ('Demand', 'Netherlands', 'Gas-to-Power (calculated to 30/6/22 then actual)')
    ]
    
    gtp_total = pd.Series(0.0, index=index)
    for cat, region, subcat in gtp_mappings:
        mask = (full_data.columns.get_level_values(2) == cat) & \
               (full_data.columns.get_level_values(3) == region) & \
               (full_data.columns.get_level_values(4) == subcat)
        if mask.any():
            gtp_total += full_data.iloc[:, mask].sum(axis=1, skipna=True)
    
    demand_df['Gas-to-Power'] = gtp_total
    
    # LDZ
    ldz_mappings = [
        ('Demand', 'France', 'LDZ'),
        ('Demand', 'Belgium', 'LDZ'),
        ('Demand', 'Italy', 'LDZ'),
        ('Demand', 'Italy', 'Other'),
        ('Demand', 'Netherlands', 'LDZ'),
        ('Demand', 'GB', 'LDZ'),
        ('Demand', 'Austria', 'Austria'),
        ('Demand', 'Germany', 'LDZ'),
        ('Demand', 'Switzerland', 'Switzerland'),
        ('Demand', 'Luxembourg', 'Luxembourg'),
        ('Demand (Net)', 'Island of Ireland', 'Island of Ireland'),
        ('Demand', '#Germany', 'Implied LDZ')
    ]
    
    ldz_total = pd.Series(0.0, index=index)
    for cat, region, subcat in ldz_mappings:
        mask = (full_data.columns.get_level_values(2) == cat) & \
               (full_data.columns.get_level_values(3) == region) & \
               (full_data.columns.get_level_values(4) == subcat)
        if mask.any():
            ldz_total += full_data.iloc[:, mask].sum(axis=1, skipna=True)
    
    demand_df['LDZ'] = ldz_total
    
    return demand_df

def verify_against_target_real_data(demand_df):
    """Verify real data results against target"""
    print("\n🎯 VERIFICATION WITH REAL DATA")
    print("=" * 60)
    
    # Target values (from Excel Demand sheet)
    targets = {
        'Italy': 115.24,
        'Total': 614.04,
        'Industrial': 207.43,
        'LDZ': 243.89,
        'Gas-to-Power': 159.34
    }
    
    first_row = demand_df.iloc[0]
    
    print("📊 REAL DATA vs TARGET (first row):")
    perfect_matches = 0
    close_matches = 0
    
    for key, target_val in targets.items():
        if key in first_row.index:
            our_val = first_row[key]
            diff = abs(our_val - target_val)
            
            if diff < 1.0:
                status = "🎯 PERFECT"
                perfect_matches += 1
            elif diff < 5.0:
                status = "✅ CLOSE"
                close_matches += 1
            else:
                status = "❌ OFF"
            
            print(f"   {key:<12}: Target={target_val:8.2f}, Real={our_val:8.2f}, Diff={diff:6.2f} {status}")
    
    print(f"\n📊 ACCURACY SUMMARY:")
    print(f"   🎯 Perfect matches (<1.0 diff): {perfect_matches}/5")
    print(f"   ✅ Close matches (<5.0 diff): {close_matches}/5")
    print(f"   📈 Total acceptable: {perfect_matches + close_matches}/5")
    
    # Check category sum
    categories_sum = first_row['Industrial'] + first_row['LDZ'] + first_row['Gas-to-Power']
    total_val = first_row['Total']
    sum_diff = abs(total_val - categories_sum)
    
    print(f"\n📊 CATEGORY SUM CHECK:")
    print(f"   Industrial + LDZ + Gas-to-Power = {categories_sum:.2f}")
    print(f"   Total = {total_val:.2f}")
    print(f"   Difference = {sum_diff:.2f}")
    print(f"   Status: {'✅ GOOD' if sum_diff < 10.0 else '❌ CHECK'}")
    
    return (perfect_matches + close_matches) >= 4  # At least 4/5 should be close

def main():
    """Main execution with real data"""
    print("🎯 GAS MARKET PROCESSOR - REAL DATA VERSION")
    print("=" * 80)
    print("Uses downloaded Bloomberg data for exact accuracy")
    print("=" * 80)
    
    try:
        # Step 1: Load real Bloomberg data
        data, is_normalized = load_real_bloomberg_data()
        if data is None:
            return False
        
        # Step 2: Apply normalization if needed
        normalized_data = apply_normalization_if_needed(data, is_normalized)
        
        # Step 3: Create MultiIndex structure
        full_data = create_multiindex_structure(normalized_data)
        
        # Step 4: Create Demand sheet
        demand_df = create_demand_sheet_real_data(full_data)
        
        # Step 5: Verify against targets
        success = verify_against_target_real_data(demand_df)
        
        # Step 6: Save results
        output_file = 'DNB Markets EUROPEAN GAS BALANCE_real_data.xlsx'
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                demand_df.to_excel(writer, sheet_name='Demand')
            print(f"\n✅ Saved: {output_file}")
        except Exception as e:
            print(f"\n⚠️  Save failed: {e}")
        
        # Summary
        print(f"\n🎉 REAL DATA RESULTS")
        print("=" * 50)
        print(f"✅ Used real Bloomberg data")
        print(f"{'✅' if success else '⚠️ '} Accuracy: {success}")
        print(f"🇮🇹 Italy: {demand_df['Italy'].iloc[0]:.2f} (target: 115.24)")
        print(f"📊 Total: {demand_df['Total'].iloc[0]:.2f} (target: 614.04)")
        
        return success
        
    except Exception as e:
        print(f"\n💥 ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)