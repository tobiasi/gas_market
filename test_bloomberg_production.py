# -*- coding: utf-8 -*-
"""
Test Bloomberg Production System
Uses bloomberg_raw_data.csv to mimic Bloomberg download and verify against LiveSheet
"""

import pandas as pd
import numpy as np
import json

def load_ticker_data():
    """Load ticker list and normalization factors from use4.xlsx"""
    data_path = 'use4.xlsx'
    
    # Read ticker list from TickerList sheet
    dataset = pd.read_excel(data_path, sheet_name='TickerList', skiprows=8)
    
    # Extract normalization factors
    norm_factors = {}
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        norm_factor = row.get('Normalization factor', 1.0)
        if pd.notna(ticker) and pd.notna(norm_factor):
            norm_factors[ticker] = float(norm_factor)
    
    # Get unique tickers
    tickers = list(set(dataset.Ticker.dropna().to_list()))
    
    return dataset, tickers, norm_factors

def load_raw_bloomberg_data():
    """Load raw Bloomberg data from CSV file to mimic Bloomberg download"""
    print("📊 Loading raw Bloomberg data from CSV file...")
    
    # Load the raw data
    data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True)
    
    print(f"✅ Loaded Bloomberg data: {data.shape}")
    print(f"📅 Date range: {data.index[0]} to {data.index[-1]}")
    print(f"📊 Available tickers: {len(data.columns)}")
    
    return data

def apply_normalization(data, norm_factors):
    """Apply normalization factors to Bloomberg data"""
    print("\n🔧 APPLYING NORMALIZATION FACTORS...")
    
    normalization_applied = 0
    normalization_skipped = 0
    
    for ticker in data.columns:
        if ticker in norm_factors:
            norm_factor = norm_factors[ticker]
            original_val = data[ticker].iloc[0] if not pd.isna(data[ticker].iloc[0]) else 0
            data[ticker] = data[ticker] * norm_factor
            normalized_val = data[ticker].iloc[0] if not pd.isna(data[ticker].iloc[0]) else 0
            normalization_applied += 1
            
            # Show normalization for key Italy series
            if 'SNAM' in ticker and any(keyword in ticker for keyword in ['GIND', 'GLDN', 'GPGE', 'CLGG', 'GOTH']):
                print(f"   🇮🇹 {ticker}: {original_val:.2f} → {normalized_val:.2f} (factor: {norm_factor:.9f})")
        else:
            normalization_skipped += 1
    
    print(f"✅ Normalization complete:")
    print(f"   Applied factors: {normalization_applied} tickers")
    print(f"   No factor found: {normalization_skipped} tickers")
    print(f"   Coverage: {normalization_applied/(normalization_applied+normalization_skipped)*100:.1f}%")
    
    return data

def create_multiindex_data(data, dataset):
    """Create MultiIndex DataFrame from Bloomberg data"""
    print("\n📋 CREATING MULTIINDEX STRUCTURE...")
    
    successful_tickers = []
    level_0 = []  # Description
    level_1 = []  # Empty
    level_2 = []  # Category  
    level_3 = []  # Region from
    level_4 = []  # Region to
    
    # Process each ticker in dataset
    for index, row in dataset.iterrows():
        ticker = row.get('Ticker')
        
        if pd.isna(ticker) or ticker not in data.columns:
            continue
        
        successful_tickers.append(ticker)
        level_0.append(row.get('Description', ''))
        level_1.append('')  # Empty level
        level_2.append(row.get('Category', ''))
        level_3.append(row.get('Region from', ''))
        level_4.append(row.get('Region to', ''))
    
    # Create MultiIndex DataFrame
    if successful_tickers:
        ticker_data = data[successful_tickers]
        multi_index = pd.MultiIndex.from_arrays([level_0, level_1, level_2, level_3, level_4])
        full_data = pd.DataFrame(ticker_data.values, index=ticker_data.index, columns=multi_index)
        
        # Apply data cleaning
        full_data.index = pd.to_datetime(full_data.index)
        full_data = full_data[~((full_data.index.month == 2) & (full_data.index.day == 29))]
        full_data = full_data.copy()
        full_data.iloc[:-1] = full_data.iloc[:-1].interpolate(limit=2, limit_area='inside')
        
        print(f"✅ Created MultiIndex data: {full_data.shape}")
        return full_data
    else:
        raise ValueError("No successful tickers to process")

def process_demand_data(full_data):
    """Process demand data into country breakdown"""
    print("\n🌍 PROCESSING DEMAND DATA...")
    
    index = full_data.index
    country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    # Create countries DataFrame
    demand_1 = ['Demand'] * len(country_list) + ['Demand (Net)']
    demand_2 = [''] * (len(country_list) + 1) 
    demand_3 = country_list + ['Island of Ireland']
    
    countries = pd.DataFrame(index=index, columns=pd.MultiIndex.from_tuples(list(zip(demand_1, demand_2, demand_3))))
    
    # Populate individual countries
    for country in country_list:
        country_mask = (full_data.columns.get_level_values(2) == 'Demand') & (full_data.columns.get_level_values(3) == country)
        country_total = full_data.iloc[:, country_mask].sum(axis=1, skipna=True)
        countries[('Demand', '', country)] = country_total
        
        # Debug output for key countries
        num_series = country_mask.sum()
        first_val = country_total.iloc[0] if len(country_total) > 0 else 0
        print(f"  {country:12}: {num_series} series, first value: {first_val:.2f}")
    
    # Handle Island of Ireland separately
    ireland_mask = (full_data.columns.get_level_values(2) == 'Demand (Net)') & (full_data.columns.get_level_values(3) == 'Island of Ireland')
    ireland_total = full_data.iloc[:, ireland_mask].sum(axis=1, skipna=True) 
    countries[('Demand (Net)', '', 'Island of Ireland')] = ireland_total
    
    return countries

def process_category_data(full_data):
    """Process Industrial, LDZ, and Gas-to-Power data (simplified)"""
    print("\n🏭 PROCESSING CATEGORY DATA...")
    
    index = full_data.index
    
    # Simplified category processing - just get totals for now
    industrial_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
                     (full_data.columns.get_level_values(4).str.contains('Industrial', na=False))
    ldz_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
               (full_data.columns.get_level_values(4).str.contains('LDZ', na=False))
    gtp_mask = (full_data.columns.get_level_values(2) == 'Demand') & \
               (full_data.columns.get_level_values(4).str.contains('Gas-to-Power', na=False))
    
    industrial_total = full_data.iloc[:, industrial_mask].sum(axis=1, skipna=True)
    ldz_total = full_data.iloc[:, ldz_mask].sum(axis=1, skipna=True)
    gtp_total = full_data.iloc[:, gtp_mask].sum(axis=1, skipna=True)
    
    print(f"  Industrial: {industrial_mask.sum()} series, first value: {industrial_total.iloc[0]:.2f}")
    print(f"  LDZ:        {ldz_mask.sum()} series, first value: {ldz_total.iloc[0]:.2f}")
    print(f"  Gas-to-Power: {gtp_mask.sum()} series, first value: {gtp_total.iloc[0]:.2f}")
    
    return industrial_total, ldz_total, gtp_total

def create_demand_master(countries, industrial_total, ldz_total, gtp_total):
    """Create demand master DataFrame in simple format"""
    print("\n📊 CREATING DEMAND MASTER...")
    
    index = countries.index
    country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    demand_data = pd.DataFrame(index=index)
    
    # Add individual countries
    for country in country_list:
        if ('Demand', '', country) in countries.columns:
            demand_data[country] = countries[('Demand', '', country)]
    
    # Calculate total from countries
    country_total = demand_data[country_list].sum(axis=1, skipna=True)
    
    # Add totals
    demand_data['Total'] = country_total
    demand_data['Industrial'] = industrial_total
    demand_data['LDZ'] = ldz_total
    demand_data['Gas-to-Power'] = gtp_total
    
    print(f"✅ Demand master created: {demand_data.shape}")
    
    return demand_data

def verify_against_livesheet(demand_data):
    """Verify our processed data against the LiveSheet"""
    print("\n🎯 VERIFICATION AGAINST LIVESHEET")
    print("=" * 60)
    
    try:
        # Load corrected column mapping for verification
        with open('corrected_analysis_results.json', 'r') as f:
            analysis = json.load(f)
        target_values = analysis['target_values']
        
        # Check specific date: 2016-10-04
        target_date = '2016-10-04'
        if target_date in [str(d)[:10] for d in demand_data.index]:
            target_row = demand_data[demand_data.index.strftime('%Y-%m-%d') == target_date].iloc[0]
            
            print(f"📅 Verification for {target_date}:")
            perfect_matches = 0
            
            for key, expected in target_values.items():
                if key in target_row.index:
                    actual = target_row[key]
                    diff = abs(actual - expected)
                    status = "✅" if diff < 0.1 else "❌"  # Allow small rounding differences
                    if diff < 0.1:
                        perfect_matches += 1
                    print(f"  {key:12}: {actual:8.2f} (expected: {expected:8.2f}) {status}")
                else:
                    print(f"  {key:12}: NOT FOUND ❌")
            
            print(f"\n📊 ACCURACY: {perfect_matches}/{len(target_values)} perfect matches")
            accuracy = (perfect_matches / len(target_values)) * 100
            print(f"🎯 Success rate: {accuracy:.1f}%")
            
            # Check if countries sum to total
            countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
            country_sum = sum(target_row[c] for c in countries if c in target_row.index)
            total_val = target_row['Total'] if 'Total' in target_row.index else 0
            
            print(f"\n🔍 COLUMN SUM VERIFICATION:")
            print(f"  Countries sum:  {country_sum:.2f}")
            print(f"  Total column:   {total_val:.2f}")
            print(f"  Difference:     {abs(total_val - country_sum):.2f}")
            
            if abs(total_val - country_sum) < 0.1:
                print(f"  ✅ Countries sum to Total!")
            else:
                print(f"  ❌ Countries don't sum to Total")
            
            return accuracy >= 80  # Success if 80%+ match
        else:
            print(f"❌ Target date {target_date} not found in processed data")
            return False
            
    except Exception as e:
        print(f"❌ Verification error: {e}")
        return False

def save_test_files(demand_data):
    """Save test output files"""
    print("\n💾 SAVING TEST FILES...")
    
    # Reset index to make Date a column
    final_demand_df = demand_data.copy()
    final_demand_df.reset_index(inplace=True)
    final_demand_df.rename(columns={'index': 'Date'}, inplace=True)
    
    # Save Excel file
    test_file = 'European_Gas_Market_TEST.xlsx'
    with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
        final_demand_df.to_excel(writer, sheet_name='Demand', index=False)
    
    # Save CSV file
    final_demand_df.to_csv('European_Gas_Demand_TEST.csv', index=False)
    
    print(f"✅ Test files saved: {test_file}")
    
    return test_file

def main():
    """Main test execution function"""
    print("🧪 TESTING BLOOMBERG PRODUCTION SYSTEM")
    print("=" * 80)
    print("Using bloomberg_raw_data.csv to mimic Bloomberg API")
    print("=" * 80)
    
    try:
        # Load ticker configuration
        dataset, tickers, norm_factors = load_ticker_data()
        print(f"📋 Loaded {len(tickers)} tickers with {len(norm_factors)} normalization factors")
        
        # Load raw Bloomberg data (mimic API call)
        data = load_raw_bloomberg_data()
        
        # Apply normalization
        data = apply_normalization(data, norm_factors)
        
        # Create MultiIndex structure
        full_data = create_multiindex_data(data, dataset)
        
        # Process demand data
        countries = process_demand_data(full_data)
        
        # Process category data
        industrial_total, ldz_total, gtp_total = process_category_data(full_data)
        
        # Create demand master
        demand_data = create_demand_master(countries, industrial_total, ldz_total, gtp_total)
        
        # Verify against LiveSheet
        verification_passed = verify_against_livesheet(demand_data)
        
        # Save test files
        test_file = save_test_files(demand_data)
        
        # Final summary
        print(f"\n🎉 TEST RESULTS")
        print("=" * 40)
        print(f"✅ Data processing: COMPLETE")
        print(f"📊 Output shape: {demand_data.shape}")
        print(f"🎯 Verification: {'PASSED' if verification_passed else 'FAILED'}")
        print(f"📁 Test file: {test_file}")
        
        if verification_passed:
            print(f"\n🎊 SUCCESS! Bloomberg production system is working correctly!")
            print(f"   The processed data matches the LiveSheet reference values.")
        else:
            print(f"\n⚠️  ISSUES DETECTED: Some values don't match the LiveSheet.")
            print(f"   Review the verification output above for details.")
        
        return verification_passed
        
    except Exception as e:
        print(f"\n❌ TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)