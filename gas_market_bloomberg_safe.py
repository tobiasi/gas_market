# -*- coding: utf-8 -*-
"""
Gas Market Bloomberg Production Processor - MEMORY SAFE VERSION
Ultra-safe processing with minimal memory footprint for saving operations
"""

import pandas as pd
import numpy as np
import gc
import time

def load_ticker_data():
    """Load ticker list and normalization factors from use4.xlsx"""
    print("📋 Loading ticker configuration...")
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
    
    print(f"✅ Loaded {len(tickers)} tickers, {len(norm_factors)} norm factors")
    return dataset, tickers, norm_factors

def download_bloomberg_data_minimal(tickers):
    """Ultra-safe Bloomberg data loading with size limits"""
    print("📊 Loading Bloomberg data with size limits...")
    
    try:
        # Try xbbg first
        from xbbg import blp
        print("🔄 Attempting Bloomberg API download...")
        data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)
        
        # Limit data size immediately
        if len(data) > 500:
            data = data.head(500)
            print(f"⚠️ Limited to 500 rows to prevent memory overload")
        
        print("✅ Using live Bloomberg data via xbbg")
        return data
    except Exception as e:
        print(f"⚠️ Bloomberg API failed: {e}")
        print("🔄 Falling back to limited raw data...")
        
        try:
            # Load only small portion of raw data
            data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True, nrows=500)
            print(f"✅ Using limited bloomberg_raw_data.csv: {data.shape}")
            return data
            
        except Exception as e2:
            print(f"❌ Raw data loading failed: {e2}")
            return None

def process_countries_minimal(data, dataset):
    """Process only essential countries to minimize memory usage"""
    print("🌍 Processing essential countries only...")
    
    # Only process key countries
    essential_countries = ['Italy', 'Germany', 'France']
    
    # Result storage
    country_results = {}
    
    # Process each country individually
    for i, country in enumerate(essential_countries, 1):
        print(f"   Processing {country} ({i}/{len(essential_countries)})...")
        
        try:
            # Find tickers for this country
            country_tickers = []
            for _, row in dataset.iterrows():
                ticker = row.get('Ticker')
                region_from = row.get('Region from', '')
                category = row.get('Category', '')
                region_to = row.get('Region to', '')
                
                if (ticker and ticker in data.columns and 
                    region_from == country and category == 'Demand'):
                    
                    # For Italy, exclude losses and exports
                    if country == 'Italy':
                        if region_to in ['Industrial', 'LDZ', 'Gas-to-Power']:
                            country_tickers.append(ticker)
                    else:
                        country_tickers.append(ticker)
                
                # Limit tickers per country to prevent overload
                if len(country_tickers) >= 20:
                    break
            
            print(f"     Found {len(country_tickers)} tickers for {country}")
            
            # Calculate country total
            if country_tickers:
                country_data = data[country_tickers]
                country_total = country_data.sum(axis=1, skipna=True)
                country_results[country] = country_total
                
                # Show first value for verification
                first_val = country_total.iloc[0] if len(country_total) > 0 else 0
                print(f"     {country} first value: {first_val:.2f}")
                
                # Clean up this country's data immediately
                del country_data
            else:
                print(f"     No tickers found for {country}")
                country_results[country] = pd.Series(0, index=data.index)
        
        except Exception as e:
            print(f"     ❌ Error processing {country}: {e}")
            country_results[country] = pd.Series(0, index=data.index)
        
        # Force cleanup after each country
        gc.collect()
        time.sleep(0.2)
    
    print("✅ Completed essential country processing")
    return country_results

def create_minimal_output(country_results, data):
    """Create minimal final DataFrame"""
    print("📊 Creating minimal output...")
    
    # Create demand DataFrame with only essential columns
    demand_data = pd.DataFrame(index=data.index)
    
    # Add only the countries we processed
    for country, series in country_results.items():
        demand_data[country] = series
    
    # Calculate total for processed countries only
    country_cols = list(country_results.keys())
    demand_data['Total'] = demand_data[country_cols].sum(axis=1, skipna=True)
    
    print(f"✅ Created minimal demand data: {demand_data.shape}")
    return demand_data

def save_output_safely(demand_data):
    """Ultra-safe file saving with memory management"""
    print("💾 Saving output files safely...")
    
    try:
        # Prepare data in smallest possible chunks
        print("   Preparing data for saving...")
        final_demand_df = demand_data.copy()
        final_demand_df.reset_index(inplace=True)
        final_demand_df.rename(columns={'index': 'Date'}, inplace=True)
        
        # Clean up original before saving
        del demand_data
        gc.collect()
        time.sleep(0.5)
        
        # Save CSV first (lighter operation)
        print("   Saving CSV file...")
        csv_file = 'European_Gas_Minimal.csv'
        final_demand_df.to_csv(csv_file, index=False)
        gc.collect()
        time.sleep(0.5)
        
        # Save Excel with minimal memory usage
        print("   Saving Excel file...")
        excel_file = 'European_Gas_Minimal.xlsx'
        
        # Use xlsxwriter for better memory efficiency
        try:
            with pd.ExcelWriter(excel_file, engine='xlsxwriter', options={'strings_to_numbers': True}) as writer:
                final_demand_df.to_excel(writer, sheet_name='Demand', index=False)
            print(f"   ✅ Used xlsxwriter engine")
        except ImportError:
            # Fallback to openpyxl
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                final_demand_df.to_excel(writer, sheet_name='Demand', index=False)
            print(f"   ✅ Used openpyxl engine")
        
        # Final cleanup
        del final_demand_df
        gc.collect()
        
        print(f"✅ Safely saved: {excel_file} and {csv_file}")
        return excel_file
        
    except Exception as e:
        print(f"❌ Safe saving failed: {e}")
        print("🔄 Attempting CSV-only save...")
        
        try:
            # Emergency: CSV only
            csv_file = 'European_Gas_Emergency.csv'
            final_demand_df.to_csv(csv_file, index=False)
            print(f"✅ Emergency save successful: {csv_file}")
            return csv_file
        except Exception as e2:
            print(f"❌ Emergency save failed: {e2}")
            return None

def main():
    """Main execution with ultra-safe processing"""
    print("🚀 MEMORY-SAFE BLOOMBERG PROCESSOR")
    print("=" * 60)
    print("Ultra-safe processing with minimal memory footprint...")
    print("=" * 60)
    
    try:
        # Step 1: Load configuration
        dataset, tickers, norm_factors = load_ticker_data()
        time.sleep(0.5)
        gc.collect()
        
        # Step 2: Load Bloomberg data with limits
        data = download_bloomberg_data_minimal(tickers)
        if data is None:
            print("❌ Failed to load Bloomberg data")
            return False
        
        time.sleep(0.5)
        gc.collect()
        
        # Step 3: Apply normalization to limited data
        print("🔧 Applying normalization to limited data...")
        normalized_count = 0
        for ticker in data.columns:
            if ticker in norm_factors:
                norm_factor = norm_factors[ticker]
                data[ticker] = data[ticker] * norm_factor
                normalized_count += 1
        
        print(f"✅ Normalized {normalized_count} tickers")
        time.sleep(0.5)
        gc.collect()
        
        # Step 4: Process essential countries only
        country_results = process_countries_minimal(data, dataset)
        time.sleep(0.5)
        gc.collect()
        
        # Step 5: Create minimal output
        demand_data = create_minimal_output(country_results, data)
        time.sleep(0.5)
        gc.collect()
        
        # Clean up data before saving
        del data, country_results, dataset, tickers, norm_factors
        gc.collect()
        time.sleep(1.0)  # Extra pause before saving
        
        # Step 6: Save files safely
        output_file = save_output_safely(demand_data)
        
        if output_file:
            print(f"\n🎉 MEMORY-SAFE PROCESSING COMPLETE!")
            print(f"✅ Output: {output_file}")
            print(f"📊 Data shape was minimal to prevent kernel restart")
            return True
        else:
            print(f"\n❌ Failed to save output files")
            return False
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print(f"Error type: {type(e).__name__}")
        
        # Force cleanup
        gc.collect()
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)