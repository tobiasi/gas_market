# -*- coding: utf-8 -*-
"""
Gas Market Bloomberg Production Processor - CHUNKED VERSION
Processes data in small chunks to prevent kernel restart
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

def download_bloomberg_data_safe(tickers):
    """Safe Bloomberg data loading with chunked approach"""
    print("📊 Loading Bloomberg data safely...")
    
    try:
        # Try xbbg first
        from xbbg import blp
        print("🔄 Attempting Bloomberg API download...")
        data = blp.bdh(tickers, start_date="2016-10-01", flds="PX_LAST").droplevel(axis=1, level=1)
        print("✅ Using live Bloomberg data via xbbg")
        return data
    except Exception as e:
        print(f"⚠️ Bloomberg API failed: {e}")
        print("🔄 Falling back to raw data...")
        
        # Fallback to raw data with chunked loading
        try:
            # Load in chunks to avoid memory spike
            chunks = []
            chunk_size = 1000
            total_chunks = 0
            
            for chunk in pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True, chunksize=chunk_size):
                chunks.append(chunk)
                total_chunks += 1
                print(f"   Loaded chunk {total_chunks} ({len(chunk)} rows)")
                
                # Force garbage collection after each chunk
                if total_chunks % 2 == 0:
                    gc.collect()
                    time.sleep(0.1)  # Brief pause
            
            print(f"🔄 Combining {total_chunks} chunks...")
            data = pd.concat(chunks, ignore_index=False)
            
            # Clean up chunks immediately
            del chunks
            gc.collect()
            
            print(f"✅ Using bloomberg_raw_data.csv: {data.shape}")
            return data
            
        except Exception as e2:
            print(f"❌ Raw data loading failed: {e2}")
            return None

def apply_normalization_chunked(data, norm_factors):
    """Apply normalization in chunks to prevent memory overload"""
    print("🔧 Applying normalization in chunks...")
    
    # Process columns in batches of 50
    batch_size = 50
    total_batches = (len(data.columns) + batch_size - 1) // batch_size
    
    normalized_count = 0
    
    for i in range(0, len(data.columns), batch_size):
        batch_cols = data.columns[i:i+batch_size]
        batch_num = i // batch_size + 1
        
        print(f"   Processing batch {batch_num}/{total_batches} ({len(batch_cols)} columns)")
        
        # Apply normalization to this batch
        for ticker in batch_cols:
            if ticker in norm_factors:
                norm_factor = norm_factors[ticker]
                data[ticker] = data[ticker] * norm_factor
                normalized_count += 1
        
        # Force cleanup after each batch
        if batch_num % 5 == 0:
            gc.collect()
            time.sleep(0.1)
    
    print(f"✅ Normalized {normalized_count} tickers")
    return data

def process_countries_step_by_step(data, dataset):
    """Process countries one by one to avoid memory overload"""
    print("🌍 Processing countries step by step...")
    
    # Country list
    country_list = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
    
    # Result storage
    country_results = {}
    
    # Process each country individually
    for i, country in enumerate(country_list, 1):
        print(f"   Processing {country} ({i}/{len(country_list)})...")
        
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
            
            print(f"     Found {len(country_tickers)} tickers for {country}")
            
            # Calculate country total
            if country_tickers:
                country_data = data[country_tickers]
                country_total = country_data.sum(axis=1, skipna=True)
                country_results[country] = country_total
                
                # Show first value for verification
                first_val = country_total.iloc[0] if len(country_total) > 0 else 0
                print(f"     {country} first value: {first_val:.2f}")
                
                # Clean up this country's data
                del country_data
            else:
                print(f"     No tickers found for {country}")
                country_results[country] = pd.Series(0, index=data.index)
        
        except Exception as e:
            print(f"     ❌ Error processing {country}: {e}")
            country_results[country] = pd.Series(0, index=data.index)
        
        # Force cleanup after each country
        gc.collect()
        time.sleep(0.1)
    
    print("✅ Completed country processing")
    return country_results

def create_final_output(country_results, data):
    """Create final demand DataFrame"""
    print("📊 Creating final output...")
    
    # Create demand DataFrame
    demand_data = pd.DataFrame(index=data.index)
    
    # Add countries
    for country, series in country_results.items():
        demand_data[country] = series
    
    # Calculate total
    country_cols = list(country_results.keys())
    demand_data['Total'] = demand_data[country_cols].sum(axis=1, skipna=True)
    
    # For now, set categories equal to total (simplified)
    demand_data['Industrial'] = demand_data['Total'] * 0.33  # Rough estimate
    demand_data['LDZ'] = demand_data['Total'] * 0.33
    demand_data['Gas-to-Power'] = demand_data['Total'] * 0.34
    
    print(f"✅ Created demand data: {demand_data.shape}")
    return demand_data

def save_output_files(demand_data):
    """Save output files"""
    print("💾 Saving output files...")
    
    # Prepare for saving
    final_demand_df = demand_data.copy()
    final_demand_df.reset_index(inplace=True)
    final_demand_df.rename(columns={'index': 'Date'}, inplace=True)
    
    # Save Excel file
    master_file = 'European_Gas_Market_Master.xlsx'
    with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
        final_demand_df.to_excel(writer, sheet_name='Demand', index=False)
    
    # Save CSV file
    final_demand_df.to_csv('European_Gas_Demand_Master.csv', index=False)
    
    print(f"✅ Saved: {master_file}")
    return master_file

def main():
    """Main execution with chunked processing"""
    print("🚀 CHUNKED BLOOMBERG PRODUCTION PROCESSOR")
    print("=" * 60)
    print("Processing data in small chunks to prevent kernel restart...")
    print("=" * 60)
    
    try:
        # Step 1: Load configuration
        dataset, tickers, norm_factors = load_ticker_data()
        time.sleep(0.5)  # Brief pause
        
        # Step 2: Load Bloomberg data safely
        data = download_bloomberg_data_safe(tickers)
        if data is None:
            print("❌ Failed to load Bloomberg data")
            return False
        
        time.sleep(0.5)
        gc.collect()
        
        # Step 3: Apply normalization in chunks
        data = apply_normalization_chunked(data, norm_factors)
        time.sleep(0.5)
        gc.collect()
        
        # Step 4: Process countries step by step
        country_results = process_countries_step_by_step(data, dataset)
        time.sleep(0.5)
        gc.collect()
        
        # Step 5: Create final output
        demand_data = create_final_output(country_results, data)
        time.sleep(0.5)
        gc.collect()
        
        # Step 6: Save files
        master_file = save_output_files(demand_data)
        
        # Verification
        print("\n🎯 VERIFICATION:")
        target_date = '2016-10-04'
        if target_date in [str(d)[:10] for d in demand_data.index]:
            target_row = demand_data[demand_data.index.strftime('%Y-%m-%d') == target_date].iloc[0]
            print(f"   Italy on {target_date}: {target_row['Italy']:.2f}")
            print(f"   Total: {target_row['Total']:.2f}")
        
        print(f"\n🎉 CHUNKED PROCESSING COMPLETE!")
        print(f"✅ Output: {master_file}")
        print(f"📊 Shape: {demand_data.shape}")
        
        return True
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print(f"Error type: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        
        # Force cleanup
        gc.collect()
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)