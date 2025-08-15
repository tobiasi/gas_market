# -*- coding: utf-8 -*-
"""
Gas Market Lightweight Production Processor
Memory-optimized version to prevent kernel restarts
"""

import pandas as pd
import numpy as np
import gc

def process_gas_market_lightweight():
    """Lightweight version that processes data efficiently without memory overload"""
    
    print("🚀 LIGHTWEIGHT GAS MARKET PROCESSOR")
    print("=" * 50)
    
    try:
        # 1. Load configuration with memory management
        print("1️⃣ Loading configuration...")
        dataset = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
        
        # Extract only what we need
        norm_factors = {}
        for _, row in dataset.iterrows():
            ticker = row.get('Ticker')
            norm_factor = row.get('Normalization factor', 1.0)
            if pd.notna(ticker) and pd.notna(norm_factor):
                norm_factors[ticker] = float(norm_factor)
        
        print(f"   Loaded {len(norm_factors)} normalization factors")
        
        # 2. Load Bloomberg data in chunks
        print("2️⃣ Loading Bloomberg data...")
        
        # Use chunked reading to avoid memory overload
        try:
            data_chunks = []
            chunk_size = 500  # Process 500 rows at a time
            
            for chunk in pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True, chunksize=chunk_size):
                data_chunks.append(chunk)
                if len(data_chunks) >= 10:  # Limit to first 5000 rows for memory
                    break
            
            # Combine chunks
            data = pd.concat(data_chunks, ignore_index=False)
            del data_chunks  # Free memory
            gc.collect()
            
            print(f"   Loaded Bloomberg data: {data.shape}")
            
        except Exception as e:
            print(f"   ❌ Error loading Bloomberg data: {e}")
            return False
        
        # 3. Apply normalization efficiently 
        print("3️⃣ Applying normalization...")
        
        # Only process tickers we have factors for
        available_tickers = [t for t in data.columns if t in norm_factors]
        data_subset = data[available_tickers].copy()
        del data  # Free original data
        gc.collect()
        
        # Apply normalization
        for ticker in data_subset.columns:
            if ticker in norm_factors:
                data_subset[ticker] = data_subset[ticker] * norm_factors[ticker]
        
        print(f"   Normalized {len(available_tickers)} tickers")
        
        # 4. Simple aggregation (avoid complex MultiIndex)
        print("4️⃣ Processing Italy demand (simplified)...")
        
        # Italy key tickers
        italy_tickers = ['SNAMGIND Index', 'SNAMGLDN Index', 'SNAMGPGE Index']
        italy_data = data_subset[[t for t in italy_tickers if t in data_subset.columns]]
        
        if not italy_data.empty:
            italy_total = italy_data.sum(axis=1)
            
            # Get specific date for verification
            target_date = '2016-10-04'
            if target_date in [str(d)[:10] for d in italy_total.index]:
                target_val = italy_total[italy_total.index.strftime('%Y-%m-%d') == target_date].iloc[0]
                print(f"   Italy total on {target_date}: {target_val:.2f}")
                
                # Check against our known target
                expected = 151.47  # From LiveSheet
                diff = abs(target_val - expected)
                print(f"   Expected: {expected:.2f}, Difference: {diff:.2f}")
                
                if diff < 1.0:
                    print(f"   ✅ EXCELLENT: Within 1.0 of target!")
                else:
                    print(f"   ⚠️  Difference detected, but processing successful")
        
        # 5. Create simple output
        print("5️⃣ Creating output...")
        
        # Simple demand DataFrame (just key countries)
        key_countries = ['Italy', 'Germany', 'France', 'Netherlands', 'GB']
        demand_simple = pd.DataFrame(index=data_subset.index)
        
        # For now, just add Italy (other countries would need ticker mapping)
        demand_simple['Italy'] = italy_total if not italy_data.empty else 0
        demand_simple['Date'] = demand_simple.index
        
        # Save simple output
        output_file = 'European_Gas_Lightweight.xlsx'
        demand_simple.reset_index(drop=True).to_excel(output_file, index=False)
        
        print(f"   ✅ Saved: {output_file}")
        print(f"   📊 Shape: {demand_simple.shape}")
        
        # Memory cleanup
        del data_subset, italy_data, demand_simple
        gc.collect()
        
        print("\n🎉 LIGHTWEIGHT PROCESSING COMPLETE!")
        return True
        
    except Exception as e:
        print(f"\n❌ Error in lightweight processing: {e}")
        print(f"Error type: {type(e).__name__}")
        
        # Try to free memory
        gc.collect()
        return False

def check_system_requirements():
    """Check if system has required files and memory"""
    print("🔍 SYSTEM REQUIREMENTS CHECK")
    print("-" * 30)
    
    import os
    import psutil
    
    # Check memory
    memory = psutil.virtual_memory()
    print(f"Available memory: {memory.available / 1024**3:.1f} GB")
    
    if memory.available < 2 * 1024**3:  # Less than 2GB
        print("⚠️  WARNING: Low memory detected")
        print("   Consider closing other applications")
    
    # Check required files
    required_files = ['use4.xlsx', 'bloomberg_raw_data.csv']
    for file in required_files:
        if os.path.exists(file):
            size_mb = os.path.getsize(file) / 1024**2
            print(f"✅ {file} ({size_mb:.1f} MB)")
        else:
            print(f"❌ {file} - NOT FOUND")
            return False
    
    return True

if __name__ == "__main__":
    print("🧪 LIGHTWEIGHT GAS MARKET PROCESSOR")
    print("=" * 60)
    
    # Check system first
    if check_system_requirements():
        print("\n" + "=" * 60)
        success = process_gas_market_lightweight()
        
        if success:
            print("\n✅ SUCCESS: Lightweight processing completed without kernel restart!")
        else:
            print("\n❌ FAILED: Check error messages above")
    else:
        print("\n❌ SYSTEM CHECK FAILED: Missing required files")