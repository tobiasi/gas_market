# Memory optimization techniques for the full Bloomberg system

import pandas as pd
import numpy as np
import gc

def memory_optimized_bloomberg_system():
    """
    Key memory optimization techniques to prevent kernel restart:
    """
    
    # 1. CHUNKED DATA LOADING
    def load_data_in_chunks():
        """Load large CSV files in chunks"""
        chunks = []
        for chunk in pd.read_csv('bloomberg_raw_data.csv', chunksize=1000):
            chunks.append(chunk)
            if len(chunks) >= 3:  # Limit data size
                break
        return pd.concat(chunks)
    
    # 2. MEMORY CLEANUP
    def cleanup_memory():
        """Force garbage collection"""
        gc.collect()
        
    # 3. PROCESS IN BATCHES
    def process_countries_individually():
        """Process one country at a time instead of all together"""
        countries = ['Italy', 'Germany', 'France']  # Process subset
        results = {}
        
        for country in countries:
            # Process country data
            country_data = process_single_country(country)
            results[country] = country_data
            
            # Clean up after each country
            cleanup_memory()
        
        return results
    
    # 4. USE EFFICIENT DATA TYPES
    def optimize_datatypes(df):
        """Convert to more efficient data types"""
        for col in df.select_dtypes(include=['float64']):
            df[col] = df[col].astype('float32')  # Half the memory
        return df
    
    # 5. DELETE LARGE OBJECTS
    def safe_delete(obj_name, locals_dict):
        """Safely delete large objects"""
        if obj_name in locals_dict:
            del locals_dict[obj_name]
            cleanup_memory()

# APPLY THESE TO YOUR ORIGINAL CODE:

def fixed_download_bloomberg_data(tickers):
    """Memory-optimized Bloomberg data loading"""
    try:
        # Load in smaller chunks
        data = pd.read_csv('bloomberg_raw_data.csv', index_col=0, parse_dates=True, nrows=1000)  # Limit rows
        print(f"✅ Loaded limited data to prevent memory overload: {data.shape}")
        return data
    except Exception as e:
        print(f"❌ Memory-safe loading failed: {e}")
        return None

def fixed_create_multiindex_data(data, dataset):
    """Memory-optimized MultiIndex creation"""
    
    # Limit to essential tickers only
    essential_tickers = []
    for _, row in dataset.iterrows():
        ticker = row.get('Ticker')
        region = row.get('Region from', '')
        category = row.get('Category', '')
        
        # Only include key demand tickers
        if ticker and ticker in data.columns and category == 'Demand' and region in ['Italy', 'Germany', 'France']:
            essential_tickers.append(ticker)
            
        if len(essential_tickers) >= 50:  # Limit to 50 tickers
            break
    
    print(f"Processing {len(essential_tickers)} essential tickers (memory optimized)")
    
    # Work with subset only
    data_subset = data[essential_tickers].copy()
    
    # Clean up original large DataFrame
    del data
    gc.collect()
    
    return data_subset  # Return simplified version

# MEMORY MONITORING
def check_memory_usage():
    """Monitor memory usage during processing"""
    import psutil
    process = psutil.Process()
    memory_mb = process.memory_info().rss / 1024 / 1024
    print(f"Memory usage: {memory_mb:.1f} MB")
    
    if memory_mb > 1000:  # Over 1GB
        print("⚠️  HIGH MEMORY USAGE - Consider reducing data size")
        return False
    return True

if __name__ == "__main__":
    print("💡 MEMORY OPTIMIZATION TECHNIQUES")
    print("Apply these patterns to prevent kernel restart:")
    print("1. Load data in chunks")
    print("2. Process countries individually") 
    print("3. Delete large objects with gc.collect()")
    print("4. Use efficient data types (float32 vs float64)")
    print("5. Limit data size (rows/columns)")
    print("6. Monitor memory usage during processing")