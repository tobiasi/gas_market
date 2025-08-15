# Debug script to check memory usage and identify bottlenecks
import pandas as pd
import psutil
import os

def check_memory():
    """Check current memory usage"""
    process = psutil.Process(os.getpid())
    memory_info = process.memory_info()
    memory_mb = memory_info.rss / 1024 / 1024
    
    print(f"Current memory usage: {memory_mb:.1f} MB")
    
    # System memory
    system_memory = psutil.virtual_memory()
    print(f"System memory - Total: {system_memory.total / 1024**3:.1f} GB")
    print(f"System memory - Available: {system_memory.available / 1024**3:.1f} GB")
    print(f"System memory - Used: {system_memory.percent:.1f}%")
    
    return memory_mb

def test_file_loading():
    """Test loading the main file to see memory impact"""
    print("🔍 TESTING FILE LOADING AND MEMORY USAGE")
    print("=" * 60)
    
    initial_memory = check_memory()
    
    try:
        print("\n1. Testing LiveSheet file loading...")
        filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
        
        if os.path.exists(filename):
            file_size = os.path.getsize(filename) / 1024 / 1024
            print(f"   File size: {file_size:.1f} MB")
            
            # Load file
            print("   Loading file...")
            target_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
            print(f"   Data shape: {target_data.shape}")
            
            after_load_memory = check_memory()
            memory_increase = after_load_memory - initial_memory
            print(f"   Memory increase: {memory_increase:.1f} MB")
            
            # Test processing a small subset
            print("\n2. Testing data processing on subset...")
            subset_data = target_data.iloc[:100, :50]  # Small subset
            
            # Create small demand dataframe
            demand_data = []
            for i in range(12, min(112, target_data.shape[0])):  # First 100 rows
                date_val = target_data.iloc[i, 1]
                if pd.notna(date_val):
                    try:
                        date_parsed = pd.to_datetime(str(date_val))
                        demand_data.append({
                            'Date': date_parsed,
                            'Italy': target_data.iloc[i, 4],
                            'Total': target_data.iloc[i, 12]
                        })
                    except:
                        continue
            
            if demand_data:
                small_df = pd.DataFrame(demand_data)
                print(f"   Created small DataFrame: {small_df.shape}")
                
                after_process_memory = check_memory()
                process_increase = after_process_memory - after_load_memory
                print(f"   Processing memory increase: {process_increase:.1f} MB")
                
                # Estimate full processing memory
                full_rows = target_data.shape[0] - 12
                estimated_memory = process_increase * (full_rows / len(demand_data))
                print(f"   Estimated full processing memory: {estimated_memory:.1f} MB")
                
                total_estimated = after_load_memory + estimated_memory
                print(f"   Total estimated memory needed: {total_estimated:.1f} MB")
                
                # Check if this might cause issues
                available_mb = psutil.virtual_memory().available / 1024 / 1024
                if total_estimated > available_mb * 0.8:  # Using 80% threshold
                    print("   ⚠️  WARNING: Estimated memory usage may cause issues!")
                    print("   💡 Recommendations:")
                    print("      - Process data in smaller chunks")
                    print("      - Reduce number of simultaneous DataFrames")
                    print("      - Use data types optimization")
                else:
                    print("   ✅ Memory usage should be acceptable")
            
        else:
            print(f"   ❌ File not found: {filename}")
    
    except Exception as e:
        print(f"   ❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()

def suggest_solutions():
    """Suggest solutions based on system analysis"""
    print(f"\n💡 SOLUTIONS TO PREVENT KERNEL RESTART")
    print("=" * 60)
    
    print("1. **Memory Optimization:**")
    print("   - Process data in smaller chunks")
    print("   - Clear variables with `del` when no longer needed")
    print("   - Use `pd.read_excel(chunksize=...)` for large files")
    
    print("\n2. **Data Type Optimization:**")
    print("   - Convert float64 to float32 where possible")
    print("   - Use category data type for repeated strings")
    
    print("\n3. **Processing Strategy:**")
    print("   - Process demand and supply separately")
    print("   - Create seasonal plots one metric at a time")
    print("   - Save intermediate results and reload")
    
    print("\n4. **Environment Settings:**")
    print("   - Increase Jupyter memory limits")
    print("   - Run in terminal instead of notebook")
    print("   - Close other applications to free memory")
    
    print("\n5. **Alternative Approach:**")
    print("   - Run the modular scripts separately:")
    print("     python gas_market_master.py")
    print("     python create_monthly_tabs.py")
    print("     python create_seasonal_plots.py")

if __name__ == "__main__":
    test_file_loading()
    suggest_solutions()