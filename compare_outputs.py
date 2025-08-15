# Compare outputs from different Bloomberg processing methods
import pandas as pd
import numpy as np
import json

def compare_bloomberg_outputs():
    """Compare outputs from different Bloomberg processing methods"""
    
    print("🔍 COMPARING BLOOMBERG OUTPUT METHODS")
    print("=" * 60)
    
    # Load reference values (our validated targets)
    with open('corrected_analysis_results.json', 'r') as f:
        analysis = json.load(f)
    target_values = analysis['target_values']
    
    print("📋 REFERENCE VALUES (LiveSheet validated):")
    for key, value in target_values.items():
        print(f"  {key}: {value:.6f}")
    
    # 1. Load chunked Bloomberg output (latest)
    print(f"\n1️⃣ CHUNKED BLOOMBERG OUTPUT:")
    try:
        chunked_data = pd.read_excel('European_Gas_Market_Master.xlsx', sheet_name='Demand')
        chunked_data['Date'] = pd.to_datetime(chunked_data['Date'])
        
        target_date = '2016-10-04'
        chunked_row = chunked_data[chunked_data['Date'].dt.strftime('%Y-%m-%d') == target_date].iloc[0]
        
        print(f"  Shape: {chunked_data.shape}")
        print(f"  Date range: {chunked_data['Date'].min()} to {chunked_data['Date'].max()}")
        print(f"  Values on {target_date}:")
        
        chunked_values = {}
        for key in target_values.keys():
            if key in chunked_row.index:
                val = chunked_row[key]
                chunked_values[key] = val
                print(f"    {key}: {val:.6f}")
        
    except Exception as e:
        print(f"  ❌ Error loading chunked output: {e}")
        chunked_values = {}
    
    # 2. Load lightweight output (previous working)
    print(f"\n2️⃣ LIGHTWEIGHT OUTPUT:")
    try:
        lightweight_data = pd.read_excel('European_Gas_Lightweight.xlsx')
        lightweight_data['Date'] = pd.to_datetime(lightweight_data['Date'])
        
        lightweight_row = lightweight_data[lightweight_data['Date'].dt.strftime('%Y-%m-%d') == target_date].iloc[0]
        
        print(f"  Shape: {lightweight_data.shape}")
        print(f"  Values on {target_date}:")
        
        lightweight_values = {}
        for col in lightweight_row.index:
            if col != 'Date':
                val = lightweight_row[col]
                lightweight_values[col] = val
                print(f"    {col}: {val:.6f}")
        
    except Exception as e:
        print(f"  ❌ Error loading lightweight output: {e}")
        lightweight_values = {}
    
    # 3. Load test output (from validation)
    print(f"\n3️⃣ TEST OUTPUT:")
    try:
        test_data = pd.read_excel('European_Gas_Market_TEST.xlsx', sheet_name='Demand')
        test_data['Date'] = pd.to_datetime(test_data['Date'])
        
        test_row = test_data[test_data['Date'].dt.strftime('%Y-%m-%d') == target_date].iloc[0]
        
        print(f"  Shape: {test_data.shape}")
        print(f"  Values on {target_date}:")
        
        test_values = {}
        for key in target_values.keys():
            if key in test_row.index:
                val = test_row[key]
                test_values[key] = val
                print(f"    {key}: {val:.6f}")
        
    except Exception as e:
        print(f"  ❌ Error loading test output: {e}")
        test_values = {}
    
    # 4. Comparison analysis
    print(f"\n📊 DETAILED COMPARISON:")
    print(f"{'Metric':<15} {'Reference':<12} {'Chunked':<12} {'Lightweight':<12} {'Test':<12} {'Status'}")
    print("-" * 75)
    
    all_match = True
    
    for key in target_values.keys():
        reference = target_values[key]
        chunked = chunked_values.get(key, np.nan)
        lightweight = lightweight_values.get(key, np.nan)
        test = test_values.get(key, np.nan)
        
        # Check if values match (within small tolerance)
        chunked_match = abs(chunked - reference) < 1.0 if not pd.isna(chunked) else False
        test_match = abs(test - reference) < 1.0 if not pd.isna(test) else False
        
        # For Italy, check if chunked matches test (both should be ~150.84)
        if key == 'Italy':
            italy_consistency = abs(chunked - test) < 0.1 if not pd.isna(chunked) and not pd.isna(test) else False
            status = "✅" if italy_consistency else "❌"
        else:
            status = "✅" if chunked_match and test_match else "❌"
        
        if not (chunked_match and test_match):
            all_match = False
        
        print(f"{key:<15} {reference:<12.2f} {chunked:<12.2f} {lightweight:<12.2f} {test:<12.2f} {status}")
    
    # 5. Cross-method consistency check
    print(f"\n🔍 CROSS-METHOD CONSISTENCY:")
    
    # Check if chunked and test methods give same results
    print(f"  Chunked vs Test (Italy): {chunked_values.get('Italy', 0):.2f} vs {test_values.get('Italy', 0):.2f}")
    italy_diff = abs(chunked_values.get('Italy', 0) - test_values.get('Italy', 0))
    print(f"  Italy difference: {italy_diff:.6f}")
    
    if italy_diff < 0.01:
        print(f"  ✅ PERFECT: Chunked and Test methods give identical results!")
    elif italy_diff < 0.1:
        print(f"  ✅ EXCELLENT: Chunked and Test methods are highly consistent")
    else:
        print(f"  ❌ INCONSISTENT: Methods giving different results")
    
    # 6. Check data completeness
    print(f"\n📈 DATA COMPLETENESS:")
    
    datasets = [
        ("Chunked", chunked_data if 'chunked_data' in locals() else None),
        ("Test", test_data if 'test_data' in locals() else None),
        ("Lightweight", lightweight_data if 'lightweight_data' in locals() else None)
    ]
    
    for name, df in datasets:
        if df is not None:
            print(f"  {name}: {df.shape[0]} rows, {df.shape[1]} columns")
            
            # Check for missing data
            missing_pct = (df.isnull().sum().sum() / (df.shape[0] * df.shape[1])) * 100
            print(f"    Missing data: {missing_pct:.1f}%")
            
            # Check countries sum to total
            if 'Total' in df.columns:
                country_cols = [col for col in df.columns if col in ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']]
                if country_cols:
                    sample_row = df.iloc[0]
                    countries_sum = sample_row[country_cols].sum()
                    total_val = sample_row['Total']
                    sum_diff = abs(total_val - countries_sum)
                    print(f"    Countries sum to Total: {sum_diff:.2f} difference")
        else:
            print(f"  {name}: NOT AVAILABLE")
    
    # 7. Final assessment
    print(f"\n🎯 FINAL ASSESSMENT:")
    
    if all_match:
        print(f"  ✅ ALL METHODS CONSISTENT: Results match within tolerance")
    else:
        print(f"  ⚠️  SOME DIFFERENCES: Check individual metrics above")
    
    # Specific assessment for production readiness
    chunked_italy = chunked_values.get('Italy', 0)
    target_italy = target_values['Italy']
    
    print(f"\n🚀 PRODUCTION READINESS:")
    print(f"  Chunked Italy: {chunked_italy:.2f}")
    print(f"  Target Italy: {target_italy:.2f}")
    print(f"  Difference: {abs(chunked_italy - target_italy):.2f}")
    
    if abs(chunked_italy - target_italy) < 1.0:
        print(f"  ✅ CHUNKED METHOD IS PRODUCTION READY!")
        print(f"     Achieves excellent accuracy while preventing kernel restart")
    else:
        print(f"  ❌ CHUNKED METHOD NEEDS ADJUSTMENT")
    
    return chunked_values, test_values, lightweight_values

if __name__ == "__main__":
    chunked_vals, test_vals, light_vals = compare_bloomberg_outputs()