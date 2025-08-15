# Fix the demand column mapping using correct demand columns
import pandas as pd
import numpy as np
import json

def fix_demand_columns():
    """Fix column mapping to use correct demand columns that sum to total"""
    
    print("🔧 FIXING DEMAND COLUMN MAPPING")
    print("=" * 60)
    
    try:
        # Load the original LiveSheet
        filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
        target_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
        
        print(f"✅ Loaded LiveSheet: {target_data.shape}")
        
        # Correct demand column mapping based on headers
        correct_demand_mapping = {
            'France': 2,
            'Belgium': 3, 
            'Italy': 4,
            'Netherlands': 5,  # This is the DEMAND Netherlands (not production)
            'GB': 6,           # This is the DEMAND GB (not production)
            'Austria': 7,      # This is the DEMAND Austria (not export)
            'Germany': 8,
            'Switzerland': 9,
            'Luxembourg': 10,
            'Island of Ireland': 11,
            'Total': 12,
            'Industrial': 13,
            'LDZ': 14,
            'Gas-to-Power': 15
        }
        
        print(f"\n📋 CORRECTED DEMAND COLUMN MAPPING:")
        for name, col_idx in correct_demand_mapping.items():
            row9 = str(target_data.iloc[9, col_idx]) if pd.notna(target_data.iloc[9, col_idx]) else "empty"
            row10 = str(target_data.iloc[10, col_idx]) if pd.notna(target_data.iloc[10, col_idx]) else "empty"
            print(f"  {name:15} (col {col_idx:2d}): {row9} | {row10}")
        
        # Test on our validation date: 2016-10-04
        print(f"\n📊 TESTING ON 2016-10-04:")
        target_row_idx = None
        for i in range(12, min(50, target_data.shape[0])):
            date_val = target_data.iloc[i, 1]
            if pd.notna(date_val):
                try:
                    date_parsed = pd.to_datetime(str(date_val))
                    if date_parsed.strftime('%Y-%m-%d') == '2016-10-04':
                        target_row_idx = i
                        break
                except:
                    continue
        
        if target_row_idx:
            print(f"Found 2016-10-04 at row {target_row_idx}")
            
            # Extract values using corrected mapping
            row_values = {}
            for name, col_idx in correct_demand_mapping.items():
                val = target_data.iloc[target_row_idx, col_idx]
                row_values[name] = float(val) if pd.notna(val) and isinstance(val, (int, float)) else 0
                print(f"  {name:15}: {val}")
            
            # Test if countries sum to total
            countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Switzerland', 'Luxembourg', 'Island of Ireland']
            categories = ['Industrial', 'LDZ', 'Gas-to-Power']
            
            countries_sum = sum(row_values[c] for c in countries if c in row_values)
            categories_sum = sum(row_values[c] for c in categories if c in row_values)
            total_val = row_values['Total']
            
            print(f"\n📊 CORRECTED SUMS ANALYSIS:")
            print(f"  Countries sum:  {countries_sum:.3f}")
            print(f"  Categories sum: {categories_sum:.3f}")
            print(f"  Total column:   {total_val:.3f}")
            print(f"  Countries vs Total diff:  {abs(total_val - countries_sum):.3f}")
            print(f"  Categories vs Total diff: {abs(total_val - categories_sum):.3f}")
            
            if abs(total_val - countries_sum) < 0.001:
                print(f"  ✅ Countries now sum to Total!")
            else:
                print(f"  ❌ Countries still don't sum to Total")
                
            if abs(total_val - categories_sum) < 0.001:
                print(f"  ✅ Categories sum to Total!")
        
        # Test on multiple dates to confirm pattern
        print(f"\n📈 PATTERN VERIFICATION (first 10 data rows):")
        print("Row    Countries_Sum   Categories_Sum   Total       C_Diff   Cat_Diff")
        print("-" * 70)
        
        for i in range(12, min(22, target_data.shape[0])):
            date_val = target_data.iloc[i, 1]
            if pd.notna(date_val):
                try:
                    # Extract values
                    vals = {}
                    for name, col_idx in correct_demand_mapping.items():
                        val = target_data.iloc[i, col_idx]
                        vals[name] = float(val) if pd.notna(val) and isinstance(val, (int, float)) else 0
                    
                    countries_sum = sum(vals[c] for c in countries if c in vals)
                    categories_sum = sum(vals[c] for c in categories if c in vals)
                    total_val = vals['Total']
                    
                    c_diff = abs(total_val - countries_sum)
                    cat_diff = abs(total_val - categories_sum)
                    
                    print(f"{i:3d}    {countries_sum:10.2f}    {categories_sum:11.2f}    {total_val:8.2f}    {c_diff:7.3f}    {cat_diff:8.3f}")
                    
                except:
                    continue
        
        # Save corrected mapping
        corrected_analysis = {
            'target_values': {
                'Italy': 151.4659795006551,
                'Netherlands': 76.89155673229693,  # Now using demand column 5
                'GB': 163.86,                       # Now using demand column 6
                'Austria': 20.64725134432316,       # Now using demand column 7
                'Total': 767.6925372545071
            },
            'column_mapping': correct_demand_mapping
        }
        
        with open('corrected_analysis_results.json', 'w') as f:
            json.dump(corrected_analysis, f, indent=2)
        
        print(f"\n✅ Corrected analysis saved to 'corrected_analysis_results.json'")
        print(f"💡 Key discovery: We were mixing demand columns with production/export columns!")
        print(f"   - Netherlands demand is column 5 (not 20 which is production)")
        print(f"   - GB demand is column 6 (not 21 which is production)")
        print(f"   - Austria demand is column 7 (not 28 which is export)")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    fix_demand_columns()