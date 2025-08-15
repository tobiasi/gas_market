# Check how the columns sum in the original LiveSheet
import pandas as pd
import numpy as np

def check_livesheet_totals():
    """Check the original LiveSheet to see how columns sum"""
    
    print("🔍 CHECKING ORIGINAL LIVESHEET COLUMN SUMS")
    print("=" * 70)
    
    try:
        # Load the original LiveSheet
        filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
        target_data = pd.read_excel(filename, sheet_name='Daily historic data by category', header=None)
        
        print(f"✅ Loaded LiveSheet: {target_data.shape}")
        
        # Show the header structure around our columns
        print(f"\n📋 HEADER STRUCTURE (rows 9-11):")
        print("Column indexes we're using:")
        
        our_mapping = {
            'France': 2, 'Belgium': 3, 'Italy': 4, 'Netherlands': 20, 'GB': 21,
            'Austria': 28, 'Germany': 8, 'Total': 12, 'Industrial': 13, 'LDZ': 14, 'Gas-to-Power': 15
        }
        
        for name, col_idx in our_mapping.items():
            row9 = str(target_data.iloc[9, col_idx]) if pd.notna(target_data.iloc[9, col_idx]) else "empty"
            row10 = str(target_data.iloc[10, col_idx]) if pd.notna(target_data.iloc[10, col_idx]) else "empty"
            row11 = str(target_data.iloc[11, col_idx]) if pd.notna(target_data.iloc[11, col_idx]) else "empty"
            
            print(f"  {name:12} (col {col_idx:2d}): {row9} | {row10} | {row11}")
        
        # Check a specific date - Oct 4, 2016 (row we validated)
        print(f"\n📊 CHECKING SPECIFIC DATE: 2016-10-04 (around row with date)")
        
        # Find the row with 2016-10-04
        target_row_idx = None
        for i in range(12, min(50, target_data.shape[0])):  # Search first ~40 data rows
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
            
            # Extract the values
            row_values = {}
            for name, col_idx in our_mapping.items():
                val = target_data.iloc[target_row_idx, col_idx]
                row_values[name] = float(val) if pd.notna(val) and isinstance(val, (int, float)) else np.nan
                print(f"  {name:12}: {val}")
            
            # Calculate sums
            countries = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany']
            categories = ['Industrial', 'LDZ', 'Gas-to-Power']
            
            countries_sum = sum(row_values[c] for c in countries if not pd.isna(row_values[c]))
            categories_sum = sum(row_values[c] for c in categories if not pd.isna(row_values[c]))
            total_val = row_values['Total']
            
            print(f"\n📊 SUMS ANALYSIS:")
            print(f"  Countries sum:  {countries_sum:.3f}")
            print(f"  Categories sum: {categories_sum:.3f}")
            print(f"  Total column:   {total_val:.3f}")
            print(f"  Countries vs Total diff:  {total_val - countries_sum:.3f}")
            print(f"  Categories vs Total diff: {total_val - categories_sum:.3f}")
            
        # Check if there are other country columns we might be missing
        print(f"\n🔍 SCANNING FOR ADDITIONAL COUNTRY COLUMNS")
        print("=" * 50)
        
        # Look at row 10 headers for country names around our range
        print("Headers around demand columns (columns 2-30):")
        for col in range(2, 31):
            if col not in [c for c in our_mapping.values()]:  # Skip columns we already know
                row9 = str(target_data.iloc[9, col]) if pd.notna(target_data.iloc[9, col]) else ""
                row10 = str(target_data.iloc[10, col]) if pd.notna(target_data.iloc[10, col]) else ""
                
                row9_clean = row9.replace('nan', '').strip()
                row10_clean = row10.replace('nan', '').strip()
                
                if row10_clean and row10_clean not in ['', 'nan']:
                    print(f"  Col {col:2d}: {row10_clean} ({row9_clean})")
        
        # Look at totals across several rows to see the pattern
        print(f"\n📈 PATTERN ANALYSIS (first 10 data rows)")
        print("=" * 60)
        print("Row    Countries_Sum   Categories_Sum   Total       C_Diff   Cat_Diff")
        print("-" * 70)
        
        for i in range(12, min(22, target_data.shape[0])):
            date_val = target_data.iloc[i, 1]
            if pd.notna(date_val):
                try:
                    # Extract values
                    vals = {}
                    for name, col_idx in our_mapping.items():
                        val = target_data.iloc[i, col_idx]
                        vals[name] = float(val) if pd.notna(val) and isinstance(val, (int, float)) else 0
                    
                    countries_sum = sum(vals[c] for c in countries)
                    categories_sum = sum(vals[c] for c in categories)
                    total_val = vals['Total']
                    
                    c_diff = total_val - countries_sum
                    cat_diff = total_val - categories_sum
                    
                    print(f"{i:3d}    {countries_sum:10.2f}    {categories_sum:11.2f}    {total_val:8.2f}    {c_diff:7.2f}    {cat_diff:8.2f}")
                    
                except:
                    continue
        
        # Check if we might have wrong column mapping
        print(f"\n🔍 VERIFYING OUR COLUMN MAPPING")
        print("=" * 50)
        print("Let's double-check our known validation point...")
        
        # We know Italy should be 151.466 on 2016-10-04
        if target_row_idx:
            # Check surrounding columns to see if we have the right Italy column
            print(f"Values around Italy column (4) on 2016-10-04:")
            for col in range(2, 8):
                val = target_data.iloc[target_row_idx, col]
                header = str(target_data.iloc[10, col]) if pd.notna(target_data.iloc[10, col]) else "empty"
                print(f"  Col {col}: {val} (header: {header})")
                
                # Check if this matches our validation
                if pd.notna(val) and isinstance(val, (int, float)):
                    if abs(float(val) - 151.466) < 0.001:
                        print(f"    ✅ This matches our Italy validation value!")
        
    except Exception as e:
        print(f"❌ Error checking LiveSheet: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_livesheet_totals()