# Check a YOY percentage sheet structure
import pandas as pd

file_path = 'European_Gas_Seasonal_Analysis.xlsx'

# Check a YOY percentage sheet
sheet_name = 'D_YOY_Pct_Italy'
print(f"Sheet: {sheet_name}")

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)
    print(f"Shape: {df.shape}")
    print("Columns:")
    for col in df.columns:
        print(f"  - {col}")
    
    print(f"\nFirst few rows and columns (YOY % values):")
    print(df.iloc[:5, :5])
    
    # Check baseline columns
    print(f"\nBaseline statistics columns:")
    baseline_cols = [col for col in df.columns if '2018-2021' in str(col)]
    for col in baseline_cols:
        print(f"  - {col}")
    
    if baseline_cols:
        print(f"\nSample baseline values:")
        print(df[baseline_cols].head())

except Exception as e:
    print(f"Error: {e}")