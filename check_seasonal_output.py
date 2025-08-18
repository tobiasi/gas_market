# Check the seasonal output file structure
import pandas as pd

file_path = 'European_Gas_Seasonal_Analysis.xlsx'

# Get all sheet names
try:
    excel_file = pd.ExcelFile(file_path)
    sheet_names = excel_file.sheet_names
    print(f"Total sheets: {len(sheet_names)}")
    print("\nSheet names:")
    for i, sheet in enumerate(sheet_names, 1):
        print(f"{i:2d}. {sheet}")
    
    # Check a sample sheet structure
    if len(sheet_names) > 1:
        sample_sheet = sheet_names[0]  # First non-summary sheet
        print(f"\nSample sheet '{sample_sheet}' structure:")
        sample_df = pd.read_excel(file_path, sheet_name=sample_sheet, index_col=0)
        print(f"Shape: {sample_df.shape}")
        print("Columns:")
        for col in sample_df.columns:
            print(f"  - {col}")
        print(f"\nFirst few rows and columns:")
        print(sample_df.iloc[:5, :5])

except Exception as e:
    print(f"Error: {e}")