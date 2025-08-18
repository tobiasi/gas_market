# Check if demand country columns match with total column
import pandas as pd
import numpy as np

def check_demand_totals():
    """Check if country columns sum to total in demand data"""
    
    print("🔍 CHECKING DEMAND DATA COLUMN MATCHING")
    print("=" * 60)
    
    try:
        # Load demand data
        master_file = 'European_Gas_Market_Master.xlsx'
        demand_df = pd.read_excel(master_file, sheet_name='Demand')
        demand_df['Date'] = pd.to_datetime(demand_df['Date'])
        
        print(f"✅ Loaded demand data: {demand_df.shape}")
        print(f"📅 Date range: {demand_df['Date'].min()} to {demand_df['Date'].max()}")
        
        # Show column names
        print(f"\n📊 DEMAND COLUMNS:")
        for i, col in enumerate(demand_df.columns, 1):
            print(f"{i:2d}. {col}")
        
        # Check a few sample rows
        print(f"\n📋 SAMPLE DATA (first 5 rows):")
        sample_cols = ['Date', 'France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany', 'Total']
        available_cols = [col for col in sample_cols if col in demand_df.columns]
        print(demand_df[available_cols].head().to_string(index=False))
        
        # Check what the Total column represents
        print(f"\n🔍 ANALYZING TOTAL COLUMN COMPOSITION")
        print("=" * 50)
        
        # Get country columns (excluding Total, Industrial, LDZ, Gas-to-Power)
        country_columns = ['France', 'Belgium', 'Italy', 'Netherlands', 'GB', 'Austria', 'Germany']
        available_countries = [col for col in country_columns if col in demand_df.columns]
        
        print(f"Country columns found: {available_countries}")
        
        # Calculate sum of countries for each row
        test_rows = demand_df.head(10).copy()
        
        if available_countries:
            test_rows['Countries_Sum'] = test_rows[available_countries].sum(axis=1)
            
            print(f"\nComparison for first 10 rows:")
            print("Date                Countries_Sum    Total       Difference")
            print("-" * 60)
            
            for _, row in test_rows.iterrows():
                countries_sum = row['Countries_Sum']
                total_val = row['Total']
                diff = total_val - countries_sum
                date_str = row['Date'].strftime('%Y-%m-%d')
                print(f"{date_str}    {countries_sum:10.2f}    {total_val:8.2f}    {diff:10.2f}")
        
        # Check if Total might be sum of categories instead
        print(f"\n🔍 CHECKING IF TOTAL = INDUSTRIAL + LDZ + GAS-TO-POWER")
        print("=" * 60)
        
        category_columns = ['Industrial', 'LDZ', 'Gas-to-Power']
        available_categories = [col for col in category_columns if col in demand_df.columns]
        
        if available_categories:
            test_rows['Categories_Sum'] = test_rows[available_categories].sum(axis=1)
            
            print(f"Category columns found: {available_categories}")
            print(f"\nComparison for first 10 rows:")
            print("Date                Categories_Sum   Total       Difference")
            print("-" * 60)
            
            for _, row in test_rows.iterrows():
                categories_sum = row['Categories_Sum']
                total_val = row['Total']
                diff = total_val - categories_sum
                date_str = row['Date'].strftime('%Y-%m-%d')
                print(f"{date_str}    {categories_sum:11.2f}    {total_val:8.2f}    {diff:10.2f}")
        
        # Statistical analysis
        print(f"\n📊 STATISTICAL ANALYSIS")
        print("=" * 40)
        
        if available_countries:
            countries_sum = demand_df[available_countries].sum(axis=1)
            total_col = demand_df['Total']
            
            diff_countries = total_col - countries_sum
            
            print(f"Countries sum vs Total:")
            print(f"  Mean difference: {diff_countries.mean():.2f}")
            print(f"  Std difference: {diff_countries.std():.2f}")
            print(f"  Min difference: {diff_countries.min():.2f}")
            print(f"  Max difference: {diff_countries.max():.2f}")
            
            # Check correlation
            correlation = countries_sum.corr(total_col)
            print(f"  Correlation: {correlation:.4f}")
        
        if available_categories:
            categories_sum = demand_df[available_categories].sum(axis=1)
            total_col = demand_df['Total']
            
            diff_categories = total_col - categories_sum
            
            print(f"\nCategories sum vs Total:")
            print(f"  Mean difference: {diff_categories.mean():.2f}")
            print(f"  Std difference: {diff_categories.std():.2f}")
            print(f"  Min difference: {diff_categories.min():.2f}")
            print(f"  Max difference: {diff_categories.max():.2f}")
            
            # Check correlation
            correlation = categories_sum.corr(total_col)
            print(f"  Correlation: {correlation:.4f}")
            
            # Check if they match closely
            close_matches = (abs(diff_categories) < 1.0).sum()
            total_rows = len(demand_df)
            match_percentage = (close_matches / total_rows) * 100
            print(f"  Rows matching within 1.0: {close_matches}/{total_rows} ({match_percentage:.1f}%)")
        
        # Check what the LiveSheet documentation says
        print(f"\n📖 INTERPRETING RESULTS")
        print("=" * 40)
        
        if available_countries and available_categories:
            countries_sum_sample = demand_df[available_countries].sum(axis=1).mean()
            categories_sum_sample = demand_df[available_categories].sum(axis=1).mean()
            total_sample = demand_df['Total'].mean()
            
            print(f"Average values:")
            print(f"  Countries sum: {countries_sum_sample:.2f}")
            print(f"  Categories sum: {categories_sum_sample:.2f}")
            print(f"  Total column: {total_sample:.2f}")
            
            if abs(categories_sum_sample - total_sample) < abs(countries_sum_sample - total_sample):
                print(f"\n💡 CONCLUSION: Total appears to be sum of CATEGORIES (Industrial + LDZ + Gas-to-Power)")
                print(f"   Countries represent geographical breakdown")
                print(f"   Categories represent consumption type breakdown")
                print(f"   Both are different views of the same total demand")
            else:
                print(f"\n💡 CONCLUSION: Total appears to be sum of COUNTRIES")
    
    except Exception as e:
        print(f"❌ Error checking demand totals: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_demand_totals()