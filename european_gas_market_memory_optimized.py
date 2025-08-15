# -*- coding: utf-8 -*-
"""
European Gas Market Memory-Optimized Analysis System
===================================================
A memory-efficient version that processes data in chunks to prevent kernel restarts.
This version processes seasonal data one metric at a time and clears memory between operations.

Usage: python european_gas_market_memory_optimized.py
"""

import pandas as pd
import numpy as np
import json
from datetime import datetime
import os
import gc  # For garbage collection

class MemoryOptimizedGasAnalyzer:
    """Memory-optimized European Gas Market Analysis System"""
    
    def __init__(self):
        self.livSheet_file = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
        self.analysis_results_file = 'analysis_results.json'
        self.master_output_file = 'European_Gas_Market_Master.xlsx'
        self.seasonal_output_file = 'European_Gas_Seasonal_Analysis.xlsx'
        
        # Known correct mapping
        self.known_column_mapping = {
            'France': 2, 'Belgium': 3, 'Italy': 4, 'Netherlands': 20, 'GB': 21,
            'Austria': 28, 'Germany': 8, 'Total': 12, 'Industrial': 13, 
            'LDZ': 14, 'Gas-to-Power': 15
        }
        
        self.validation_targets = {
            'Italy': 151.466, 'Netherlands': 90.493, 'GB': 97.740, 'Austria': -9.313, 'Total': 767.693
        }

    def log_memory_usage(self, step=""):
        """Log current memory usage"""
        import psutil
        process = psutil.Process(os.getpid())
        memory_mb = process.memory_info().rss / 1024 / 1024
        print(f"   Memory: {memory_mb:.1f} MB {step}")
        return memory_mb

    def setup_analysis_file(self):
        """Create analysis results file"""
        analysis_data = {
            'target_values': self.validation_targets,
            'column_mapping': self.known_column_mapping
        }
        with open(self.analysis_results_file, 'w') as f:
            json.dump(analysis_data, f, indent=2)
        print(f"✅ Created analysis configuration")

    def load_and_process_core_data(self):
        """Load and process core demand/supply data with memory optimization"""
        
        print("\n📊 LOADING AND PROCESSING CORE DATA")
        print("=" * 60)
        self.log_memory_usage("(start)")
        
        # Load target data
        print("Loading LiveSheet data...")
        target_data = pd.read_excel(self.livSheet_file, 
                                  sheet_name='Daily historic data by category', 
                                  header=None)
        self.log_memory_usage("(after loading)")
        
        # Extract demand data
        print("Processing demand data...")
        demand_df = self.extract_demand_data(target_data)
        self.log_memory_usage("(after demand)")
        
        # Extract supply data  
        print("Processing supply data...")
        supply_df, supply_columns = self.extract_supply_data(target_data)
        self.log_memory_usage("(after supply)")
        
        # Clear target_data to free memory
        del target_data
        gc.collect()
        self.log_memory_usage("(after cleanup)")
        
        return demand_df, supply_df, supply_columns

    def extract_demand_data(self, target_data):
        """Extract demand data efficiently"""
        extracted_data = []
        dates = []
        
        for i in range(12, target_data.shape[0]):
            date_val = target_data.iloc[i, 1]
            if pd.notna(date_val):
                try:
                    date_parsed = pd.to_datetime(str(date_val))
                    dates.append(date_parsed)
                    
                    row_data = {}
                    for key, col_idx in self.known_column_mapping.items():
                        val = target_data.iloc[i, col_idx]
                        row_data[key] = float(val) if pd.notna(val) and isinstance(val, (int, float)) else np.nan
                    
                    extracted_data.append(row_data)
                except:
                    continue
        
        demand_df = pd.DataFrame(extracted_data, index=dates)
        print(f"✅ Demand data: {demand_df.shape}")
        
        # Verify accuracy
        self.verify_accuracy(demand_df)
        return demand_df

    def extract_supply_data(self, target_data):
        """Extract supply data efficiently"""
        supply_columns = {}
        
        # Build supply column mapping
        for j in range(17, 39):
            row10 = str(target_data.iloc[10, j]) if pd.notna(target_data.iloc[10, j]) else ""
            row9 = str(target_data.iloc[9, j]) if pd.notna(target_data.iloc[9, j]) else ""
            
            header_name = row10.replace('nan', '').strip() or row9.replace('nan', '').strip() or f'Col_{j}'
            supply_columns[j] = {'header': header_name}
        
        # Extract supply data
        supply_data = []
        dates = []
        
        for i in range(12, target_data.shape[0]):
            date_val = target_data.iloc[i, 1]
            if pd.notna(date_val):
                try:
                    date_parsed = pd.to_datetime(str(date_val))
                    dates.append(date_parsed)
                    
                    row_data = {}
                    for j in range(17, 39):
                        header = supply_columns[j]['header']
                        val = target_data.iloc[i, j]
                        row_data[header] = float(val) if pd.notna(val) and isinstance(val, (int, float)) else np.nan
                    
                    supply_data.append(row_data)
                except:
                    continue
        
        supply_df = pd.DataFrame(supply_data, index=dates)
        print(f"✅ Supply data: {supply_df.shape}")
        return supply_df, supply_columns

    def verify_accuracy(self, demand_df):
        """Verify demand accuracy"""
        target_date = '2016-10-04'
        if target_date in [str(d)[:10] for d in demand_df.index]:
            target_row = demand_df[demand_df.index.strftime('%Y-%m-%d') == target_date].iloc[0]
            
            matches = 0
            for key, expected in self.validation_targets.items():
                if key in target_row.index and pd.notna(target_row[key]):
                    actual = target_row[key]
                    if abs(actual - expected) < 0.001:
                        matches += 1
            
            accuracy = (matches / len(self.validation_targets)) * 100
            print(f"✅ Accuracy: {accuracy:.1f}% ({matches}/{len(self.validation_targets)} matches)")

    def create_monthly_data(self, demand_df, supply_df):
        """Create monthly aggregations efficiently"""
        
        print("\n📅 CREATING MONTHLY DATA")
        print("=" * 60)
        self.log_memory_usage("(start monthly)")
        
        # Create monthly averages
        demand_monthly = demand_df.resample('ME').mean().round(2)
        supply_monthly = supply_df.resample('ME').mean().round(2)
        
        # Add time columns and reset index
        for df in [demand_monthly, supply_monthly]:
            df['Year'] = df.index.year
            df['Month'] = df.index.month
            df['Year-Month'] = df.index.strftime('%Y-%m')
            df.reset_index(inplace=True)
        
        self.log_memory_usage("(after monthly averages)")
        
        # Create YOY calculations
        demand_yoy = self.create_yoy_data(demand_monthly.copy())
        supply_yoy = self.create_yoy_data(supply_monthly.copy())
        
        self.log_memory_usage("(after YOY)")
        
        print(f"✅ Monthly data created: demand {demand_monthly.shape}, supply {supply_monthly.shape}")
        return demand_monthly, supply_monthly, demand_yoy, supply_yoy

    def create_yoy_data(self, df):
        """Create YOY calculations for a dataframe"""
        date_cols = ['Date', 'Year', 'Month', 'Year-Month']
        data_cols = [col for col in df.columns if col not in date_cols]
        
        for col in data_cols:
            df[f'{col}_YOY_Abs'] = df[col] - df[col].shift(12)
            df[f'{col}_YOY_Pct'] = ((df[col] / df[col].shift(12)) - 1) * 100
        
        # Round and clean
        yoy_cols = [col for col in df.columns if '_YOY_' in col]
        df[yoy_cols] = df[yoy_cols].round(2)
        
        if data_cols:
            df.dropna(subset=[f'{data_cols[0]}_YOY_Abs'], inplace=True)
        
        return df

    def save_master_file(self, demand_df, supply_df, demand_monthly, supply_monthly, demand_yoy, supply_yoy, supply_columns):
        """Save master file efficiently"""
        
        print(f"\n💾 SAVING MASTER FILE")
        print("=" * 60)
        self.log_memory_usage("(start save)")
        
        # Prepare data
        final_demand_df = demand_df.copy().reset_index().rename(columns={'index': 'Date'})
        final_supply_df = supply_df.copy().reset_index().rename(columns={'index': 'Date'})
        
        # Save to Excel
        with pd.ExcelWriter(self.master_output_file, engine='openpyxl') as writer:
            final_demand_df.to_excel(writer, sheet_name='Demand', index=False)
            final_supply_df.to_excel(writer, sheet_name='Supply', index=False)
            demand_monthly.to_excel(writer, sheet_name='Demand Monthly', index=False)
            supply_monthly.to_excel(writer, sheet_name='Supply Monthly', index=False)
            demand_yoy.to_excel(writer, sheet_name='Demand Monthly YOY', index=False)
            supply_yoy.to_excel(writer, sheet_name='Supply Monthly YOY', index=False)
        
        self.log_memory_usage("(after save)")
        print(f"✅ Master file saved: {self.master_output_file}")

    def create_seasonal_data_memory_efficient(self, demand_df, supply_df, demand_monthly, supply_monthly):
        """Create seasonal data one metric at a time to avoid memory issues"""
        
        print(f"\n🌍 CREATING SEASONAL DATA (MEMORY OPTIMIZED)")
        print("=" * 80)
        self.log_memory_usage("(start seasonal)")
        
        # Define metrics to process
        demand_metrics = ['Italy', 'Netherlands', 'GB', 'Austria', 'Total']  # Reduced set for memory
        supply_metrics = ['Norway', 'LNG', 'Total']  # Reduced set for memory
        
        # Filter to existing columns
        demand_metrics = [col for col in demand_metrics if col in demand_df.columns]
        supply_metrics = [col for col in supply_metrics if col in supply_df.columns]
        
        print(f"📊 Processing {len(demand_metrics)} demand + {len(supply_metrics)} supply metrics")
        
        # Process seasonal data metric by metric and save immediately
        seasonal_sheets = {}
        
        # Daily YOY Percentage for demand
        print(f"Creating daily YOY percentage plots...")
        for i, metric in enumerate(demand_metrics):
            print(f"  Processing demand {metric} ({i+1}/{len(demand_metrics)})...")
            yoy_df = self.create_single_daily_yoy_pct(demand_df, metric)
            seasonal_sheets[f'D_YOY_Pct_{metric}'] = yoy_df
            
            # Clear memory every few metrics
            if (i + 1) % 3 == 0:
                gc.collect()
                self.log_memory_usage(f"(after {i+1} demand metrics)")
        
        # Daily YOY Percentage for supply
        for i, metric in enumerate(supply_metrics):
            print(f"  Processing supply {metric} ({i+1}/{len(supply_metrics)})...")
            yoy_df = self.create_single_daily_yoy_pct(supply_df, metric)
            seasonal_sheets[f'D_YOY_Pct_{metric}'] = yoy_df
            
            if (i + 1) % 3 == 0:
                gc.collect()
                self.log_memory_usage(f"(after {i+1} supply metrics)")
        
        # Save seasonal file immediately
        self.save_seasonal_file_efficient(seasonal_sheets)
        
        # Clear all seasonal data from memory
        del seasonal_sheets
        gc.collect()
        self.log_memory_usage("(after seasonal cleanup)")
        
        print(f"✅ Seasonal analysis complete!")

    def create_single_daily_yoy_pct(self, df, metric):
        """Create daily YOY percentage for a single metric"""
        years = list(range(2017, 2031))
        days = list(range(1, 366))
        seasonal_df = pd.DataFrame(index=days, columns=years)
        seasonal_df.index.name = 'Day_of_Year'
        
        # Build data lookup for this metric only
        data_by_day_year = {}
        for _, row in df.iterrows():
            date = row.name if isinstance(row.name, pd.Timestamp) else pd.to_datetime(row.name)
            year = date.year
            
            if date.month == 2 and date.day == 29:
                continue
            
            if date.month > 2 and pd.Timestamp(year, 12, 31).dayofyear == 366:
                day_of_year = date.dayofyear - 1
            else:
                day_of_year = date.dayofyear
            
            if year in years and 1 <= day_of_year <= 365:
                value = row[metric]
                if pd.notna(value):
                    data_by_day_year[(day_of_year, year)] = value
        
        # Calculate YOY percentages
        for day in days:
            for year in years[1:]:
                current_key = (day, year)
                prev_key = (day, year - 1)
                
                if current_key in data_by_day_year and prev_key in data_by_day_year:
                    current_val = data_by_day_year[current_key]
                    prev_val = data_by_day_year[prev_key]
                    if prev_val != 0:
                        yoy_pct = ((current_val / prev_val) - 1) * 100
                        seasonal_df.at[day, year] = round(yoy_pct, 2)
        
        # Add baseline statistics
        baseline_years = [2018, 2019, 2020, 2021]
        baseline_cols = [year for year in baseline_years if year in seasonal_df.columns]
        
        if baseline_cols:
            for col in baseline_cols:
                seasonal_df[col] = pd.to_numeric(seasonal_df[col], errors='coerce')
            
            seasonal_df['Avg_2018-2021'] = seasonal_df[baseline_cols].mean(axis=1, skipna=True).round(2)
            seasonal_df['Min_2018-2021'] = seasonal_df[baseline_cols].min(axis=1, skipna=True).round(2)
            seasonal_df['Max_2018-2021'] = seasonal_df[baseline_cols].max(axis=1, skipna=True).round(2)
            seasonal_df['Range_2018-2021'] = (seasonal_df['Max_2018-2021'] - seasonal_df['Min_2018-2021']).round(2)
        
        return seasonal_df

    def save_seasonal_file_efficient(self, seasonal_sheets):
        """Save seasonal file efficiently"""
        print(f"💾 Saving seasonal file with {len(seasonal_sheets)} sheets...")
        
        with pd.ExcelWriter(self.seasonal_output_file, engine='openpyxl') as writer:
            for sheet_name, df in seasonal_sheets.items():
                # Clean sheet name
                clean_name = sheet_name.replace('/', '_').replace('(', '').replace(')', '')[:31]
                df.to_excel(writer, sheet_name=clean_name)
            
            # Add summary
            summary_data = [{'Sheet': name, 'Type': 'Daily_YOY_Pct'} for name in seasonal_sheets.keys()]
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"✅ Seasonal file saved: {self.seasonal_output_file}")

    def run_memory_optimized_analysis(self):
        """Run the complete analysis with memory optimization"""
        
        print("🌍 MEMORY-OPTIMIZED EUROPEAN GAS MARKET ANALYSIS")
        print("=" * 100)
        print(f"Starting analysis at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 100)
        
        try:
            # Step 1: Setup
            self.setup_analysis_file()
            
            # Step 2: Load and process core data
            demand_df, supply_df, supply_columns = self.load_and_process_core_data()
            
            # Step 3: Create monthly data
            demand_monthly, supply_monthly, demand_yoy, supply_yoy = self.create_monthly_data(
                demand_df, supply_df)
            
            # Step 4: Save master file
            self.save_master_file(demand_df, supply_df, demand_monthly, supply_monthly, 
                                demand_yoy, supply_yoy, supply_columns)
            
            # Step 5: Clear some memory before seasonal analysis
            del demand_monthly, supply_monthly, demand_yoy, supply_yoy
            gc.collect()
            self.log_memory_usage("(before seasonal)")
            
            # Step 6: Create seasonal analysis (memory optimized)
            self.create_seasonal_data_memory_efficient(demand_df, supply_df, None, None)
            
            # Final summary
            print(f"\n🎉 MEMORY-OPTIMIZED ANALYSIS COMPLETE")
            print("=" * 100)
            print(f"✅ Master file: {self.master_output_file}")
            print(f"✅ Seasonal file: {self.seasonal_output_file}")
            print(f"✅ Analysis completed successfully!")
            
            self.log_memory_usage("(final)")
            
            return True
            
        except Exception as e:
            print(f"\n💥 ANALYSIS FAILED: {e}")
            import traceback
            traceback.print_exc()
            return False

def main():
    """Main execution function"""
    analyzer = MemoryOptimizedGasAnalyzer()
    success = analyzer.run_memory_optimized_analysis()
    return success

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)