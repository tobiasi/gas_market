#!/usr/bin/env python3
"""
LiveSheet Supply Replicator
===========================

Replicates all supply columns in the LiveSheet using validated SUMIFS logic.
First calculates each supply source individually, then aggregates to total.

Key findings applied:
1. Raw MultiTicker values are already in correct units (MCM/d)
2. No scaling factors needed (they're metadata only)
3. Use exact 3-level criteria matching from LiveSheet
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

class LiveSheetSupplyReplicator:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.multiticker_df = None
        self.livesheet_df = None
        self.supply_results = None
        
        # Define the 18 supply routes with their exact criteria
        self.supply_routes = [
            {'name': 'Slovakia_Austria', 'col': 'R', 'criteria': ('Import', 'Slovakia', 'Austria')},
            {'name': 'Russia_NordStream_Germany', 'col': 'S', 'criteria': ('Import', 'Russia (Nord Stream)', 'Germany')},
            {'name': 'Norway_Europe', 'col': 'T', 'criteria': ('Import', 'Norway', 'Europe')},
            {'name': 'Netherlands_Production', 'col': 'U', 'criteria': ('Production', 'Netherlands', 'Netherlands')},
            {'name': 'GB_Production', 'col': 'V', 'criteria': ('Production', 'GB', 'GB')},
            {'name': 'LNG_Total', 'col': 'W', 'criteria': ('Import', 'LNG', '*')},  # Wildcard
            {'name': 'Algeria_Italy', 'col': 'X', 'criteria': ('Import', 'Algeria', 'Italy')},
            {'name': 'Libya_Italy', 'col': 'Y', 'criteria': ('Import', 'Libya', 'Italy')},
            {'name': 'Spain_France', 'col': 'Z', 'criteria': ('Import', 'Spain', 'France')},
            {'name': 'Denmark_Germany', 'col': 'AA', 'criteria': ('Import', 'Denmark', 'Germany')},
            {'name': 'Czech_Poland_Germany', 'col': 'AB', 'criteria': ('Import', 'Czech and Poland', 'Germany')},
            {'name': 'Austria_Hungary_Export', 'col': 'AC', 'criteria': ('Export', 'Austria', 'Hungary')},
            {'name': 'Slovenia_Austria', 'col': 'AD', 'criteria': ('Import', 'Slovenia', 'Austria')},
            {'name': 'MAB_Austria', 'col': 'AE', 'criteria': ('Import', 'MAB', 'Austria')},
            {'name': 'TAP_Italy', 'col': 'AF', 'criteria': ('Import', 'TAP', 'Italy')},
            {'name': 'Austria_Production', 'col': 'AG', 'criteria': ('Production', 'Austria', 'Austria')},
            {'name': 'Italy_Production', 'col': 'AH', 'criteria': ('Production', 'Italy', 'Italy')},
            {'name': 'Germany_Production', 'col': 'AI', 'criteria': ('Production', 'Germany', 'Germany')}
        ]
        
    def load_data(self):
        """Load MultiTicker and LiveSheet data from Excel."""
        print("📂 Loading Excel data...")
        
        # Load MultiTicker sheet
        self.multiticker_df = pd.read_excel(self.excel_file, sheet_name='MultiTicker', header=None)
        print(f"  ✓ MultiTicker loaded: {self.multiticker_df.shape}")
        
        # Load LiveSheet for validation
        self.livesheet_df = pd.read_excel(self.excel_file, sheet_name='Daily historic data by category', header=None)
        print(f"  ✓ LiveSheet loaded: {self.livesheet_df.shape}")
        
        return True
    
    def extract_dates(self):
        """Extract date series from MultiTicker."""
        print("📅 Extracting dates...")
        
        # Dates are in column B (index 1) starting from row 26 (index 25)
        date_series = pd.to_datetime(self.multiticker_df.iloc[25:, 1], errors='coerce')
        valid_dates = date_series[date_series.notna()]
        
        print(f"  ✓ Found {len(valid_dates)} valid dates")
        print(f"  ✓ Date range: {valid_dates.min().date()} to {valid_dates.max().date()}")
        
        return valid_dates
    
    def apply_sumifs(self, data_row_idx, criteria1, criteria2, criteria3):
        """
        Apply Excel's SUMIFS logic for a specific data row.
        
        This is the core function that replicates Excel's exact behavior:
        - No scaling factors applied
        - Exact 3-level criteria matching
        - Wildcard support for LNG
        """
        
        # Column range: C to end (columns 2 onwards in 0-indexed)
        start_col = 2
        end_col = self.multiticker_df.shape[1] - 1
        
        # Criteria rows (14, 15, 16 in Excel = 13, 14, 15 in 0-indexed)
        criteria_rows = [13, 14, 15]
        
        route_total = 0.0
        matches_found = 0
        matched_columns = []
        
        for col_idx in range(start_col, end_col + 1):
            # Get header values
            header1 = self.multiticker_df.iloc[criteria_rows[0], col_idx] if col_idx < self.multiticker_df.shape[1] else None
            header2 = self.multiticker_df.iloc[criteria_rows[1], col_idx] if col_idx < self.multiticker_df.shape[1] else None
            header3 = self.multiticker_df.iloc[criteria_rows[2], col_idx] if col_idx < self.multiticker_df.shape[1] else None
            
            # Check criteria matches
            match1 = str(header1).strip() == str(criteria1).strip() if pd.notna(header1) else False
            match2 = str(header2).strip() == str(criteria2).strip() if pd.notna(header2) else False
            
            # Handle wildcard for criteria3 (LNG case)
            if str(criteria3).strip() == '*':
                match3 = True
            else:
                match3 = str(header3).strip() == str(criteria3).strip() if pd.notna(header3) else False
            
            if match1 and match2 and match3:
                # Get data value - NO SCALING APPLIED
                data_value = self.multiticker_df.iloc[data_row_idx, col_idx]
                if pd.notna(data_value):
                    route_total += float(data_value)
                    matches_found += 1
                    matched_columns.append(col_idx)
        
        return route_total, matches_found, matched_columns
    
    def replicate_all_supply_routes(self):
        """Replicate all 18 supply routes for entire time series."""
        print("\\n🔄 Replicating all supply routes...")
        
        # Get dates
        dates = self.extract_dates()
        
        # Initialize results DataFrame
        self.supply_results = pd.DataFrame(index=dates)
        
        # Process each supply route
        for route_config in self.supply_routes:
            route_name = route_config['name']
            criteria = route_config['criteria']
            
            print(f"\\n📊 Processing {route_name}:")
            print(f"   Criteria: {criteria[0]} | {criteria[1]} | {criteria[2]}")
            
            route_values = []
            
            # Process each date
            for i, date in enumerate(dates):
                # MultiTicker data starts at row 26 (index 25)
                data_row_idx = 25 + i
                
                # Apply SUMIFS
                value, matches, cols = self.apply_sumifs(data_row_idx, criteria[0], criteria[1], criteria[2])
                route_values.append(value)
                
                # Show progress for first date
                if i == 0:
                    print(f"   First date ({date.date()}): {value:.2f} MCM/d ({matches} columns)")
            
            # Store results
            self.supply_results[route_name] = route_values
            
            # Calculate statistics
            non_zero = (pd.Series(route_values) != 0).sum()
            mean_val = pd.Series(route_values).mean()
            max_val = pd.Series(route_values).max()
            min_val = pd.Series(route_values).min()
            
            print(f"   Statistics: {non_zero} non-zero days, Mean: {mean_val:.2f}, Max: {max_val:.2f}, Min: {min_val:.2f}")
        
        # Calculate Total Supply
        print("\\n📊 Calculating Total Supply...")
        self.supply_results['Total_Supply'] = self.supply_results.sum(axis=1)
        
        print(f"✅ Supply replication complete: {len(self.supply_results.columns)} columns")
        
        return self.supply_results
    
    def validate_against_livesheet(self, test_dates=None):
        """Validate replicated values against LiveSheet."""
        print("\\n🔍 Validating against LiveSheet...")
        
        if test_dates is None:
            # Default test dates
            test_dates = ['2017-01-01', '2016-10-08', '2016-12-31']
        
        validation_results = []
        
        for test_date in test_dates:
            print(f"\\n📅 Testing {test_date}:")
            
            # Find date in our results
            test_date_parsed = pd.to_datetime(test_date)
            if test_date_parsed not in self.supply_results.index:
                print(f"  ❌ Date not found in results")
                continue
            
            # Find corresponding row in LiveSheet
            # Map date to LiveSheet row (approximate)
            # 2017-01-01 is at row 105 in LiveSheet
            if test_date == '2017-01-01':
                livesheet_row = 104  # 0-indexed
            elif test_date == '2016-10-08':
                livesheet_row = 19  # 0-indexed
            else:
                print(f"  ⚠️ LiveSheet row mapping not defined for {test_date}")
                continue
            
            # Compare values
            print(f"  {'Route':<25} {'LiveSheet':>10} {'MyCalc':>10} {'Diff':>8} {'Status'}")
            print("  " + "-" * 60)
            
            total_diff = 0
            for i, route in enumerate(self.supply_routes):
                route_name = route['name']
                
                # Get LiveSheet value (columns R-AI = 17-34 in 0-indexed)
                livesheet_col = 17 + i
                livesheet_val = float(self.livesheet_df.iloc[livesheet_row, livesheet_col]) if pd.notna(self.livesheet_df.iloc[livesheet_row, livesheet_col]) else 0.0
                
                # Get my calculated value
                my_val = self.supply_results.loc[test_date_parsed, route_name]
                
                diff = my_val - livesheet_val
                total_diff += abs(diff)
                status = "✅" if abs(diff) < 0.5 else "❌"
                
                print(f"  {route_name:<25} {livesheet_val:>10.2f} {my_val:>10.2f} {diff:>8.2f} {status}")
            
            # Total supply (column AJ = 35 in 0-indexed)
            livesheet_total = float(self.livesheet_df.iloc[livesheet_row, 35]) if pd.notna(self.livesheet_df.iloc[livesheet_row, 35]) else 0.0
            my_total = self.supply_results.loc[test_date_parsed, 'Total_Supply']
            
            print("  " + "-" * 60)
            print(f"  {'TOTAL':<25} {livesheet_total:>10.2f} {my_total:>10.2f} {my_total-livesheet_total:>8.2f}")
            
            accuracy = (1 - abs(my_total - livesheet_total) / livesheet_total) * 100 if livesheet_total != 0 else 0
            print(f"  Accuracy: {accuracy:.1f}%")
            
            validation_results.append({
                'date': test_date,
                'livesheet_total': livesheet_total,
                'my_total': my_total,
                'accuracy': accuracy
            })
        
        return validation_results
    
    def save_results(self, output_file='livesheet_supply_replicated.csv'):
        """Save replicated supply data to CSV."""
        print(f"\\n💾 Saving results to {output_file}...")
        
        self.supply_results.to_csv(output_file)
        print(f"  ✓ Saved {len(self.supply_results)} rows × {len(self.supply_results.columns)} columns")
        
        # Also save a summary
        summary_file = output_file.replace('.csv', '_summary.csv')
        summary = pd.DataFrame({
            'Route': self.supply_results.columns,
            'Non_Zero_Days': [(self.supply_results[col] != 0).sum() for col in self.supply_results.columns],
            'Mean': [self.supply_results[col].mean() for col in self.supply_results.columns],
            'Max': [self.supply_results[col].max() for col in self.supply_results.columns],
            'Min': [self.supply_results[col].min() for col in self.supply_results.columns]
        })
        summary.to_csv(summary_file, index=False)
        print(f"  ✓ Summary saved to {summary_file}")
        
        return True
    
    def run_complete_replication(self):
        """Run the complete supply replication pipeline."""
        print("=" * 80)
        print("LIVESHEET SUPPLY REPLICATION")
        print("=" * 80)
        print(f"Excel file: {self.excel_file}")
        print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        try:
            # Step 1: Load data
            self.load_data()
            
            # Step 2: Replicate all supply routes
            supply_df = self.replicate_all_supply_routes()
            
            # Step 3: Validate against LiveSheet
            validation = self.validate_against_livesheet()
            
            # Step 4: Save results
            self.save_results()
            
            # Final summary
            print("\\n" + "=" * 80)
            print("REPLICATION COMPLETE")
            print("=" * 80)
            
            for val in validation:
                print(f"{val['date']}: Accuracy = {val['accuracy']:.1f}%")
            
            print("\\n✅ Supply columns successfully replicated!")
            
            return supply_df
            
        except Exception as e:
            print(f"\\n❌ Error: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

def main():
    """Execute the LiveSheet supply replication."""
    
    excel_file = "2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx"
    
    if not Path(excel_file).exists():
        print(f"❌ Error: {excel_file} not found")
        return
    
    replicator = LiveSheetSupplyReplicator(excel_file)
    supply_data = replicator.run_complete_replication()
    
    if supply_data is not None:
        print("\\n🎯 SUCCESS: LiveSheet supply columns replicated!")
        print(f"Output files:")
        print(f"  • livesheet_supply_replicated.csv")
        print(f"  • livesheet_supply_replicated_summary.csv")
        
        # Show sample of recent data
        print("\\nRecent data sample (last 5 days):")
        print(supply_data.tail())

if __name__ == "__main__":
    main()