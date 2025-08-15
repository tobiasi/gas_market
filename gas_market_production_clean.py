# -*- coding: utf-8 -*-
"""
Gas Market Production Processor - Clean production version
Extracts demand and supply data from European Gas Balance LiveSheet
"""

import pandas as pd
import numpy as np
import json

def load_target_data():
    """Load the target LiveSheet data"""
    filename = '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx'
    sheet_name = 'Daily historic data by category'
    
    target_data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
    return target_data

def create_demand_tab(target_data):
    """Create demand tab using correct column mapping"""
    
    # Load the corrected column mapping
    with open('corrected_analysis_results.json', 'r') as f:
        analysis = json.load(f)
    
    column_mapping = analysis['column_mapping']
    data_start_row = 12
    
    # Extract demand data using CORRECT column mapping
    extracted_data = []
    dates = []
    
    for i in range(data_start_row, target_data.shape[0]):
        date_val = target_data.iloc[i, 1]  # Date is in column 1
        
        if pd.notna(date_val):
            # Handle date conversion carefully
            try:
                if hasattr(date_val, 'date'):
                    date_parsed = pd.to_datetime(date_val.date())
                else:
                    date_parsed = pd.to_datetime(str(date_val))
                
                dates.append(date_parsed)
                row_data = {}
                
                # Extract values using correct column mapping
                for key, col_idx in column_mapping.items():
                    val = target_data.iloc[i, col_idx]
                    if pd.notna(val) and isinstance(val, (int, float)):
                        row_data[key] = float(val)
                    else:
                        row_data[key] = np.nan
                
                extracted_data.append(row_data)
                
            except (ValueError, TypeError):
                continue
    
    # Create demand DataFrame
    demand_df = pd.DataFrame(extracted_data, index=dates)
    
    
    return demand_df

def create_supply_tab(target_data):
    """Create supply tab from columns 17-38"""
    
    # Extract supply columns 17-38
    supply_start_col = 17
    supply_end_col = 38
    
    # Build supply column mapping
    supply_columns = {}
    
    for j in range(supply_start_col, supply_end_col + 1):
        row9 = str(target_data.iloc[9, j]) if pd.notna(target_data.iloc[9, j]) else ""
        row10 = str(target_data.iloc[10, j]) if pd.notna(target_data.iloc[10, j]) else ""
        row11 = str(target_data.iloc[11, j]) if pd.notna(target_data.iloc[11, j]) else ""
        
        # Clean up
        row9_clean = row9.replace('nan', '').strip()
        row10_clean = row10.replace('nan', '').strip()
        row11_clean = row11.replace('nan', '').strip()
        
        # Create header name - use the most descriptive available
        # For columns 36 and 38, Row 10 is empty so use Row 9
        if row10_clean:
            header_name = row10_clean
        elif row9_clean:
            header_name = row9_clean
        elif row11_clean:
            header_name = row11_clean
        else:
            header_name = f'Col_{j}'
        
        supply_columns[j] = {
            'header': header_name,
            'category': row9_clean,
            'subcategory': row10_clean,
            'detail': row11_clean
        }
    
    # Extract supply data starting from row 12
    data_start_row = 12
    dates = []
    supply_data = []
    
    for i in range(data_start_row, target_data.shape[0]):
        date_val = target_data.iloc[i, 1]  # Date in column 1
        
        if pd.notna(date_val):
            try:
                if hasattr(date_val, 'date'):
                    date_parsed = pd.to_datetime(date_val.date())
                else:
                    date_parsed = pd.to_datetime(str(date_val))
                
                dates.append(date_parsed)
                
                # Extract supply values for this row
                row_data = {}
                for j in range(supply_start_col, supply_end_col + 1):
                    header = supply_columns[j]['header']
                    val = target_data.iloc[i, j]
                    
                    if pd.notna(val) and isinstance(val, (int, float)):
                        row_data[header] = float(val)
                    else:
                        row_data[header] = np.nan
                
                supply_data.append(row_data)
                
            except (ValueError, TypeError):
                continue
    
    # Create supply DataFrame
    supply_df = pd.DataFrame(supply_data, index=dates)
    
    
    return supply_df, supply_columns

def save_master_output(demand_df, supply_df, supply_columns):
    """Save both demand and supply tabs to master file"""
    
    
    # Prepare demand DataFrame for output
    final_demand_df = demand_df.copy()
    final_demand_df.reset_index(inplace=True)
    final_demand_df.rename(columns={'index': 'Date'}, inplace=True)
    
    # Prepare supply DataFrame for output
    final_supply_df = supply_df.copy()
    final_supply_df.reset_index(inplace=True)
    final_supply_df.rename(columns={'index': 'Date'}, inplace=True)
    
    # Keep supply columns in original order (17-38)
    column_order = ['Date']
    
    # Add supply columns in their original sequential order (17-38)
    for j in range(17, 39):  # Columns 17-38
        if j in supply_columns:
            header = supply_columns[j]['header']
            column_order.append(header)
    
    # Ensure we have all columns (safety check)
    for col in final_supply_df.columns:
        if col not in column_order:
            column_order.append(col)
    
    final_supply_df = final_supply_df[column_order]
    
    # Save master file with both tabs
    master_file = 'European_Gas_Market_Master.xlsx'
    
    with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
        final_demand_df.to_excel(writer, sheet_name='Demand', index=False)
        final_supply_df.to_excel(writer, sheet_name='Supply', index=False)
    
    # Also save individual CSV files
    final_demand_df.to_csv('European_Gas_Demand_Master.csv', index=False)
    final_supply_df.to_csv('European_Gas_Supply_Master.csv', index=False)
    
    return master_file

def main():
    """Main execution function"""
    
    try:
        # Load target data
        target_data = load_target_data()
        
        # Create demand tab
        demand_df = create_demand_tab(target_data)
        
        # Create supply tab
        supply_df, supply_columns = create_supply_tab(target_data)
        
        # Save master output
        master_file = save_master_output(demand_df, supply_df, supply_columns)
        
        print(f"✅ Processing complete: {master_file}")
        print(f"📊 Demand data: {demand_df.shape}")
        print(f"⛽ Supply data: {supply_df.shape}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)