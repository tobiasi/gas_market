# -*- coding: utf-8 -*-
"""
Create seasonal plots for European gas market data
- Daily seasonal: 365 days (excluding Feb 29) x years (2017-2030)
- Monthly seasonal: 12 months x years (2017-2030)
- Both for demand, supply, and YOY changes
"""

import pandas as pd
import numpy as np
from datetime import datetime

def load_data():
    """Load existing demand and supply data"""
    
    print("📊 LOADING DATA FOR SEASONAL ANALYSIS")
    print("=" * 60)
    
    master_file = 'European_Gas_Market_Master.xlsx'
    
    # Load daily data
    demand_daily = pd.read_excel(master_file, sheet_name='Demand')
    supply_daily = pd.read_excel(master_file, sheet_name='Supply')
    
    # Load monthly data
    demand_monthly = pd.read_excel(master_file, sheet_name='Demand Monthly')
    supply_monthly = pd.read_excel(master_file, sheet_name='Supply Monthly')
    
    # Convert dates
    demand_daily['Date'] = pd.to_datetime(demand_daily['Date'])
    supply_daily['Date'] = pd.to_datetime(supply_daily['Date'])
    demand_monthly['Date'] = pd.to_datetime(demand_monthly['Date'])
    supply_monthly['Date'] = pd.to_datetime(supply_monthly['Date'])
    
    print(f"✅ Daily demand: {demand_daily.shape}")
    print(f"✅ Daily supply: {supply_daily.shape}")
    print(f"✅ Monthly demand: {demand_monthly.shape}")
    print(f"✅ Monthly supply: {supply_monthly.shape}")
    
    return demand_daily, supply_daily, demand_monthly, supply_monthly

def create_daily_seasonal(df, data_columns, name_prefix):
    """Create daily seasonal plots (365 days x years)"""
    
    print(f"\n📅 CREATING DAILY SEASONAL FOR {name_prefix}")
    print("=" * 60)
    
    seasonal_dfs = {}
    
    for col in data_columns:
        print(f"  Processing {col}...")
        
        # Create empty DataFrame for years 2017-2030
        years = list(range(2017, 2031))
        days = list(range(1, 366))  # 365 days
        seasonal_df = pd.DataFrame(index=days, columns=years)
        seasonal_df.index.name = 'Day_of_Year'
        
        # Fill with data
        for _, row in df.iterrows():
            date = row['Date']
            year = date.year
            
            # Skip Feb 29 (leap year day)
            if date.month == 2 and date.day == 29:
                continue
            
            # Calculate day of year (adjusting for leap years)
            if date.month > 2 and pd.Timestamp(year, 12, 31).dayofyear == 366:
                # After Feb in leap year, subtract 1 from day of year
                day_of_year = date.dayofyear - 1
            else:
                day_of_year = date.dayofyear
            
            if year in years and 1 <= day_of_year <= 365:
                value = row[col]
                if pd.notna(value):
                    seasonal_df.at[day_of_year, year] = round(value, 2)
        
        # Add 2017-2021 statistical columns
        baseline_years = [2017, 2018, 2019, 2020, 2021]
        baseline_cols = [year for year in baseline_years if year in seasonal_df.columns]
        
        if baseline_cols:
            # Convert to numeric first (in case of mixed types)
            for col in baseline_cols:
                seasonal_df[col] = pd.to_numeric(seasonal_df[col], errors='coerce')
            
            # Calculate row-wise statistics for each day
            seasonal_df['Avg_2017-2021'] = seasonal_df[baseline_cols].mean(axis=1, skipna=True).round(2)
            seasonal_df['Min_2017-2021'] = seasonal_df[baseline_cols].min(axis=1, skipna=True).round(2)
            seasonal_df['Max_2017-2021'] = seasonal_df[baseline_cols].max(axis=1, skipna=True).round(2)
            seasonal_df['Range_2017-2021'] = (seasonal_df['Max_2017-2021'] - seasonal_df['Min_2017-2021']).round(2)
        
        seasonal_dfs[col] = seasonal_df
    
    print(f"✅ Created {len(seasonal_dfs)} daily seasonal plots")
    return seasonal_dfs

def create_monthly_seasonal(df, data_columns, name_prefix):
    """Create monthly seasonal plots (12 months x years)"""
    
    print(f"\n📅 CREATING MONTHLY SEASONAL FOR {name_prefix}")
    print("=" * 60)
    
    seasonal_dfs = {}
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    for col in data_columns:
        print(f"  Processing {col}...")
        
        # Create empty DataFrame for years 2017-2030
        years = list(range(2017, 2031))
        seasonal_df = pd.DataFrame(index=month_names, columns=years)
        seasonal_df.index.name = 'Month'
        
        # Fill with data
        for _, row in df.iterrows():
            date = row['Date']
            year = date.year
            month_idx = date.month - 1  # 0-based index
            
            if year in years:
                value = row[col]
                if pd.notna(value):
                    seasonal_df.at[month_names[month_idx], year] = round(value, 2)
        
        # Add 2017-2021 statistical columns
        baseline_years = [2017, 2018, 2019, 2020, 2021]
        baseline_cols = [year for year in baseline_years if year in seasonal_df.columns]
        
        if baseline_cols:
            # Convert to numeric first (in case of mixed types)
            for col in baseline_cols:
                seasonal_df[col] = pd.to_numeric(seasonal_df[col], errors='coerce')
            
            # Calculate row-wise statistics for each month
            seasonal_df['Avg_2017-2021'] = seasonal_df[baseline_cols].mean(axis=1, skipna=True).round(2)
            seasonal_df['Min_2017-2021'] = seasonal_df[baseline_cols].min(axis=1, skipna=True).round(2)
            seasonal_df['Max_2017-2021'] = seasonal_df[baseline_cols].max(axis=1, skipna=True).round(2)
            seasonal_df['Range_2017-2021'] = (seasonal_df['Max_2017-2021'] - seasonal_df['Min_2017-2021']).round(2)
        
        seasonal_dfs[col] = seasonal_df
    
    print(f"✅ Created {len(seasonal_dfs)} monthly seasonal plots")
    return seasonal_dfs

def create_daily_seasonal_yoy_abs(df, data_columns, name_prefix):
    """Create daily seasonal YOY absolute change plots"""
    
    print(f"\n📈 CREATING DAILY SEASONAL YOY ABSOLUTE FOR {name_prefix}")
    print("=" * 60)
    
    seasonal_yoy_dfs = {}
    
    for col in data_columns:
        print(f"  Processing {col} YOY...")
        
        # First create the seasonal data
        years = list(range(2017, 2031))
        days = list(range(1, 366))
        seasonal_df = pd.DataFrame(index=days, columns=years)
        seasonal_df.index.name = 'Day_of_Year'
        
        # Fill with data
        data_by_day_year = {}
        for _, row in df.iterrows():
            date = row['Date']
            year = date.year
            
            # Skip Feb 29
            if date.month == 2 and date.day == 29:
                continue
            
            # Calculate day of year
            if date.month > 2 and pd.Timestamp(year, 12, 31).dayofyear == 366:
                day_of_year = date.dayofyear - 1
            else:
                day_of_year = date.dayofyear
            
            if year in years and 1 <= day_of_year <= 365:
                value = row[col]
                if pd.notna(value):
                    data_by_day_year[(day_of_year, year)] = value
        
        # Calculate YOY changes
        for day in days:
            for year in years[1:]:  # Start from 2018 for YOY
                current_key = (day, year)
                prev_key = (day, year - 1)
                
                if current_key in data_by_day_year and prev_key in data_by_day_year:
                    current_val = data_by_day_year[current_key]
                    prev_val = data_by_day_year[prev_key]
                    # Calculate absolute YOY change
                    yoy_abs = current_val - prev_val
                    seasonal_df.at[day, year] = round(yoy_abs, 2)
        
        # Add 2017-2021 statistical columns for YOY data
        baseline_years = [2018, 2019, 2020, 2021]  # YOY starts from 2018
        baseline_cols = [year for year in baseline_years if year in seasonal_df.columns]
        
        if baseline_cols:
            # Convert to numeric first (in case of mixed types)
            for col in baseline_cols:
                seasonal_df[col] = pd.to_numeric(seasonal_df[col], errors='coerce')
            
            seasonal_df['Avg_2018-2021'] = seasonal_df[baseline_cols].mean(axis=1, skipna=True).round(2)
            seasonal_df['Min_2018-2021'] = seasonal_df[baseline_cols].min(axis=1, skipna=True).round(2)
            seasonal_df['Max_2018-2021'] = seasonal_df[baseline_cols].max(axis=1, skipna=True).round(2)
            seasonal_df['Range_2018-2021'] = (seasonal_df['Max_2018-2021'] - seasonal_df['Min_2018-2021']).round(2)
        
        seasonal_yoy_dfs[col] = seasonal_df
    
    print(f"✅ Created {len(seasonal_yoy_dfs)} daily seasonal YOY absolute plots")
    return seasonal_yoy_dfs

def create_daily_seasonal_yoy_pct(df, data_columns, name_prefix):
    """Create daily seasonal YOY percentage change plots"""
    
    print(f"\n📈 CREATING DAILY SEASONAL YOY PERCENTAGE FOR {name_prefix}")
    print("=" * 60)
    
    seasonal_yoy_dfs = {}
    
    for col in data_columns:
        print(f"  Processing {col} YOY %...")
        
        # First create the seasonal data
        years = list(range(2017, 2031))
        days = list(range(1, 366))
        seasonal_df = pd.DataFrame(index=days, columns=years)
        seasonal_df.index.name = 'Day_of_Year'
        
        # Fill with data
        data_by_day_year = {}
        for _, row in df.iterrows():
            date = row['Date']
            year = date.year
            
            # Skip Feb 29
            if date.month == 2 and date.day == 29:
                continue
            
            # Calculate day of year
            if date.month > 2 and pd.Timestamp(year, 12, 31).dayofyear == 366:
                day_of_year = date.dayofyear - 1
            else:
                day_of_year = date.dayofyear
            
            if year in years and 1 <= day_of_year <= 365:
                value = row[col]
                if pd.notna(value):
                    data_by_day_year[(day_of_year, year)] = value
        
        # Calculate YOY percentage changes
        for day in days:
            for year in years[1:]:  # Start from 2018 for YOY
                current_key = (day, year)
                prev_key = (day, year - 1)
                
                if current_key in data_by_day_year and prev_key in data_by_day_year:
                    current_val = data_by_day_year[current_key]
                    prev_val = data_by_day_year[prev_key]
                    # Handle division by zero
                    if prev_val != 0:
                        yoy_pct = ((current_val / prev_val) - 1) * 100
                        seasonal_df.at[day, year] = round(yoy_pct, 2)
        
        # Add 2017-2021 statistical columns for YOY data
        baseline_years = [2018, 2019, 2020, 2021]  # YOY starts from 2018
        baseline_cols = [year for year in baseline_years if year in seasonal_df.columns]
        
        if baseline_cols:
            # Convert to numeric first (in case of mixed types)
            for col_year in baseline_cols:
                seasonal_df[col_year] = pd.to_numeric(seasonal_df[col_year], errors='coerce')
            
            seasonal_df['Avg_2018-2021'] = seasonal_df[baseline_cols].mean(axis=1, skipna=True).round(2)
            seasonal_df['Min_2018-2021'] = seasonal_df[baseline_cols].min(axis=1, skipna=True).round(2)
            seasonal_df['Max_2018-2021'] = seasonal_df[baseline_cols].max(axis=1, skipna=True).round(2)
            seasonal_df['Range_2018-2021'] = (seasonal_df['Max_2018-2021'] - seasonal_df['Min_2018-2021']).round(2)
        
        seasonal_yoy_dfs[col] = seasonal_df
    
    print(f"✅ Created {len(seasonal_yoy_dfs)} daily seasonal YOY percentage plots")
    return seasonal_yoy_dfs

def create_monthly_seasonal_yoy_abs(df, data_columns, name_prefix):
    """Create monthly seasonal YOY absolute change plots"""
    
    print(f"\n📈 CREATING MONTHLY SEASONAL YOY ABSOLUTE FOR {name_prefix}")
    print("=" * 60)
    
    seasonal_yoy_dfs = {}
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    for col in data_columns:
        print(f"  Processing {col} YOY...")
        
        # Create empty DataFrame
        years = list(range(2017, 2031))
        seasonal_df = pd.DataFrame(index=month_names, columns=years)
        seasonal_df.index.name = 'Month'
        
        # Fill with data
        data_by_month_year = {}
        for _, row in df.iterrows():
            date = row['Date']
            year = date.year
            month_idx = date.month - 1
            
            if year in years:
                value = row[col]
                if pd.notna(value):
                    data_by_month_year[(month_idx, year)] = value
        
        # Calculate YOY changes
        for month_idx, month_name in enumerate(month_names):
            for year in years[1:]:  # Start from 2018 for YOY
                current_key = (month_idx, year)
                prev_key = (month_idx, year - 1)
                
                if current_key in data_by_month_year and prev_key in data_by_month_year:
                    current_val = data_by_month_year[current_key]
                    prev_val = data_by_month_year[prev_key]
                    # Calculate absolute YOY change
                    yoy_abs = current_val - prev_val
                    seasonal_df.at[month_name, year] = round(yoy_abs, 2)
        
        # Add 2017-2021 statistical columns for YOY data
        baseline_years = [2018, 2019, 2020, 2021]  # YOY starts from 2018
        baseline_cols = [year for year in baseline_years if year in seasonal_df.columns]
        
        if baseline_cols:
            # Convert to numeric first (in case of mixed types)
            for col in baseline_cols:
                seasonal_df[col] = pd.to_numeric(seasonal_df[col], errors='coerce')
            
            seasonal_df['Avg_2018-2021'] = seasonal_df[baseline_cols].mean(axis=1, skipna=True).round(2)
            seasonal_df['Min_2018-2021'] = seasonal_df[baseline_cols].min(axis=1, skipna=True).round(2)
            seasonal_df['Max_2018-2021'] = seasonal_df[baseline_cols].max(axis=1, skipna=True).round(2)
            seasonal_df['Range_2018-2021'] = (seasonal_df['Max_2018-2021'] - seasonal_df['Min_2018-2021']).round(2)
        
        seasonal_yoy_dfs[col] = seasonal_df
    
    print(f"✅ Created {len(seasonal_yoy_dfs)} monthly seasonal YOY absolute plots")
    return seasonal_yoy_dfs

def create_monthly_seasonal_yoy_pct(df, data_columns, name_prefix):
    """Create monthly seasonal YOY percentage change plots"""
    
    print(f"\n📈 CREATING MONTHLY SEASONAL YOY PERCENTAGE FOR {name_prefix}")
    print("=" * 60)
    
    seasonal_yoy_dfs = {}
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    for col in data_columns:
        print(f"  Processing {col} YOY %...")
        
        # Create empty DataFrame
        years = list(range(2017, 2031))
        seasonal_df = pd.DataFrame(index=month_names, columns=years)
        seasonal_df.index.name = 'Month'
        
        # Fill with data
        data_by_month_year = {}
        for _, row in df.iterrows():
            date = row['Date']
            year = date.year
            month_idx = date.month - 1
            
            if year in years:
                value = row[col]
                if pd.notna(value):
                    data_by_month_year[(month_idx, year)] = value
        
        # Calculate YOY percentage changes
        for month_idx, month_name in enumerate(month_names):
            for year in years[1:]:  # Start from 2018 for YOY
                current_key = (month_idx, year)
                prev_key = (month_idx, year - 1)
                
                if current_key in data_by_month_year and prev_key in data_by_month_year:
                    current_val = data_by_month_year[current_key]
                    prev_val = data_by_month_year[prev_key]
                    # Handle division by zero
                    if prev_val != 0:
                        yoy_pct = ((current_val / prev_val) - 1) * 100
                        seasonal_df.at[month_name, year] = round(yoy_pct, 2)
        
        # Add 2017-2021 statistical columns for YOY data
        baseline_years = [2018, 2019, 2020, 2021]  # YOY starts from 2018
        baseline_cols = [year for year in baseline_years if year in seasonal_df.columns]
        
        if baseline_cols:
            # Convert to numeric first (in case of mixed types)
            for col_year in baseline_cols:
                seasonal_df[col_year] = pd.to_numeric(seasonal_df[col_year], errors='coerce')
            
            seasonal_df['Avg_2018-2021'] = seasonal_df[baseline_cols].mean(axis=1, skipna=True).round(2)
            seasonal_df['Min_2018-2021'] = seasonal_df[baseline_cols].min(axis=1, skipna=True).round(2)
            seasonal_df['Max_2018-2021'] = seasonal_df[baseline_cols].max(axis=1, skipna=True).round(2)
            seasonal_df['Range_2018-2021'] = (seasonal_df['Max_2018-2021'] - seasonal_df['Min_2018-2021']).round(2)
        
        seasonal_yoy_dfs[col] = seasonal_df
    
    print(f"✅ Created {len(seasonal_yoy_dfs)} monthly seasonal YOY percentage plots")
    return seasonal_yoy_dfs

def save_seasonal_plots(all_seasonal_data):
    """Save all seasonal plots to Excel"""
    
    print(f"\n💾 SAVING SEASONAL PLOTS TO EXCEL")
    print("=" * 60)
    
    output_file = 'European_Gas_Seasonal_Analysis.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        sheet_count = 0
        
        # Save each seasonal plot as a separate sheet
        for category, data_dict in all_seasonal_data.items():
            for metric, df in data_dict.items():
                # Ensure metric is a string
                metric_str = str(metric)
                
                # Create sheet name (Excel limit is 31 characters)
                if 'daily' in category.lower():
                    if 'yoy_abs' in category.lower():
                        sheet_name = f"D_YOY_Abs_{metric_str[:16]}"
                    elif 'yoy_pct' in category.lower():
                        sheet_name = f"D_YOY_Pct_{metric_str[:16]}"
                    elif 'yoy' in category.lower():
                        sheet_name = f"D_YOY_{metric_str[:20]}"
                    else:
                        sheet_name = f"Daily_{metric_str[:24]}"
                elif 'monthly' in category.lower():
                    if 'yoy_abs' in category.lower():
                        sheet_name = f"M_YOY_Abs_{metric_str[:16]}"
                    elif 'yoy_pct' in category.lower():
                        sheet_name = f"M_YOY_Pct_{metric_str[:16]}"
                    elif 'yoy' in category.lower():
                        sheet_name = f"M_YOY_{metric_str[:20]}"
                    else:
                        sheet_name = f"Monthly_{metric_str[:22]}"
                else:
                    sheet_name = metric_str[:31]
                
                # Remove invalid characters for Excel sheet names
                sheet_name = sheet_name.replace('/', '_').replace('(', '').replace(')', '')
                
                df.to_excel(writer, sheet_name=sheet_name)
                sheet_count += 1
                print(f"  ✅ Saved: {sheet_name}")
        
        # Create summary sheet
        summary_data = {
            'Category': [],
            'Sheet_Type': [],
            'Metric': [],
            'Sheet_Name': []
        }
        
        for category, data_dict in all_seasonal_data.items():
            for metric in data_dict.keys():
                metric_str = str(metric)
                summary_data['Category'].append(category.split('_')[0])
                summary_data['Sheet_Type'].append(category)
                summary_data['Metric'].append(metric_str)
                
                # Generate sheet name
                if 'daily' in category.lower():
                    if 'yoy_abs' in category.lower():
                        sheet_name = f"D_YOY_Abs_{metric_str[:16]}"
                    elif 'yoy_pct' in category.lower():
                        sheet_name = f"D_YOY_Pct_{metric_str[:16]}"
                    elif 'yoy' in category.lower():
                        sheet_name = f"D_YOY_{metric_str[:20]}"
                    else:
                        sheet_name = f"Daily_{metric_str[:24]}"
                elif 'monthly' in category.lower():
                    if 'yoy_abs' in category.lower():
                        sheet_name = f"M_YOY_Abs_{metric_str[:16]}"
                    elif 'yoy_pct' in category.lower():
                        sheet_name = f"M_YOY_Pct_{metric_str[:16]}"
                    elif 'yoy' in category.lower():
                        sheet_name = f"M_YOY_{metric_str[:20]}"
                    else:
                        sheet_name = f"Monthly_{metric_str[:22]}"
                else:
                    sheet_name = metric_str[:31]
                
                sheet_name = sheet_name.replace('/', '_').replace('(', '').replace(')', '')
                summary_data['Sheet_Name'].append(sheet_name)
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
    print(f"\n✅ SAVED: {output_file}")
    print(f"📊 Total sheets: {sheet_count + 1} (including summary)")
    print(f"📈 Added baseline columns: Avg_2017-2021, Min_2017-2021, Max_2017-2021, Range_2017-2021")
    print(f"📈 YOY baseline columns: Avg_2018-2021, Min_2018-2021, Max_2018-2021, Range_2018-2021")
    
    return output_file

def main():
    """Main execution function"""
    
    print("🌍 EUROPEAN GAS MARKET SEASONAL ANALYSIS")
    print("=" * 80)
    print("Creating seasonal plots for demand, supply, and YOY changes")
    print("Structure: Days/Months (rows) x Years 2017-2030 (columns)")
    print("=" * 80)
    
    try:
        # Load data
        demand_daily, supply_daily, demand_monthly, supply_monthly = load_data()
        
        # Define key columns for analysis
        demand_columns = ['Italy', 'Netherlands', 'GB', 'Austria', 'Belgium', 
                         'France', 'Germany', 'Total', 'Industrial', 'LDZ', 'Gas-to-Power']
        
        supply_columns = ['Norway', 'Russia (Nord Stream)', 'Russia (Yamal)', 
                         'LNG', 'Algeria', 'Libya', 'Azerbaijan', 'Netherlands', 
                         'GB', 'Germany', 'Total']
        
        # Filter to existing columns
        demand_columns = [col for col in demand_columns if col in demand_daily.columns]
        supply_columns = [col for col in supply_columns if col in supply_daily.columns]
        
        print(f"\n📊 Processing {len(demand_columns)} demand metrics")
        print(f"⛽ Processing {len(supply_columns)} supply metrics")
        
        # Create all seasonal plots
        all_seasonal = {}
        
        # Daily seasonal
        print("\n" + "="*60)
        print("CREATING DAILY SEASONAL PLOTS")
        print("="*60)
        all_seasonal['demand_daily'] = create_daily_seasonal(
            demand_daily, demand_columns, "DEMAND"
        )
        all_seasonal['supply_daily'] = create_daily_seasonal(
            supply_daily, supply_columns, "SUPPLY"
        )
        
        # Monthly seasonal
        print("\n" + "="*60)
        print("CREATING MONTHLY SEASONAL PLOTS")
        print("="*60)
        all_seasonal['demand_monthly'] = create_monthly_seasonal(
            demand_monthly, demand_columns, "DEMAND"
        )
        all_seasonal['supply_monthly'] = create_monthly_seasonal(
            supply_monthly, supply_columns, "SUPPLY"
        )
        
        # Daily YOY seasonal - Absolute
        print("\n" + "="*60)
        print("CREATING DAILY YOY ABSOLUTE SEASONAL PLOTS")
        print("="*60)
        all_seasonal['demand_daily_yoy_abs'] = create_daily_seasonal_yoy_abs(
            demand_daily, demand_columns, "DEMAND"
        )
        all_seasonal['supply_daily_yoy_abs'] = create_daily_seasonal_yoy_abs(
            supply_daily, supply_columns, "SUPPLY"
        )
        
        # Daily YOY seasonal - Percentage
        print("\n" + "="*60)
        print("CREATING DAILY YOY PERCENTAGE SEASONAL PLOTS")
        print("="*60)
        all_seasonal['demand_daily_yoy_pct'] = create_daily_seasonal_yoy_pct(
            demand_daily, demand_columns, "DEMAND"
        )
        all_seasonal['supply_daily_yoy_pct'] = create_daily_seasonal_yoy_pct(
            supply_daily, supply_columns, "SUPPLY"
        )
        
        # Monthly YOY seasonal - Absolute
        print("\n" + "="*60)
        print("CREATING MONTHLY YOY ABSOLUTE SEASONAL PLOTS")
        print("="*60)
        all_seasonal['demand_monthly_yoy_abs'] = create_monthly_seasonal_yoy_abs(
            demand_monthly, demand_columns, "DEMAND"
        )
        all_seasonal['supply_monthly_yoy_abs'] = create_monthly_seasonal_yoy_abs(
            supply_monthly, supply_columns, "SUPPLY"
        )
        
        # Monthly YOY seasonal - Percentage
        print("\n" + "="*60)
        print("CREATING MONTHLY YOY PERCENTAGE SEASONAL PLOTS")
        print("="*60)
        all_seasonal['demand_monthly_yoy_pct'] = create_monthly_seasonal_yoy_pct(
            demand_monthly, demand_columns, "DEMAND"
        )
        all_seasonal['supply_monthly_yoy_pct'] = create_monthly_seasonal_yoy_pct(
            supply_monthly, supply_columns, "SUPPLY"
        )
        
        # Save all plots
        output_file = save_seasonal_plots(all_seasonal)
        
        # Final summary
        print(f"\n🎉 SEASONAL ANALYSIS COMPLETE")
        print("=" * 80)
        print(f"✅ Daily Seasonal: 365 days × years (2017-2030)")
        print(f"✅ Monthly Seasonal: 12 months × years (2017-2030)")
        print(f"✅ YOY Calculations: Both absolute and percentage changes vs previous year")
        print(f"✅ Feb 29 excluded: Consistent 365-day structure")
        print(f"✅ Future dates: Empty cells for dates not yet occurred")
        print(f"✅ Baseline statistics: 2017-2021 avg, min, max, and range for each day/month")
        print(f"✅ YOY baseline statistics: 2018-2021 avg, min, max, and range for YOY data")
        
        print(f"\n📁 OUTPUT FILE:")
        print(f"   🎯 {output_file}")
        print(f"\n🌟 READY FOR SEASONAL PATTERN ANALYSIS!")
        
        return True
        
    except Exception as e:
        print(f"\n💥 ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)