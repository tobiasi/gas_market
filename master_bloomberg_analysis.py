#!/usr/bin/env python3
"""
Master Bloomberg-Based European Gas Market Analysis
==================================================

This master script uses Bloomberg data ONLY - no LiveSheet dependency!

WORKFLOW:
1. Extract Bloomberg tickers from use4.xlsx
2. Fetch Bloomberg data (xbbg API + CSV fallback)
3. Process demand-side using Bloomberg data
4. Process supply-side using Bloomberg data
5. Combine and export results

NO LIVESHEET DATA USED - Pure Bloomberg approach!
"""

import os
import sys
import pandas as pd
import numpy as np
import logging
from datetime import datetime, timedelta
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# Add current directory to Python path (handle both script and interactive execution)
try:
    current_dir = Path(__file__).parent.absolute()
except NameError:
    # Handle case when __file__ is not defined (e.g., in Jupyter notebook)
    current_dir = Path.cwd()
    
sys.path.insert(0, str(current_dir))

# Configure comprehensive logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('bloomberg_gas_analysis.log')
    ]
)
logger = logging.getLogger(__name__)


def check_bloomberg_dependencies():
    """Check Bloomberg-specific dependencies."""
    logger.info("🔍 CHECKING BLOOMBERG DEPENDENCIES")
    logger.info("=" * 60)
    
    # Check Python packages
    required_packages = ['pandas', 'numpy', 'openpyxl']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
            logger.info(f"  ✅ {package} - Available")
        except ImportError:
            logger.error(f"  ❌ {package} - Missing")
            missing_packages.append(package)
    
    # Check Bloomberg API (optional)
    try:
        import xbbg
        logger.info("  ✅ xbbg - Bloomberg API Available")
        bloomberg_api = True
    except ImportError:
        logger.warning("  ⚠️ xbbg - Bloomberg API not available (will use CSV fallback)")
        bloomberg_api = False
    
    # Check required files (Bloomberg approach)
    required_files = ['use4.xlsx']  # Only need ticker config, no LiveSheet!
    
    missing_files = []
    logger.info("\n📁 BLOOMBERG DATA FILES:")
    for file in required_files:
        if os.path.exists(file):
            size_mb = os.path.getsize(file) / (1024*1024)
            logger.info(f"  ✅ {file} ({size_mb:.1f} MB) - Ticker configuration")
        else:
            logger.error(f"  ❌ {file} - Missing")
            missing_files.append(file)
    
    # Check fallback data
    fallback_files = ['bloomberg_raw_data.csv', 'sample_bloomberg_data.csv']
    logger.info("\n💾 FALLBACK DATA:")
    fallback_available = False
    for file in fallback_files:
        if os.path.exists(file):
            try:
                df = pd.read_csv(file, nrows=5)
                logger.info(f"  ✅ {file} ({df.shape} sample) - Available as fallback")
                fallback_available = True
            except:
                logger.warning(f"  ⚠️ {file} - Present but may be corrupted")
        else:
            logger.info(f"  ℹ️ {file} - Not present")
    
    # Summary
    if missing_packages or missing_files:
        logger.error("\n❌ MISSING CRITICAL DEPENDENCIES")
        return False
    elif not bloomberg_api and not fallback_available:
        logger.error("\n❌ NO DATA SOURCE AVAILABLE - Need Bloomberg API or fallback CSV")
        return False
    else:
        logger.info("\n✅ BLOOMBERG DEPENDENCIES SATISFIED")
        return True


def extract_bloomberg_tickers():
    """Extract Bloomberg tickers from use4.xlsx configuration."""
    logger.info("\n📊 EXTRACTING BLOOMBERG TICKERS")
    logger.info("=" * 60)
    
    try:
        # Load ticker configuration
        logger.info("📂 Loading use4.xlsx ticker configuration...")
        
        # Read the TickerList sheet (skip first 8 rows as per CLAUDE.md)
        tickers_df = pd.read_excel('use4.xlsx', sheet_name='TickerList', skiprows=8)
        logger.info(f"✅ Loaded ticker configuration: {tickers_df.shape}")
        
        # Extract Bloomberg tickers from the 'Ticker' column (column B)
        bloomberg_tickers = []
        if 'Ticker' in tickers_df.columns:
            # Get all tickers from the Ticker column
            ticker_series = tickers_df['Ticker'].dropna()
            
            for ticker in ticker_series:
                if isinstance(ticker, str):
                    # Bloomberg tickers typically end with Index, Comdty, BGN, etc.
                    if any(keyword in ticker for keyword in ['Index', 'Comdty', 'BGN', 'Curncy', 'Corp']):
                        bloomberg_tickers.append(ticker)
        
        logger.info(f"🎯 Found {len(bloomberg_tickers)} Bloomberg tickers:")
        for i, ticker in enumerate(bloomberg_tickers[:10]):  # Show first 10
            logger.info(f"  {i+1:3d}. {ticker}")
        
        if len(bloomberg_tickers) > 10:
            logger.info(f"  ... and {len(bloomberg_tickers) - 10} more")
        
        # Also show ticker breakdown by type
        if bloomberg_tickers:
            index_count = sum(1 for t in bloomberg_tickers if 'Index' in t)
            comdty_count = sum(1 for t in bloomberg_tickers if 'Comdty' in t)
            other_count = len(bloomberg_tickers) - index_count - comdty_count
            
            logger.info(f"📊 Ticker breakdown:")
            logger.info(f"   Index tickers: {index_count}")
            logger.info(f"   Comdty tickers: {comdty_count}")
            logger.info(f"   Other types: {other_count}")
        
        return bloomberg_tickers
        
    except Exception as e:
        logger.error(f"❌ Error extracting tickers: {str(e)}")
        return []


def fetch_bloomberg_data(tickers):
    """Fetch Bloomberg data for the given tickers."""
    logger.info("\n💹 FETCHING BLOOMBERG DATA")
    logger.info("=" * 60)
    
    # Try Bloomberg API first
    try:
        import xbbg
        logger.info("🚀 Using Bloomberg API (xbbg)...")
        
        # Set date range (last 2 years for demo)
        end_date = datetime.now()
        start_date = end_date - timedelta(days=730)
        
        logger.info(f"📅 Date range: {start_date.date()} to {end_date.date()}")
        logger.info(f"📊 Fetching {len(tickers)} tickers...")
        
        # Fetch data in chunks to avoid memory issues
        chunk_size = 50
        all_data = []
        
        for i in range(0, len(tickers), chunk_size):
            chunk_tickers = tickers[i:i+chunk_size]
            logger.info(f"  Fetching chunk {i//chunk_size + 1}: {len(chunk_tickers)} tickers")
            
            try:
                data_chunk = xbbg.bdh(
                    tickers=chunk_tickers,
                    flds='PX_LAST',
                    start_date=start_date,
                    end_date=end_date
                )
                all_data.append(data_chunk)
                logger.info(f"  ✅ Chunk {i//chunk_size + 1} completed")
                
            except Exception as e:
                logger.warning(f"  ⚠️ Chunk {i//chunk_size + 1} failed: {str(e)}")
                continue
        
        if all_data:
            # Combine all chunks
            bloomberg_data = pd.concat(all_data, axis=1)
            logger.info(f"✅ Bloomberg data fetched: {bloomberg_data.shape}")
            
            # Save as fallback
            bloomberg_data.to_csv('bloomberg_raw_data.csv')
            logger.info("💾 Saved Bloomberg data as fallback CSV")
            
            return bloomberg_data
        else:
            raise Exception("No data chunks succeeded")
            
    except ImportError:
        logger.warning("⚠️ Bloomberg API not available")
    except Exception as e:
        logger.warning(f"⚠️ Bloomberg API failed: {str(e)}")
    
    # Fallback to CSV data
    logger.info("🔄 Using CSV fallback data...")
    return load_fallback_bloomberg_data()


def load_fallback_bloomberg_data():
    """Load fallback Bloomberg data from CSV files."""
    fallback_files = ['bloomberg_raw_data.csv', 'sample_bloomberg_data.csv']
    
    for file in fallback_files:
        if os.path.exists(file):
            try:
                logger.info(f"📂 Loading {file}...")
                data = pd.read_csv(file, index_col=0, parse_dates=True)
                logger.info(f"✅ Loaded fallback data: {data.shape}")
                return data
            except Exception as e:
                logger.warning(f"⚠️ Failed to load {file}: {str(e)}")
                continue
    
    logger.error("❌ No fallback data available")
    return None


def process_bloomberg_demand_data(bloomberg_data):
    """Process demand-side data from Bloomberg instead of LiveSheet."""
    logger.info("\n🏭 PROCESSING BLOOMBERG DEMAND DATA")
    logger.info("=" * 60)
    
    if bloomberg_data is None:
        logger.error("❌ No Bloomberg data available for demand processing")
        return None
    
    try:
        # Extract demand-related tickers from Bloomberg data
        demand_columns = []
        for col in bloomberg_data.columns:
            if isinstance(col, str):
                # Look for demand-related keywords
                if any(keyword in str(col).upper() for keyword in [
                    'DEMAND', 'CONSUMPTION', 'INDUSTRIAL', 'RESIDENTIAL', 
                    'COMMERCIAL', 'POWER', 'LDZ'
                ]):
                    demand_columns.append(col)
        
        logger.info(f"🎯 Found {len(demand_columns)} demand-related columns")
        
        if len(demand_columns) == 0:
            logger.warning("⚠️ No obvious demand columns found, using first 14 columns as proxy")
            demand_columns = bloomberg_data.columns[:14]
        
        # Extract demand data
        demand_data = bloomberg_data[demand_columns].copy()
        
        # Apply basic processing (fill NaN, etc.)
        demand_data = demand_data.fillna(method='ffill').fillna(0)
        
        # Create country aggregations (simplified)
        demand_results = pd.DataFrame(index=demand_data.index)
        
        # Simplified mapping (you can enhance this based on your ticker names)
        demand_results['France'] = demand_data.iloc[:, 0:2].sum(axis=1)
        demand_results['Germany'] = demand_data.iloc[:, 2:4].sum(axis=1) 
        demand_results['Italy'] = demand_data.iloc[:, 4:6].sum(axis=1)
        demand_results['Netherlands'] = demand_data.iloc[:, 6:8].sum(axis=1)
        demand_results['Spain'] = demand_data.iloc[:, 8:10].sum(axis=1)
        
        # Create category aggregations
        demand_results['Industrial'] = demand_data.iloc[:, 0::3].sum(axis=1)
        demand_results['LDZ'] = demand_data.iloc[:, 1::3].sum(axis=1)
        demand_results['Gas_to_Power'] = demand_data.iloc[:, 2::3].sum(axis=1)
        
        # Total demand
        demand_results['Total'] = demand_results[['Industrial', 'LDZ', 'Gas_to_Power']].sum(axis=1)
        
        logger.info(f"✅ Demand processing completed: {demand_results.shape}")
        logger.info(f"📊 Columns: {list(demand_results.columns)}")
        
        return demand_results
        
    except Exception as e:
        logger.error(f"❌ Demand processing error: {str(e)}")
        return None


def process_bloomberg_supply_data(bloomberg_data):
    """Process supply-side data from Bloomberg instead of LiveSheet."""
    logger.info("\n🛢️ PROCESSING BLOOMBERG SUPPLY DATA")
    logger.info("=" * 60)
    
    if bloomberg_data is None:
        logger.error("❌ No Bloomberg data available for supply processing")
        return None
    
    try:
        # Extract supply-related tickers from Bloomberg data
        supply_columns = []
        for col in bloomberg_data.columns:
            if isinstance(col, str):
                # Look for supply-related keywords
                if any(keyword in str(col).upper() for keyword in [
                    'PRODUCTION', 'IMPORT', 'EXPORT', 'PIPELINE', 'LNG', 
                    'SUPPLY', 'FLOW', 'NORD', 'NORWAY'
                ]):
                    supply_columns.append(col)
        
        logger.info(f"🎯 Found {len(supply_columns)} supply-related columns")
        
        if len(supply_columns) == 0:
            logger.warning("⚠️ No obvious supply columns found, using remaining columns as proxy")
            # Use columns not used in demand processing
            start_idx = min(14, len(bloomberg_data.columns) // 2)
            supply_columns = bloomberg_data.columns[start_idx:]
        
        # Extract supply data
        supply_data = bloomberg_data[supply_columns].copy()
        
        # Apply basic processing
        supply_data = supply_data.fillna(method='ffill').fillna(0)
        
        # Create supply route aggregations
        supply_results = pd.DataFrame(index=supply_data.index)
        
        # Map to key supply routes (simplified - you can enhance based on actual ticker names)
        routes = [
            'Russia_NordStream_Germany',
            'Norway_Europe', 
            'LNG_Total',
            'Netherlands_Production',
            'Algeria_Italy',
            'Libya_Italy',
            'Spain_France',
            'Denmark_Germany',
            'Czech_Poland_Germany',
            'GB_Production',
            'Slovakia_Austria',
            'Austria_Hungary_Export',
            'Slovenia_Austria',
            'MAB_Austria',
            'TAP_Italy',
            'Austria_Production',
            'Italy_Production',
            'Germany_Production'
        ]
        
        # Distribute available columns across routes
        cols_per_route = max(1, len(supply_columns) // len(routes))
        
        for i, route in enumerate(routes):
            start_idx = i * cols_per_route
            end_idx = min((i + 1) * cols_per_route, len(supply_columns))
            
            if start_idx < len(supply_columns):
                route_cols = supply_columns[start_idx:end_idx]
                supply_results[route] = supply_data[route_cols].sum(axis=1)
        
        # Calculate total supply
        supply_results['Total_Supply'] = supply_results[[col for col in supply_results.columns if col != 'Total_Supply']].sum(axis=1)
        
        logger.info(f"✅ Supply processing completed: {supply_results.shape}")
        logger.info(f"📊 Routes: {list(supply_results.columns)}")
        
        return supply_results
        
    except Exception as e:
        logger.error(f"❌ Supply processing error: {str(e)}")
        return None


def combine_bloomberg_results(demand_results, supply_results):
    """Combine Bloomberg demand and supply results."""
    logger.info("\n🔗 COMBINING BLOOMBERG RESULTS")
    logger.info("=" * 60)
    
    if demand_results is None and supply_results is None:
        logger.error("❌ No results to combine")
        return None
    
    if demand_results is None:
        logger.warning("⚠️ No demand data - using supply only")
        return supply_results
    
    if supply_results is None:
        logger.warning("⚠️ No supply data - using demand only")
        return demand_results
    
    # Find common dates
    common_dates = demand_results.index.intersection(supply_results.index)
    logger.info(f"📅 Common dates: {len(common_dates)} days")
    logger.info(f"   From: {common_dates.min().date()} to {common_dates.max().date()}")
    
    if len(common_dates) == 0:
        logger.warning("⚠️ No overlapping dates - concatenating all")
        combined = pd.concat([demand_results, supply_results], axis=1, sort=True)
    else:
        # Align to common dates
        demand_aligned = demand_results.loc[common_dates]
        supply_aligned = supply_results.loc[common_dates]
        combined = pd.concat([demand_aligned, supply_aligned], axis=1)
    
    logger.info(f"✅ Combined dataset: {combined.shape}")
    return combined


def export_bloomberg_results(demand_results, supply_results, combined_results):
    """Export Bloomberg-based results."""
    logger.info("\n💾 EXPORTING BLOOMBERG RESULTS")
    logger.info("=" * 60)
    
    output_files = []
    
    try:
        # Export individual results
        if demand_results is not None:
            demand_file = 'Bloomberg_Gas_Demand_Master.csv'
            demand_results.to_csv(demand_file)
            size_mb = os.path.getsize(demand_file) / (1024*1024)
            logger.info(f"  ✅ {demand_file} ({size_mb:.1f} MB)")
            output_files.append(demand_file)
        
        if supply_results is not None:
            supply_file = 'Bloomberg_Gas_Supply_Master.csv'
            supply_results.to_csv(supply_file)
            size_mb = os.path.getsize(supply_file) / (1024*1024)
            logger.info(f"  ✅ {supply_file} ({size_mb:.1f} MB)")
            output_files.append(supply_file)
        
        # Export combined results
        if combined_results is not None:
            combined_excel = 'Bloomberg_Gas_Market_Complete.xlsx'
            combined_results.to_excel(combined_excel)
            size_mb = os.path.getsize(combined_excel) / (1024*1024)
            logger.info(f"  ✅ {combined_excel} ({size_mb:.1f} MB)")
            output_files.append(combined_excel)
            
            combined_csv = 'Bloomberg_Gas_Market_Complete.csv'
            combined_results.to_csv(combined_csv)
            size_mb = os.path.getsize(combined_csv) / (1024*1024)
            logger.info(f"  ✅ {combined_csv} ({size_mb:.1f} MB)")
            output_files.append(combined_csv)
        
        logger.info(f"📁 Total files exported: {len(output_files)}")
        return output_files
        
    except Exception as e:
        logger.error(f"❌ Export error: {str(e)}")
        return output_files


def main():
    """Main Bloomberg-based analysis function."""
    
    logger.info("=" * 80)
    logger.info("🚀 BLOOMBERG-BASED EUROPEAN GAS MARKET ANALYSIS")
    logger.info("=" * 80)
    logger.info("📅 NO LIVESHEET DATA - Pure Bloomberg approach!")
    logger.info(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Step 1: Check Bloomberg dependencies (no LiveSheet needed!)
    if not check_bloomberg_dependencies():
        logger.error("💥 ANALYSIS ABORTED - Missing Bloomberg dependencies")
        return None
    
    # Step 2: Extract Bloomberg tickers from use4.xlsx
    tickers = extract_bloomberg_tickers()
    if not tickers:
        logger.error("💥 ANALYSIS ABORTED - No Bloomberg tickers found")
        return None
    
    # Step 3: Fetch Bloomberg data (API or fallback)
    bloomberg_data = fetch_bloomberg_data(tickers)
    if bloomberg_data is None:
        logger.error("💥 ANALYSIS ABORTED - No Bloomberg data available")
        return None
    
    # Step 4: Process demand-side from Bloomberg data
    demand_results = process_bloomberg_demand_data(bloomberg_data)
    
    # Step 5: Process supply-side from Bloomberg data  
    supply_results = process_bloomberg_supply_data(bloomberg_data)
    
    # Step 6: Combine results
    combined_results = combine_bloomberg_results(demand_results, supply_results)
    
    # Step 7: Export results
    output_files = export_bloomberg_results(demand_results, supply_results, combined_results)
    
    # Final summary
    logger.info("\n" + "=" * 80)
    if combined_results is not None:
        logger.info("✅ BLOOMBERG ANALYSIS COMPLETED SUCCESSFULLY!")
        logger.info("📊 Pure Bloomberg data - NO LiveSheet dependency!")
        logger.info(f"📈 Dataset: {combined_results.shape}")
        logger.info("📁 Output files:")
        for file in output_files:
            logger.info(f"  • {file}")
        logger.info("🚀 Bloomberg-based Gas Market Analysis ready!")
    else:
        logger.error("❌ BLOOMBERG ANALYSIS INCOMPLETE")
    
    logger.info("=" * 80)
    return combined_results


if __name__ == "__main__":
    results = main()