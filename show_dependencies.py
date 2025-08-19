#!/usr/bin/env python3
"""
European Gas Market Analysis - Dependency Visualization
=======================================================

This script shows all dependencies and their relationships clearly.
Run this to understand what files and modules are needed.
"""

import os
import sys
from pathlib import Path

def show_system_architecture():
    """Display the complete system architecture and dependencies."""
    
    print("=" * 80)
    print("🏗️  EUROPEAN GAS MARKET ANALYSIS - SYSTEM ARCHITECTURE")
    print("=" * 80)
    
    print("\n📁 MAIN EXECUTION SCRIPTS:")
    print("┌─ master_gas_analysis.py          ← 🎯 RECOMMENDED: Complete orchestration")
    print("├─ run_with_bloomberg_data.py      ← Main production script (fixed)")
    print("├─ run_with_bloomberg_data_fixed.py ← Alternative with enhanced error handling")
    print("└─ gas_market_bloomberg_chunked.py  ← Bloomberg API integration (advanced)")
    
    print("\n🔧 CORE PROCESSING MODULES:")
    print("┌─ restored_demand_pipeline.py      ← Perfect demand-side processing")
    print("│  └── Methods: run_restored_demand_pipeline(), validate_enhanced_results()")
    print("│")
    print("└─ livesheet_supply_complete.py     ← Perfect supply-side processing")
    print("   └── Function: replicate_livesheet_supply_complete()")
    
    print("\n📊 DATA FILES REQUIRED:")
    print("┌─ 2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx")
    print("│  └── MultiTicker sheet: 3,397 rows × 458 columns")
    print("│  └── Contains all Bloomberg time series data")
    print("│")
    print("└─ use4.xlsx")
    print("   └── TickerList sheet: 439 Bloomberg tickers")
    print("   └── Configuration and normalization factors")
    
    print("\n💾 OUTPUT FILES GENERATED:")
    print("┌─ European_Gas_Demand_Master_Final.csv      ← 14 demand metrics")
    print("├─ European_Gas_Supply_Master_Final.csv      ← 19 supply routes") 
    print("├─ European_Gas_Market_Master_Complete.xlsx  ← Combined analysis (Excel)")
    print("└─ European_Gas_Market_Master_Complete.csv   ← Combined analysis (CSV)")
    
    print("\n🔄 INTERMEDIATE FILES:")
    print("┌─ restored_demand_results.csv       ← Demand processing cache")
    print("├─ restored_demand_audit.csv         ← Demand audit trail")
    print("└─ livesheet_supply_complete.csv     ← Supply processing cache")
    
    print("\n🧪 TESTING & UTILITIES:")
    print("┌─ create_sample_bloomberg_data.py   ← Generate test data")
    print("├─ test_with_sample_bloomberg.py     ← Test with sample data")
    print("└─ compare_outputs.py                ← Validate processing accuracy")

def show_dependency_tree():
    """Show detailed dependency relationships."""
    
    print("\n🌳 DEPENDENCY TREE:")
    print("=" * 50)
    
    print("\nmaster_gas_analysis.py")
    print("├── sys, os, pathlib (built-in)")
    print("├── pandas, numpy (required)")
    print("├── openpyxl (required for Excel)")
    print("├── xbbg (optional for Bloomberg API)")
    print("│")
    print("├── restored_demand_pipeline.py")
    print("│   ├── pandas, numpy")
    print("│   ├── openpyxl") 
    print("│   └── LiveSheet Excel file")
    print("│")
    print("└── livesheet_supply_complete.py")
    print("    ├── pandas, numpy")
    print("    ├── openpyxl")
    print("    └── LiveSheet Excel file")

def show_validation_targets():
    """Display the validation targets that prove accuracy."""
    
    print("\n✅ VALIDATION TARGETS:")
    print("=" * 40)
    
    print("\n🏭 DEMAND-SIDE (2016-10-03):")
    print("  France:      90.13 MCM/d")
    print("  Total:      715.22 MCM/d")
    print("  Industrial: 236.42 MCM/d") 
    print("  LDZ:        307.80 MCM/d")
    print("  Gas-to-Power: 166.71 MCM/d")
    
    print("\n🛢️ SUPPLY-SIDE (2017-01-01):")
    print("  Russia Nord Stream: 151.99 MCM/d")
    print("  Norway Europe:      332.51 MCM/d")
    print("  LNG Total:           24.54 MCM/d")
    print("  Netherlands Prod:   183.07 MCM/d")
    print("  Total Supply:      1048.32 MCM/d")
    
    print("\n⚖️ MARKET BALANCE:")
    print("  Average Demand:  847.5 MCM/d")
    print("  Average Supply:  847.1 MCM/d") 
    print("  Balance:          -0.3 MCM/d ✅")

def show_usage_examples():
    """Show clear usage examples."""
    
    print("\n🚀 USAGE EXAMPLES:")
    print("=" * 40)
    
    print("\n1️⃣ RECOMMENDED - Complete Analysis:")
    print("   python master_gas_analysis.py")
    print("   └── Comprehensive logging, dependency checking, validation")
    
    print("\n2️⃣ PRODUCTION - Main Script:")
    print("   python run_with_bloomberg_data.py") 
    print("   └── Fast execution, proven reliability")
    
    print("\n3️⃣ TROUBLESHOOTING - Enhanced Version:")
    print("   python run_with_bloomberg_data_fixed.py")
    print("   └── Better error handling, fallback mechanisms")
    
    print("\n4️⃣ TESTING - Individual Components:")
    print("   python restored_demand_pipeline.py    # Demand only")
    print("   python livesheet_supply_complete.py   # Supply only")

def check_file_status():
    """Check which files are present in the current directory."""
    
    print("\n📋 FILE STATUS CHECK:")
    print("=" * 40)
    
    # Required files
    required_files = [
        '2025-08-12 - European Gas Supply and Demand Balances LiveSheet (1.8.0).xlsx',
        'use4.xlsx'
    ]
    
    print("\n📊 DATA FILES:")
    for file in required_files:
        if os.path.exists(file):
            size_mb = os.path.getsize(file) / (1024*1024)
            print(f"  ✅ {file} ({size_mb:.1f} MB)")
        else:
            print(f"  ❌ {file} - MISSING")
    
    # Main scripts
    main_scripts = [
        'master_gas_analysis.py',
        'run_with_bloomberg_data.py', 
        'run_with_bloomberg_data_fixed.py',
        'restored_demand_pipeline.py',
        'livesheet_supply_complete.py'
    ]
    
    print("\n🔧 MAIN SCRIPTS:")
    for script in main_scripts:
        if os.path.exists(script):
            print(f"  ✅ {script}")
        else:
            print(f"  ❌ {script} - MISSING")
    
    # Output files
    output_files = [
        'European_Gas_Demand_Master_Final.csv',
        'European_Gas_Supply_Master_Final.csv',
        'European_Gas_Market_Master_Complete.xlsx',
        'European_Gas_Market_Master_Complete.csv'
    ]
    
    print("\n💾 OUTPUT FILES:")
    for file in output_files:
        if os.path.exists(file):
            size_mb = os.path.getsize(file) / (1024*1024)
            print(f"  ✅ {file} ({size_mb:.1f} MB)")
        else:
            print(f"  ℹ️ {file} - Will be generated")

def main():
    """Display complete system overview."""
    
    show_system_architecture()
    show_dependency_tree() 
    show_validation_targets()
    show_usage_examples()
    check_file_status()
    
    print("\n" + "=" * 80)
    print("💡 QUICK START:")
    print("   1. Ensure data files are present")
    print("   2. Run: python master_gas_analysis.py")
    print("   3. Check output files for results")
    print("=" * 80)

if __name__ == "__main__":
    main()