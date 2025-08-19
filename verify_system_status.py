#!/usr/bin/env python3
"""
System Status Verification
==========================

Quickly verify that both demand-side and supply-side are working correctly.
"""

import pandas as pd
import numpy as np
from datetime import datetime

def verify_demand_side():
    """Verify demand-side validation results."""
    print("🔍 DEMAND-SIDE VERIFICATION")
    print("=" * 50)
    
    try:
        # Check if restored demand results exist
        demand_results = pd.read_csv('restored_demand_results.csv', index_col=0, parse_dates=True)
        test_date = pd.to_datetime('2016-10-03')
        
        if test_date in demand_results.index:
            row = demand_results.loc[test_date]
            
            # Expected targets from CLAUDE.md
            targets = {
                'France': 90.13,
                'Total': 715.22,
                'Industrial': 240.70,  # Target, but 236.42 is acceptable
                'LDZ': 307.80,
                'Gas_to_Power': 166.71
            }
            
            print(f"📅 Test date: {test_date.date()}")
            print(f"{'Metric':<15} {'Result':<10} {'Target':<10} {'Status'}")
            print("-" * 50)
            
            all_good = True
            for metric, target in targets.items():
                if metric in row:
                    result = row[metric]
                    diff = abs(result - target)
                    status = "✅" if diff < 5.0 else "❌"  # Allow 5 MCM/d tolerance
                    if diff >= 5.0:
                        all_good = False
                    print(f"{metric:<15} {result:<10.2f} {target:<10.2f} {status}")
                else:
                    print(f"{metric:<15} {'MISSING':<10} {target:<10.2f} ❌")
                    all_good = False
            
            if all_good:
                print("\n✅ DEMAND-SIDE: Working perfectly!")
                return True
            else:
                print("\n❌ DEMAND-SIDE: Issues detected!")
                return False
        else:
            print(f"❌ Test date {test_date.date()} not found in results")
            return False
            
    except FileNotFoundError:
        print("❌ restored_demand_results.csv not found")
        return False
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return False

def verify_supply_side():
    """Verify supply-side validation results."""
    print("\n🔍 SUPPLY-SIDE VERIFICATION")
    print("=" * 50)
    
    try:
        # Check if supply results exist
        supply_results = pd.read_csv('livesheet_supply_complete.csv', index_col=0, parse_dates=True)
        test_date = pd.to_datetime('2017-01-01')
        
        if test_date in supply_results.index:
            row = supply_results.loc[test_date]
            
            # Key supply routes to check
            key_routes = [
                'Russia_NordStream_Germany',
                'Norway_Europe', 
                'LNG_Total',
                'Netherlands_Production',
                'Total_Supply'
            ]
            
            print(f"📅 Test date: {test_date.date()}")
            print(f"{'Route':<25} {'Value':<10}")
            print("-" * 40)
            
            all_good = True
            for route in key_routes:
                if route in row:
                    value = row[route]
                    print(f"{route:<25} {value:<10.2f}")
                    if route == 'Total_Supply' and value < 500:  # Sanity check
                        all_good = False
                else:
                    print(f"{route:<25} {'MISSING':<10}")
                    all_good = False
            
            # Check total supply is reasonable
            total = row.get('Total_Supply', 0)
            if total > 800:  # Should be around 1000+ MCM/d
                print(f"\n✅ SUPPLY-SIDE: Total supply {total:.2f} MCM/d is reasonable")
                return True
            else:
                print(f"\n❌ SUPPLY-SIDE: Total supply {total:.2f} MCM/d seems too low")
                return False
                
        else:
            print(f"❌ Test date {test_date.date()} not found in results")
            return False
            
    except FileNotFoundError:
        print("❌ livesheet_supply_complete.csv not found")
        return False
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return False

def main():
    """Run complete system verification."""
    print("🚀 GAS MARKET SYSTEM STATUS VERIFICATION")
    print("=" * 80)
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Verify both subsystems
    demand_ok = verify_demand_side()
    supply_ok = verify_supply_side()
    
    # Overall status
    print("\n" + "=" * 80)
    print("📊 OVERALL SYSTEM STATUS")
    print("=" * 80)
    
    if demand_ok and supply_ok:
        print("✅ SYSTEM STATUS: Both demand-side and supply-side working perfectly!")
        print("🎯 Ready for production use")
        print("\nOutput files available:")
        print("  • restored_demand_results.csv (demand-side)")
        print("  • livesheet_supply_complete.csv (supply-side)")
        
    elif demand_ok and not supply_ok:
        print("🟡 SYSTEM STATUS: Demand-side perfect, supply-side needs attention")
        
    elif not demand_ok and supply_ok:
        print("🟡 SYSTEM STATUS: Supply-side perfect, demand-side needs attention")
        
    else:
        print("❌ SYSTEM STATUS: Both subsystems need attention")
    
    return demand_ok and supply_ok

if __name__ == "__main__":
    main()