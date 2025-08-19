#!/usr/bin/env python3
"""
Simple Bloomberg Analysis Runner
===============================

Bulletproof wrapper for running Bloomberg gas market analysis.
Works from any directory and any Python environment.
"""

import os
import sys

def main():
    """Run Bloomberg analysis with proper path handling."""
    
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Change to the script directory
    original_cwd = os.getcwd()
    os.chdir(script_dir)
    
    # Add to Python path
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)
    
    print("🚀 Bloomberg Gas Market Analysis")
    print("=" * 50)
    print(f"📂 Working directory: {script_dir}")
    print(f"🐍 Python version: {sys.version.split()[0]}")
    print()
    
    try:
        # Import and run the Bloomberg analysis
        from master_bloomberg_analysis import main as bloomberg_main
        
        print("▶️ Starting Bloomberg analysis...")
        print()
        
        results = bloomberg_main()
        
        print()
        if results is not None:
            print("🎉 Analysis completed successfully!")
            print(f"📊 Results shape: {results.shape}")
        else:
            print("⚠️ Analysis completed with warnings - check output above")
            
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("💡 Make sure you're running from the correct directory")
        return 1
    except Exception as e:
        print(f"❌ Error during analysis: {e}")
        return 1
    finally:
        # Restore original working directory
        os.chdir(original_cwd)
    
    return 0

if __name__ == "__main__":
    sys.exit(main())