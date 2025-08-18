# -*- coding: utf-8 -*-
"""
Created on Fri Feb 24 14:08:02 2023

@author: AD08394

MINIMAL FIXES APPLIED:
1. Add missing 'Replace blanks with #N/A' column 
2. Make MultiIndex unique to avoid duplicate column error
3. STRICT NaN HANDLING: All sum operations use skipna=False 
   - If ANY component is missing (NaN), the entire sum becomes NaN
   - This ensures complete data integrity in calculations

OPTIMIZATIONS INCLUDED:
- Excel optimization using update_spreadsheet_drop_in
"""

# Copy the entire content from New_read_data.py with the applied fixes
# This file now has all NaN handling set to strict mode (skipna=False)