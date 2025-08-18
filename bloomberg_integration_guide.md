
# Bloomberg Data Integration Instructions

## Files Created:
1. `multiticker_tab.xlsx` - Ready-to-populate MultiTicker structure
2. `complete_ticker_list.csv` - Full ticker list with metadata
3. `bloomberg_integration_guide.md` - This instruction file

## Bloomberg Data Download Steps:

### Option 1: Using Bloomberg Terminal
1. Open Bloomberg Terminal
2. Load ticker list from `complete_ticker_list.csv`
3. Use BDH (Bloomberg Data History) function:
   ```
   =BDH(ticker_list, "PX_LAST", start_date, end_date)
   ```
4. Export data and populate MultiTicker tab

### Option 2: Using Python xbbg Library
```python
import xbbg
import pandas as pd

# Load ticker list
tickers = pd.read_csv('complete_ticker_list.csv')['ticker'].tolist()

# Download data
data = xbbg.blp.bdh(
    tickers=tickers,
    flds='PX_LAST',
    start_date='2013-01-01',
    end_date='2025-08-18'
)

# Format and populate MultiTicker tab
```

### Option 3: Manual Population
1. Open `multiticker_tab.xlsx`
2. MultiTicker sheet has date column (A) and ticker columns (B+)
3. Populate data starting from Row 21
4. Maintain date format: YYYY-MM-DD

## Data Quality Checks:
- Ensure no missing dates
- Handle Bloomberg holidays appropriately
- Apply normalization factors from Metadata sheet
- Validate against existing data patterns

## Pipeline Integration:
- File is compatible with `master_aggregation_script.py`
- Use the created MultiTicker sheet as input
- All existing aggregation logic will work unchanged
