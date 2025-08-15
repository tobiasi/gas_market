# -*- coding: utf-8 -*-
"""
FIX: Expand category keywords to capture all demand series
"""

# UPDATED KEYWORDS based on the investigation results:

# Current keywords only capture 17/67 series. We need to expand them.

# EXPANDED INDUSTRIAL KEYWORDS:
industrial_keywords = [
    'Industrial', 'industry', 'industrial and power',
    'RLMmT',  # German industrial consumption codes
    'RLMoT',  # German industrial consumption codes  
    'Gas Industrial',  # Belgium industrial
    'Industrial Demand',  # Generic industrial
    'Consumption',  # Generic consumption (often industrial)
]

# EXPANDED LDZ KEYWORDS:
ldz_keywords = [
    'LDZ', 'ldz', 'Low Distribution Zone',
    'Domestic', 'domestic',  # Domestic demand = LDZ
    'Public Distribution',  # France domestic
    'Gas Domestic',  # Belgium domestic
    'Local Distribution',  # Belgium LDZ
    'SLP',  # German standard load profiles (residential/commercial)
    'Residual Load',  # German LDZ calculations
    'PIRR',  # France domestic demand code
]

# EXPANDED GAS-TO-POWER KEYWORDS:
gtp_keywords = [
    'Gas-to-Power', 'gas-to-power', 'Power', 'electricity',
    'Power Gen', 'Power Generation',  # Power generation
    'HMS',  # France power generation code
    'PP',  # Power Plant abbreviation
]

# NEW CATEGORY: OTHER/FLOWS (for cross-border flows, exports, etc.)
other_keywords = [
    'Flow Exit', 'Flow Entry', 'Net Entry', 'Net Exit',
    'Exit Allocation', 'Entry Allocation',
    'Exports', 'Imports',
    'Losses',  # System losses
    'Cross-border', 'Transit',
]

# GEOGRAPHIC AGGREGATION (by country/region name in description):
geographic_keywords = [
    'Austria Gas', 'Austria',  # Austria regional demand
    'Switzerland', 'Swiss',    # Switzerland flows  
    'Luxembourg',              # Luxembourg flows
]

print("🔧 EXPANDED KEYWORDS TO CAPTURE MORE SERIES:")
print(f"Industrial keywords: {industrial_keywords}")
print(f"LDZ keywords: {ldz_keywords}")
print(f"Gas-to-Power keywords: {gtp_keywords}")
print(f"Other/Flows keywords: {other_keywords}")
print(f"Geographic keywords: {geographic_keywords}")

print("\n💡 RECOMMENDATION:")
print("Replace the keyword lists in your gas_market_consolidated_use4.py with these expanded versions")
print("This should capture most of the 54 uncaptured series and reduce the difference significantly")

# You can also create a "catch-all" category for remaining unmatched series:
print("\n🗂️ SUGGESTED CATEGORY STRUCTURE:")
print("1. Industrial (expanded keywords)")
print("2. LDZ/Domestic (expanded keywords)")  
print("3. Gas-to-Power (expanded keywords)")
print("4. Cross-border Flows (new category)")
print("5. Other/Unmatched (catch remaining series)")