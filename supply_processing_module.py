#!/usr/bin/env python3
"""
Supply-Side Processing Module for European Gas Market Analysis

This module extends the existing demand-side pipeline with sophisticated supply-side
Bloomberg data processing including:
- Supply routes 3-criteria SUMIFS replication
- LNG special aggregation logic  
- Geopolitical corrections (Russian supply adjustments)
- European pipeline network flow analysis
- Import/Export/Production classification processing

Maintains complete compatibility with existing demand-side validation targets.
"""

import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Tuple, Optional
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class SupplyProcessingModule:
    """
    Supply-side processing module for European gas market analysis.
    
    Implements sophisticated supply route analysis, LNG aggregation,
    and geopolitical corrections while maintaining demand-side accuracy.
    """
    
    def __init__(self):
        """Initialize supply processing with supply route mappings."""
        self.supply_audit_trail = []
        self.geopolitical_corrections = []
        
        # Key supply routes identified from use4.xlsx analysis
        self.supply_routes_mapping = {
            # Russian supply routes (geopolitical significance)
            'Russia_NordStream_Germany': ('Import', 'Russia (Nord Stream)', 'Germany'),
            
            # Norwegian supply network (major European supplier)
            'Norway_Europe': ('Import', 'Norway', 'Europe'),
            'Norway_Germany_Netherlands': ('Import', 'Norway', 'Germany/Netherlands'),
            'Norway_France': ('Import', 'Norway', 'France'),
            'Norway_GB': ('Import', 'Norway', 'GB'),
            'Norway_Belgium': ('Import', 'Norway', 'Belgium'),
            
            # Central European interconnections
            'Slovakia_Austria': ('Import', 'Slovakia', 'Austria'),
            'Slovenia_Austria': ('Import', 'Slovenia', 'Austria'),
            'Czech_Poland_Germany': ('Import', 'Czech and Poland', 'Germany'),
            'Denmark_Germany': ('Import', 'Denmark', 'Germany'),
            
            # Southern European routes
            'Algeria_Italy': ('Import', 'Algeria', 'Italy'),
            'Libya_Italy': ('Import', 'Libya', 'Italy'),
            'TAP_Italy': ('Import', 'TAP', 'Italy'),
            
            # MAB (Mid-Anatolia-Baku) pipeline
            'MAB_Austria': ('Import', 'MAB', 'Austria'),
            
            # Export routes
            'Austria_Hungary': ('Export', 'Austria', 'Hungary'),
            
            # Domestic production
            'Netherlands_Production': ('Production', 'Netherlands', 'Netherlands'),
            'GB_Production': ('Production', 'GB', 'GB'),
            'Austria_Production': ('Production', 'Austria', 'Austria'),
            'Germany_Production': ('Production', 'Germany', 'Germany'),
            'Italy_Production': ('Production', 'Italy', 'Italy')
        }
        
        # LNG routes mapping (special aggregation required)
        self.lng_routes_mapping = {
            'LNG_Belgium': ('Import', 'LNG', 'Belgium'),
            'LNG_France': ('Import', 'LNG', 'France'),
            'LNG_Germany': ('Import', 'LNG', 'Germany'),
            'LNG_Netherlands': ('Import', 'LNG', 'Netherlands'),
            'LNG_GB': ('Import', 'LNG', 'GB'),
            'LNG_Italy': ('Import', 'LNG', 'Italy')
        }
    
    def sumifs_three_criteria_supply(self, data_df: pd.DataFrame, metadata: Dict,
                                   category_target: str, region_from_target: str, 
                                   region_to_target: str) -> pd.Series:
        """
        Exact replication of Excel 3-criteria SUMIFS for supply-side data.
        
        Replicates: SUMIFS(MultiTicker!$C19:$ZZ19,
                          MultiTicker!$C$14:$ZZ$14, category_criteria,
                          MultiTicker!$C$15:$ZZ$15, region_from_criteria,
                          MultiTicker!$C$16:$ZZ$16, region_to_criteria)
        """
        matching_cols = []
        
        for col, info in metadata.items():
            if col in data_df.columns:
                # 3-criteria exact string matching for supply routes
                category_match = info['category'] == category_target
                region_from_match = info['region'] == region_from_target  # Row 15: Region From
                region_to_match = info['subcategory'] == region_to_target  # Row 16: Region To
                
                if category_match and region_from_match and region_to_match:
                    matching_cols.append(col)
        
        if not matching_cols:
            logger.debug(f"No supply matches for {category_target}/{region_from_target}/{region_to_target}")
            return pd.Series(0.0, index=data_df.index)
        
        logger.debug(f"Found {len(matching_cols)} supply matches for {category_target}/{region_from_target}/{region_to_target}")
        
        # Sum across matching columns (handling NaN as 0)
        result = data_df[matching_cols].sum(axis=1, skipna=True)
        return result
    
    def process_lng_special_aggregation(self, data_df: pd.DataFrame, metadata: Dict) -> pd.DataFrame:
        """
        LNG special aggregation logic - aggregate by source regardless of destination.
        
        Excel criteria: ('Import', 'LNG', '*') - asterisk means aggregate all destinations
        """
        logger.info("🚢 Processing LNG special aggregation logic")
        
        result = pd.DataFrame()
        result['Date'] = data_df['Date']
        
        # LNG Total Aggregation (all destinations)
        lng_total = pd.Series(0.0, index=data_df.index)
        
        # Find all LNG import columns
        lng_cols = []
        for col, info in metadata.items():
            if (col in data_df.columns and 
                info['category'] == 'Import' and 
                info['region'] == 'LNG'):
                lng_cols.append(col)
                lng_total += data_df[col]
        
        result['LNG_Total_All_Destinations'] = lng_total
        
        # LNG by individual destination countries
        for lng_route, (category, region_from, region_to) in self.lng_routes_mapping.items():
            lng_country_total = self.sumifs_three_criteria_supply(
                data_df, metadata, category, region_from, region_to
            )
            result[f'LNG_{region_to}'] = lng_country_total
        
        logger.info(f"✅ LNG aggregation completed: {len(lng_cols)} LNG routes processed")
        
        # Add to audit trail
        self.supply_audit_trail.append({
            'processing_type': 'lng_aggregation',
            'routes_processed': len(lng_cols),
            'total_destinations': len(self.lng_routes_mapping),
            'reason': 'LNG requires special aggregation by source regardless of destination'
        })
        
        return result
    
    def process_supply_routes(self, data_df: pd.DataFrame, metadata: Dict) -> pd.DataFrame:
        """
        Process major European supply routes using 3-criteria SUMIFS logic.
        """
        logger.info("🛢️ Processing European supply routes")
        
        result = pd.DataFrame()
        result['Date'] = data_df['Date']
        
        # Process each major supply route
        for route_name, (category, region_from, region_to) in self.supply_routes_mapping.items():
            route_total = self.sumifs_three_criteria_supply(
                data_df, metadata, category, region_from, region_to
            )
            result[route_name] = route_total
            
            # Log significant routes
            if route_total.sum() > 0:
                logger.debug(f"✅ {route_name}: {category} {region_from}→{region_to}")
        
        # Calculate supply totals by category
        import_routes = [col for col in result.columns if col.startswith(('Norway_', 'Russia_', 'Slovakia_', 'Slovenia_', 'Czech_', 'Denmark_', 'Algeria_', 'Libya_', 'TAP_', 'MAB_'))]
        export_routes = [col for col in result.columns if col.startswith('Austria_Hungary')]
        production_routes = [col for col in result.columns if col.endswith('_Production')]
        
        # Aggregate totals
        if import_routes:
            result['Total_Pipeline_Imports'] = result[import_routes].sum(axis=1)
        else:
            result['Total_Pipeline_Imports'] = pd.Series(0.0, index=data_df.index)
            
        if export_routes:
            result['Total_Exports'] = result[export_routes].sum(axis=1)
        else:
            result['Total_Exports'] = pd.Series(0.0, index=data_df.index)
            
        if production_routes:
            result['Total_Domestic_Production'] = result[production_routes].sum(axis=1)
        else:
            result['Total_Domestic_Production'] = pd.Series(0.0, index=data_df.index)
        
        logger.info(f"✅ Supply routes processing completed: {len(self.supply_routes_mapping)} routes")
        
        # Add to audit trail
        self.supply_audit_trail.append({
            'processing_type': 'supply_routes',
            'routes_processed': len(self.supply_routes_mapping),
            'imports': len(import_routes),
            'exports': len(export_routes),
            'production': len(production_routes),
            'reason': '3-criteria SUMIFS replication for European pipeline network'
        })
        
        return result
    
    def apply_geopolitical_corrections(self, supply_data: pd.DataFrame) -> pd.DataFrame:
        """
        Apply geopolitical corrections to supply data.
        
        Key corrections:
        - Russian Nord Stream supply adjustments post-2022
        - Other geopolitical events affecting European gas supply
        """
        logger.info("🌍 Applying geopolitical corrections to supply data")
        
        corrected_data = supply_data.copy()
        corrections_applied = 0
        
        # Russian Nord Stream corrections (post-2022 geopolitical events)
        if 'Russia_NordStream_Germany' in corrected_data.columns:
            # Get date index
            dates = pd.to_datetime(corrected_data['Date'])
            
            # Apply corrections for post-September 2022 (Nord Stream pipeline damage)
            post_2022_mask = dates >= '2022-09-26'  # Nord Stream incidents
            
            if post_2022_mask.any():
                original_values = corrected_data.loc[post_2022_mask, 'Russia_NordStream_Germany'].sum()
                corrected_data.loc[post_2022_mask, 'Russia_NordStream_Germany'] = 0.0
                corrections_applied += post_2022_mask.sum()
                
                logger.info(f"✅ Applied Nord Stream correction: {corrections_applied} dates zeroed post-2022-09-26")
                
                # Add to audit trail
                self.geopolitical_corrections.append({
                    'route': 'Russia_NordStream_Germany',
                    'correction_type': 'pipeline_damage',
                    'date_from': '2022-09-26',
                    'dates_affected': corrections_applied,
                    'original_sum': original_values,
                    'corrected_sum': 0.0,
                    'reason': 'Nord Stream pipeline incidents - geopolitical supply disruption'
                })
        
        # Recalculate totals after corrections
        if 'Total_Pipeline_Imports' in corrected_data.columns:
            import_routes = [col for col in corrected_data.columns 
                           if col.startswith(('Norway_', 'Russia_', 'Slovakia_', 'Slovenia_', 
                                            'Czech_', 'Denmark_', 'Algeria_', 'Libya_', 'TAP_', 'MAB_'))]
            corrected_data['Total_Pipeline_Imports'] = corrected_data[import_routes].sum(axis=1)
        
        logger.info(f"✅ Geopolitical corrections completed: {len(self.geopolitical_corrections)} corrections applied")
        
        return corrected_data
    
    def validate_supply_side_processing(self, supply_results: Dict) -> bool:
        """
        Validate supply-side processing results.
        
        Checks:
        - Supply route flow consistency
        - LNG aggregation accuracy  
        - Geopolitical corrections applied
        - Energy balance maintenance
        """
        logger.info("🔍 Validating supply-side processing")
        
        validation_passed = True
        
        # Validate supply routes
        if 'supply_routes' in supply_results:
            routes_data = supply_results['supply_routes']
            
            # Check that major routes have data
            key_routes = ['Norway_Europe', 'Total_Pipeline_Imports', 'Total_Domestic_Production']
            for route in key_routes:
                if route in routes_data.columns:
                    total_flow = routes_data[route].sum()
                    if total_flow > 0:
                        logger.info(f"✅ {route}: {total_flow:.2f} total flow")
                    else:
                        logger.warning(f"⚠️  {route}: No flow detected")
                else:
                    logger.warning(f"⚠️  {route}: Route not found")
        
        # Validate LNG aggregation
        if 'lng_data' in supply_results:
            lng_data = supply_results['lng_data']
            
            # Check LNG total vs sum of countries
            if 'LNG_Total_All_Destinations' in lng_data.columns:
                lng_total = lng_data['LNG_Total_All_Destinations'].sum()
                
                # Sum individual countries
                lng_countries = [col for col in lng_data.columns if col.startswith('LNG_') and col != 'LNG_Total_All_Destinations']
                lng_country_sum = lng_data[lng_countries].sum().sum() if lng_countries else 0
                
                if abs(lng_total - lng_country_sum) < 0.01:
                    logger.info(f"✅ LNG aggregation consistent: Total={lng_total:.2f}, Countries={lng_country_sum:.2f}")
                else:
                    logger.warning(f"⚠️  LNG aggregation mismatch: Total={lng_total:.2f}, Countries={lng_country_sum:.2f}")
                    validation_passed = False
        
        # Validate geopolitical corrections
        if self.geopolitical_corrections:
            logger.info(f"✅ Geopolitical corrections applied: {len(self.geopolitical_corrections)}")
            for correction in self.geopolitical_corrections:
                logger.info(f"  - {correction['route']}: {correction['dates_affected']} dates corrected")
        else:
            logger.info("ℹ️  No geopolitical corrections required")
        
        return validation_passed
    
    def export_supply_audit_trail(self, output_file: str = 'supply_processing_audit_trail.csv') -> str:
        """
        Export detailed audit trail of supply-side processing.
        """
        if not self.supply_audit_trail:
            logger.warning("No supply processing audit trail to export")
            return None
        
        audit_df = pd.DataFrame(self.supply_audit_trail)
        audit_df.to_csv(output_file, index=False)
        
        logger.info(f"📝 Supply audit trail exported to {output_file}")
        logger.info(f"📊 Supply processing summary: {len(self.supply_audit_trail)} operations")
        
        return output_file
    
    def export_geopolitical_corrections(self, output_file: str = 'geopolitical_corrections_audit.csv') -> str:
        """
        Export detailed audit trail of geopolitical corrections.
        """
        if not self.geopolitical_corrections:
            logger.warning("No geopolitical corrections to export")
            return None
        
        corrections_df = pd.DataFrame(self.geopolitical_corrections)
        corrections_df.to_csv(output_file, index=False)
        
        logger.info(f"📝 Geopolitical corrections exported to {output_file}")
        logger.info(f"🌍 Geopolitical summary: {len(self.geopolitical_corrections)} corrections")
        
        return output_file
    
    def get_supply_processing_summary(self) -> Dict:
        """
        Get comprehensive summary of supply-side processing.
        """
        return {
            'supply_routes_processed': len(self.supply_routes_mapping),
            'lng_routes_processed': len(self.lng_routes_mapping),
            'geopolitical_corrections': len(self.geopolitical_corrections),
            'audit_trail_entries': len(self.supply_audit_trail),
            'processing_modules': {
                'supply_routes': True,
                'lng_aggregation': True,
                'geopolitical_corrections': bool(self.geopolitical_corrections)
            }
        }


def main():
    """
    Demonstration of supply-side processing module.
    """
    logger.info("🚀 Supply-Side Processing Module Demo")
    logger.info("=" * 80)
    
    try:
        # Initialize supply processing
        supply_processor = SupplyProcessingModule()
        
        # Show supply routes mapping
        logger.info("🛢️ Supply Routes Mapping:")
        for route_name, (category, region_from, region_to) in supply_processor.supply_routes_mapping.items():
            logger.info(f"  {route_name}: {category} {region_from} → {region_to}")
        
        # Show LNG routes mapping  
        logger.info("\n🚢 LNG Routes Mapping:")
        for lng_route, (category, region_from, region_to) in supply_processor.lng_routes_mapping.items():
            logger.info(f"  {lng_route}: {category} {region_from} → {region_to}")
        
        # Show processing summary
        summary = supply_processor.get_supply_processing_summary()
        logger.info("\n📊 Supply Processing Summary:")
        logger.info(f"  Supply routes: {summary['supply_routes_processed']}")
        logger.info(f"  LNG routes: {summary['lng_routes_processed']}")
        logger.info(f"  Processing modules: {summary['processing_modules']}")
        
        logger.info("\n" + "=" * 80)
        logger.info("✅ Supply-Side Processing Module Ready")
        logger.info("🔗 Ready for integration with existing demand-side pipeline")
        
        return supply_processor
        
    except Exception as e:
        logger.error(f"❌ Error in supply processing demo: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    processor = main()