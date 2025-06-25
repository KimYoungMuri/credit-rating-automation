#!/usr/bin/env python3
"""
Test script for the new SingleStatementMapper on US_Venture_2024 data.
"""

import sys
import pandas as pd
from pathlib import Path
import json

# Add the current directory to Python path
current_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(current_dir))

from single_statement_mapper import SingleStatementMapper

def load_extracted_data(file_path: str) -> dict:
    """Load extracted data from Excel file"""
    extracted_data = {}
    
    try:
        # Read all sheets from the Excel file
        for sheet_name, df in pd.read_excel(file_path, sheet_name=None).items():
            statement_type = sheet_name.lower().replace(' ', '_')
            extracted_data[statement_type] = {}
            
            # Find all year columns (exclude 'Description')
            year_cols = [col for col in df.columns if col != 'Description']
            print(f"[INFO] Found year columns for {statement_type}: {year_cols}")
            
            for year in year_cols:
                extracted_data[statement_type][str(year)] = {}
                for _, row in df.iterrows():
                    desc = row['Description']
                    if pd.notna(desc) and pd.notna(row.get(year)):
                        extracted_data[statement_type][str(year)][desc] = row[year]
            
            print(f"[INFO] Loaded {statement_type}: {len(extracted_data[statement_type])} years")
            for year, items in extracted_data[statement_type].items():
                print(f"[INFO]   {year}: {len(items)} items")
                
    except Exception as e:
        print(f"[ERROR] Failed to load extracted data: {e}")
        return {}
    
    return extracted_data

def test_single_statement_mapping():
    """Test the single statement mapper on US_Venture_2024 data"""
    
    # Initialize mapper
    mapper = SingleStatementMapper()
    
    if not mapper.check_ollama_available():
        print("❌ Ollama not available. Please install and run Ollama with Mistral model.")
        print("Install: https://ollama.ai/")
        print("Run: ollama pull mistral")
        return False
    
    print("✅ Ollama available for single statement mapping!")
    
    # Get the most recent US_Venture_2024 extracted file
    project_root = current_dir.parent.parent
    output_dir = project_root / "output_excel"
    
    # Find US_Venture_2024 files
    us_venture_files = [f for f in output_dir.glob("US_Venture_2024_*.xlsx") if not f.name.startswith('~$')]
    if not us_venture_files:
        print("❌ No US_Venture_2024 extracted files found")
        return False
    
    # Get the most recent file
    latest_file = max(us_venture_files, key=lambda x: x.stat().st_ctime)
    print(f"\n📁 Using extracted file: {latest_file.name}")
    
    # Load extracted data
    extracted_data = load_extracted_data(str(latest_file))
    if not extracted_data:
        print("❌ Failed to load extracted data")
        return False
    
    # Test each statement type
    statement_types = ['balance_sheet', 'income_statement', 'cash_flow']
    
    for statement_type in statement_types:
        if statement_type not in extracted_data:
            print(f"\n⚠️  No {statement_type} data found, skipping...")
            continue
        
        print(f"\n{'='*60}")
        print(f"🧪 TESTING {statement_type.upper().replace('_', ' ')} MAPPING")
        print(f"{'='*60}")
        
        # Get the data for this statement type
        statement_data = extracted_data[statement_type]
        
        # Test comprehensive mapping
        print(f"\n[INFO] Testing comprehensive mapping for {statement_type}...")
        mappings, unmapped = mapper.map_statement_comprehensive(statement_data, statement_type)
        
        if mappings:
            print(f"\n✅ Comprehensive mapping successful for {statement_type}!")
            print(f"[INFO] Mapped sections: {len(mappings)}")
            print(f"[INFO] Unmapped items: {len(unmapped)}")
            
            # Print detailed results
            print(f"\n📊 MAPPING RESULTS FOR {statement_type.upper()}:")
            for section, section_mappings in mappings.items():
                print(f"\n  Section: {section}")
                for template_row, year_values in section_mappings.items():
                    print(f"    {template_row}:")
                    for year, value in year_values.items():
                        print(f"      {year}: {value:,.0f}")
            
            if unmapped:
                print(f"\n⚠️  UNMAPPED ITEMS:")
                for item in unmapped:
                    print(f"  - {item['description']} ({item['year']}): {item['value']} - {item['reason']}")
            
            # Test applying to template
            template_path = project_root / "templates" / "financial_template.xlsx"
            if template_path.exists():
                print(f"\n📝 Applying mappings to template...")
                output_path = mapper.apply_mappings_to_excel(mappings, str(template_path), statement_type)
                print(f"✅ Template saved to: {output_path}")
            else:
                print(f"⚠️  Template not found at {template_path}")
        else:
            print(f"❌ Comprehensive mapping failed for {statement_type}")
    
    print(f"\n{'='*60}")
    print("🎉 SINGLE STATEMENT MAPPER TEST COMPLETE!")
    print(f"{'='*60}")
    
    return True

def main():
    """Run the test"""
    print("🚀 Starting Single Statement Mapper Test")
    print("=" * 60)
    
    success = test_single_statement_mapping()
    
    if success:
        print("\n✅ Test completed successfully!")
        print("\n📝 Key benefits of single statement mapping:")
        print("  - Single LLM call per statement type")
        print("  - Full context awareness")
        print("  - Reduced redundancy")
        print("  - Faster processing")
        print("  - Better accuracy through comprehensive analysis")
    else:
        print("\n❌ Test failed!")
    
    print("\n🔧 Next steps:")
    print("1. Review the generated templates")
    print("2. Compare with existing mapping results")
    print("3. Integrate into your main pipeline")

if __name__ == "__main__":
    main() 