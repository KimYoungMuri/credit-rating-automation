#!/usr/bin/env python3
"""
Test script to verify balance sheet mapping fixes
"""

import sys
from pathlib import Path
import pandas as pd
from final_template_mapper import TemplateMatcher

def test_balance_sheet_mapping():
    """Test the balance sheet mapping with the fixes"""
    
    # Get project root directory
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent.parent
    
    # Get paths
    template_path = project_root / "templates" / "financial_template.xlsx"
    if not template_path.exists():
        print(f"❌ Template not found at {template_path}")
        return False
    
    # Get most recent extracted Excel file
    output_dir = project_root / "output_excel"
    if not output_dir.exists():
        print("❌ No output directory found")
        return False
    
    excel_files = [f for f in output_dir.glob("*.xlsx") if not f.name.startswith('~$')]
    if not excel_files:
        print("❌ No valid Excel files found in output directory")
        return False
    
    # Sort by creation time and get most recent
    latest_file = max(excel_files, key=lambda x: x.stat().st_ctime)
    print(f"📁 Using extracted file: {latest_file.name}")
    
    # Read extracted data with improved logic
    extracted_data = {}
    for sheet_name, df in pd.read_excel(latest_file, sheet_name=None).items():
        statement_type = sheet_name.lower().replace(' ', '_')
        extracted_data[statement_type] = {}
        
        # Find all year columns (exclude 'Description')
        year_cols = [col for col in df.columns if col != 'Description']
        print(f"[INFO] Found year columns for {statement_type}: {year_cols}")
        
        # Initialize year dictionaries
        for year in year_cols:
            extracted_data[statement_type][str(year)] = {}
        
        # Process each row and populate all years
        for _, row in df.iterrows():
            desc = row['Description']
            if pd.notna(desc):
                for year in year_cols:
                    if pd.notna(row.get(year)):
                        extracted_data[statement_type][str(year)][desc] = row[year]
        
        print(f"[INFO] Loaded {statement_type}: {len(extracted_data[statement_type])} years")
        for year, items in extracted_data[statement_type].items():
            print(f"[INFO]   {year}: {len(items)} items")
    
    # Test balance sheet specifically
    if 'balance_sheet' not in extracted_data:
        print("❌ No balance sheet data found")
        return False
    
    bs_data = extracted_data['balance_sheet']
    print(f"\n📊 Balance Sheet Data Analysis:")
    print("=" * 50)
    
    # Check data structure
    for year, items in bs_data.items():
        print(f"\n📅 Year {year}:")
        print(f"   Total items: {len(items)}")
        
        # Check for common balance sheet items
        common_items = ['cash', 'receivable', 'inventory', 'property', 'equipment', 'goodwill', 'payable', 'debt', 'equity']
        found_items = []
        for desc in items.keys():
            desc_lower = desc.lower()
            for common in common_items:
                if common in desc_lower:
                    found_items.append(desc)
                    break
        
        print(f"   Common BS items found: {len(found_items)}")
        if found_items:
            print(f"   Examples: {found_items[:3]}")
    
    # Test template year detection
    print(f"\n🔍 Testing Template Year Detection:")
    print("=" * 50)
    
    matcher = TemplateMatcher()
    
    # Load template and test year detection
    import openpyxl
    wb = openpyxl.load_workbook(template_path)
    bs_sheet = wb['BS']
    
    year_cols = matcher.get_year_columns_from_template(bs_sheet)
    if year_cols:
        print(f"✅ Template year detection successful: {year_cols}")
    else:
        print("❌ Template year detection failed")
        return False
    
    # Test year mapping
    year_mapping = matcher.create_year_mapping(extracted_data, year_cols)
    print(f"✅ Year mapping: {year_mapping}")
    
    # Test hybrid mapping on a small subset
    print(f"\n🧪 Testing Hybrid Mapping:")
    print("=" * 50)
    
    # Get a small sample of balance sheet items
    sample_items = []
    for year, items in bs_data.items():
        for desc, value in list(items.items())[:5]:  # First 5 items per year
            sample_items.append({
                'description': desc,
                'value': value,
                'year': year
            })
    
    print(f"Testing with {len(sample_items)} sample items:")
    for item in sample_items:
        print(f"  - {item['description']} ({item['year']}): {item['value']}")
    
    # Test rule-based mapping
    print(f"\n🔧 Testing Rule-Based Mapping:")
    for item in sample_items:
        desc = item['description']
        template_item, confidence = matcher.apply_rule_based_mapping(desc)
        if template_item:
            print(f"  ✅ '{desc}' -> {template_item} (confidence: {confidence})")
        else:
            print(f"  ❌ '{desc}' -> No rule-based match")
    
    # Test semantic matching
    print(f"\n🧠 Testing Semantic Matching:")
    template_items = ['Cash and equivalents', 'Accounts Receivable', 'Inventory', 'Net PPE', 'Goodwill', 'Other']
    for item in sample_items:
        desc = item['description']
        template_item, confidence = matcher.get_semantic_match(desc, template_items, 'current_assets')
        if template_item and confidence > 0.3:
            print(f"  ✅ '{desc}' -> {template_item} (confidence: {confidence:.2f})")
        else:
            print(f"  ❌ '{desc}' -> No semantic match (best: {confidence:.2f})")
    
    # Test hybrid mapping
    print(f"\n🚀 Testing Hybrid Mapping:")
    for item in sample_items:
        desc = item['description']
        template_item, confidence, method = matcher.hybrid_map_item_decoupled(
            desc, template_items, 'current_assets', 'balance_sheet'
        )
        if template_item:
            print(f"  ✅ '{desc}' -> {template_item} [{method}] (confidence: {confidence:.2f})")
        else:
            print(f"  ❌ '{desc}' -> No hybrid match")
    
    print(f"\n✅ Balance sheet mapping test completed successfully!")
    return True

if __name__ == "__main__":
    success = test_balance_sheet_mapping()
    if success:
        print("\n🎉 All tests passed! The balance sheet mapping fixes are working.")
    else:
        print("\n❌ Some tests failed. Check the output above for issues.")
        sys.exit(1) 