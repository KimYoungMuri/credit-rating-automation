#!/usr/bin/env python3
"""
Test script to verify text extraction quality from financial statements.
"""

import sys
from pathlib import Path
from final_extractor import TextExtractor
from final_find_fs import FinancialStatementFinder

def test_extraction():
    """Test the text extraction process and show detailed results."""
    
    # Get project root directory
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent.parent
    
    # Use the most recent PDF
    input_pdfs_dir = project_root / "input_pdfs"
    pdf_files = list(input_pdfs_dir.glob("*.pdf"))
    if not pdf_files:
        print("No PDF files found in input_pdfs directory")
        return
    
    # Use the most recent PDF
    pdf_path = max(pdf_files, key=lambda x: x.stat().st_mtime)
    print(f"Testing extraction on: {pdf_path.name}")
    
    # Step 1: Find financial statement pages
    print("\n" + "="*60)
    print("STEP 1: FINDING FINANCIAL STATEMENT PAGES")
    print("="*60)
    
    finder = FinancialStatementFinder()
    lines, toc_pages, statement_pages = finder.extractContent(str(pdf_path))
    
    high_conf_pages = finder.get_statement_pages()
    print("\nFinancial Statement Detection Results:")
    for stmt_type, info in high_conf_pages.items():
        if info['pages']:
            conf_percent = info['confidence'] * 100
            print(f"  {stmt_type.replace('_', ' ').title()}: Page {info['pages']} ({conf_percent:.1f}% confidence)")
        else:
            print(f"  {stmt_type.replace('_', ' ').title()}: Not Found")
    
    # Convert to format expected by extractor
    pages_to_extract = {}
    for stmt_type, info in high_conf_pages.items():
        if info['pages']:
            pages_to_extract[stmt_type] = info['pages']
    
    if not pages_to_extract:
        print("ERROR: No financial statements found!")
        return
    
    # Step 2: Extract text
    print("\n" + "="*60)
    print("STEP 2: EXTRACTING TEXT FROM IDENTIFIED PAGES")
    print("="*60)
    
    extractor = TextExtractor()
    _, extracted_data = extractor.extract_text(str(pdf_path), process_numbers=True, statement_pages=pages_to_extract)
    
    if not extracted_data:
        print("ERROR: No data extracted!")
        return
    
    # Step 3: Analyze extracted data
    print("\n" + "="*60)
    print("STEP 3: ANALYZING EXTRACTED DATA")
    print("="*60)
    
    for stmt_type, year_data in extracted_data.items():
        print(f"\n--- {stmt_type.upper().replace('_', ' ')} ---")
        
        if not year_data:
            print("  No data found for this statement type")
            continue
            
        for year, items in year_data.items():
            print(f"\n  Year {year}:")
            print(f"    Number of line items: {len(items)}")
            
            # Show first 10 items as examples
            print("    Sample items:")
            for i, (desc, value) in enumerate(list(items.items())[:10]):
                print(f"      {i+1:2d}. {desc:<50} | {value}")
            
            if len(items) > 10:
                print(f"      ... and {len(items) - 10} more items")
            
            # Show some statistics
            numeric_values = [v for v in items.values() if v is not None and str(v).replace('-', '').replace('.', '').isdigit()]
            if numeric_values:
                print(f"    Numeric values found: {len(numeric_values)}")
                print(f"    Value range: {min(numeric_values)} to {max(numeric_values)}")
            else:
                print("    No numeric values found")
    
    # Step 4: Check for potential issues
    print("\n" + "="*60)
    print("STEP 4: POTENTIAL ISSUES ANALYSIS")
    print("="*60)
    
    for stmt_type, year_data in extracted_data.items():
        print(f"\n--- {stmt_type.upper().replace('_', ' ')} ---")
        
        for year, items in year_data.items():
            print(f"\n  Year {year}:")
            
            # Check for empty descriptions
            empty_descs = [desc for desc in items.keys() if not desc.strip()]
            if empty_descs:
                print(f"    ⚠️  Found {len(empty_descs)} items with empty descriptions")
            
            # Check for very short descriptions (potential fragments)
            short_descs = [desc for desc in items.keys() if len(desc.strip()) < 3 and desc.strip()]
            if short_descs:
                print(f"    ⚠️  Found {len(short_descs)} items with very short descriptions: {short_descs}")
            
            # Check for missing values
            missing_values = [desc for desc, val in items.items() if val is None or val == '']
            if missing_values:
                print(f"    ⚠️  Found {len(missing_values)} items with missing values")
            
            # Check for non-numeric values where we expect numbers
            non_numeric = [desc for desc, val in items.items() if val is not None and not str(val).replace('-', '').replace('.', '').isdigit()]
            if non_numeric:
                print(f"    ⚠️  Found {len(non_numeric)} items with non-numeric values")
                for desc in non_numeric[:5]:  # Show first 5
                    print(f"      Example: '{desc}' = '{items[desc]}'")
    
    print("\n" + "="*60)
    print("EXTRACTION TEST COMPLETE")
    print("="*60)

if __name__ == "__main__":
    test_extraction() 