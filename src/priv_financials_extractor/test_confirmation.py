#!/usr/bin/env python3
"""
Test script to demonstrate the new page confirmation system.
"""

import sys
from pathlib import Path
from final_find_fs import FinancialStatementFinder

def test_confirmation():
    """Test the page confirmation system."""
    
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
    print(f"Testing confirmation system on: {pdf_path.name}")
    
    # Step 1: Find financial statement pages
    print("\n" + "="*60)
    print("STEP 1: FINDING FINANCIAL STATEMENT PAGES")
    print("="*60)
    
    finder = FinancialStatementFinder()
    lines, toc_pages, statement_pages = finder.extractContent(str(pdf_path))
    
    # Step 2: Use the confirmation system
    print("\n" + "="*60)
    print("STEP 2: PAGE CONFIRMATION")
    print("="*60)
    
    confirmed_pages = finder.confirm_statement_pages(pdf_path.name)
    
    # Step 3: Show results
    print("\n" + "="*60)
    print("STEP 3: CONFIRMATION RESULTS")
    print("="*60)
    
    if confirmed_pages:
        print("Confirmed pages:")
        for stmt_type, pages in confirmed_pages.items():
            print(f"  {stmt_type.replace('_', ' ').title()}: {pages}")
    else:
        print("No pages were confirmed.")
    
    print("\n" + "="*60)
    print("CONFIRMATION TEST COMPLETE")
    print("="*60)

if __name__ == "__main__":
    test_confirmation() 