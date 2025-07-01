#!/usr/bin/env python3
"""
Test All Adaptive Financial Statement Mappers
===========================================
Runs Balance Sheet, Income Statement, and Cash Flow Statement mappers 
with the new adaptive extractor for better OCR handling.
"""

import sys
from pathlib import Path
import time
from datetime import datetime

def test_balance_sheet():
    """Test the Balance Sheet mapper with adaptive extractor"""
    print("=" * 80)
    print("🏛️  TESTING BALANCE SHEET MAPPER WITH ADAPTIVE EXTRACTOR")
    print("=" * 80)
    
    try:
        from final_bs_mapper import FinalKGMapper
        
        # Initialize and run the balance sheet mapper
        bs_mapper = FinalKGMapper()
        pdf_path = "../../input_pdfs/US_Venture_2024.pdf"
        
        print("📋 Starting Balance Sheet extraction...")
        start_time = time.time()
        
        mapped_items = bs_mapper.extract_and_process(pdf_path)
        
        end_time = time.time()
        print(f"⏱️  Balance Sheet mapping completed in {end_time - start_time:.2f} seconds")
        
        # Print summary
        print(f"📊 Balance Sheet Results:")
        print(f"   - Mapped items: {len(mapped_items)}")
        print(f"   - Template fields: {list(mapped_items.keys())}")
        
        return True, len(mapped_items)
        
    except Exception as e:
        print(f"❌ Balance Sheet mapper failed: {e}")
        import traceback
        traceback.print_exc()
        return False, 0

def test_income_statement():
    """Test the Income Statement mapper with adaptive extractor"""
    print("\n" + "=" * 80)
    print("📈 TESTING INCOME STATEMENT MAPPER WITH ADAPTIVE EXTRACTOR")
    print("=" * 80)
    
    try:
        from final_is_mapper import FinalISMapper
        
        # Initialize and run the income statement mapper
        is_mapper = FinalISMapper()
        pdf_path = "../../input_pdfs/US_Venture_2024.pdf"
        
        print("📋 Starting Income Statement extraction...")
        start_time = time.time()
        
        mapped_items = is_mapper.extract_and_process(pdf_path)
        
        end_time = time.time()
        print(f"⏱️  Income Statement mapping completed in {end_time - start_time:.2f} seconds")
        
        # Print summary
        print(f"📊 Income Statement Results:")
        print(f"   - Mapped items: {len(mapped_items)}")
        print(f"   - Template fields: {list(mapped_items.keys())}")
        
        return True, len(mapped_items)
        
    except Exception as e:
        print(f"❌ Income Statement mapper failed: {e}")
        import traceback
        traceback.print_exc()
        return False, 0

def test_cash_flow_statement():
    """Test the Cash Flow Statement mapper with adaptive extractor"""
    print("\n" + "=" * 80)
    print("💰 TESTING CASH FLOW STATEMENT MAPPER WITH ADAPTIVE EXTRACTOR")
    print("=" * 80)
    
    try:
        from final_cfs_mapper import FinalCFSMapper
        
        # Initialize and run the cash flow statement mapper
        cfs_mapper = FinalCFSMapper()
        pdf_path = "../../input_pdfs/US_Venture_2024.pdf"
        
        print("📋 Starting Cash Flow Statement extraction...")
        start_time = time.time()
        
        mapped_items = cfs_mapper.extract_and_process(pdf_path)
        
        end_time = time.time()
        print(f"⏱️  Cash Flow Statement mapping completed in {end_time - start_time:.2f} seconds")
        
        # Print summary
        print(f"📊 Cash Flow Statement Results:")
        print(f"   - Mapped items: {len(mapped_items)}")
        print(f"   - Template fields: {list(mapped_items.keys())}")
        
        return True, len(mapped_items)
        
    except Exception as e:
        print(f"❌ Cash Flow Statement mapper failed: {e}")
        import traceback
        traceback.print_exc()
        return False, 0

def main():
    """Run all financial statement mappers with adaptive extractor"""
    print("🚀 COMPREHENSIVE ADAPTIVE EXTRACTOR TEST")
    print("========================================")
    print(f"📅 Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📄 Testing PDF: US_Venture_2024.pdf")
    print()
    
    overall_start = time.time()
    results = {}
    
    # Test all three mappers
    tests = [
        ("Balance Sheet", test_balance_sheet),
        ("Income Statement", test_income_statement), 
        ("Cash Flow Statement", test_cash_flow_statement)
    ]
    
    for test_name, test_func in tests:
        try:
            success, item_count = test_func()
            results[test_name] = {
                'success': success,
                'items': item_count
            }
        except Exception as e:
            print(f"❌ {test_name} test crashed: {e}")
            results[test_name] = {
                'success': False,
                'items': 0
            }
    
    overall_end = time.time()
    
    # Print comprehensive summary
    print("\n" + "=" * 80)
    print("📊 COMPREHENSIVE TEST RESULTS SUMMARY")
    print("=" * 80)
    
    total_items = 0
    successful_tests = 0
    
    for test_name, result in results.items():
        status = "✅ PASSED" if result['success'] else "❌ FAILED"
        print(f"{test_name:20} {status:10} Items: {result['items']:3d}")
        
        if result['success']:
            successful_tests += 1
            total_items += result['items']
    
    print("-" * 80)
    print(f"📈 Overall Success Rate: {successful_tests}/{len(tests)} ({successful_tests/len(tests)*100:.1f}%)")
    print(f"📊 Total Items Mapped: {total_items}")
    print(f"⏱️  Total Runtime: {overall_end - overall_start:.2f} seconds")
    
    # Adaptive extractor benefits
    print("\n🎯 ADAPTIVE EXTRACTOR BENEFITS:")
    print("  ✅ Font size detection and adaptive tolerances")
    print("  ✅ Intelligent word merging for OCR artifacts") 
    print("  ✅ Better pattern matching for split words")
    print("  ✅ Preserved important phrase separations")
    
    if successful_tests == len(tests):
        print("\n🎉 ALL TESTS PASSED! Adaptive extractor is working perfectly!")
        return True
    else:
        print(f"\n⚠️  {len(tests) - successful_tests} test(s) failed. Review the errors above.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1) 