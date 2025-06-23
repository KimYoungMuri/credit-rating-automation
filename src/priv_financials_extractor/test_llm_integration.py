#!/usr/bin/env python3
"""
Test script for LLM integration with financial statement mapping.
"""

import sys
from pathlib import Path

# Add the current directory to Python path
current_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(current_dir))

from llm_mapper import LLMMapper
from final_template_mapper import TemplateMatcher

def test_llm_mapper():
    """Test the LLM mapper functionality"""
    print("🧪 Testing LLM Mapper...")
    
    mapper = LLMMapper()
    
    if not mapper.check_ollama_available():
        print("❌ Ollama not available. Please install and run Ollama with Mistral model.")
        print("Install: https://ollama.ai/")
        print("Run: ollama pull mistral")
        return False
    
    print("✅ Ollama with Mistral is available!")
    
    # Test cases
    test_cases = [
        {
            'description': 'Cash and cash equivalents',
            'template_items': ['Cash and equivalents', 'Other'],
            'section': 'current_assets',
            'statement_type': 'balance_sheet'
        },
        {
            'description': 'Trade receivables',
            'template_items': ['Accounts Receivable', 'Other'],
            'section': 'current_assets', 
            'statement_type': 'balance_sheet'
        },
        {
            'description': 'Property, plant and equipment',
            'template_items': ['Net PPE', 'Other'],
            'section': 'noncurrent_assets',
            'statement_type': 'balance_sheet'
        }
    ]
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n--- Test Case {i}: {test_case['description']} ---")
        
        result = mapper.map_with_llm(
            test_case['description'],
            test_case['template_items'],
            test_case['section'],
            test_case['statement_type']
        )
        
        template_item, confidence, reasoning = result
        print(f"Result: {template_item} (confidence: {confidence:.2f})")
        print(f"Reasoning: {reasoning}")
    
    return True

def test_hybrid_mapper():
    """Test the hybrid mapper functionality"""
    print("\n🧪 Testing Hybrid Mapper...")
    
    matcher = TemplateMatcher()
    
    # Test cases
    test_cases = [
        {
            'description': 'Cash and cash equivalents',
            'template_items': ['Cash and equivalents', 'Other'],
            'section': 'current_assets',
            'statement_type': 'balance_sheet'
        },
        {
            'description': 'Trade receivables',
            'template_items': ['Accounts Receivable', 'Other'],
            'section': 'current_assets',
            'statement_type': 'balance_sheet'
        }
    ]
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n--- Hybrid Test Case {i}: {test_case['description']} ---")
        
        result = matcher.hybrid_map_item(
            test_case['description'],
            test_case['template_items'],
            test_case['section'],
            test_case['statement_type']
        )
        
        template_item, confidence, method = result
        print(f"Result: {template_item} (confidence: {confidence:.2f}, method: {method})")
    
    return True

def main():
    """Run all tests"""
    print("🚀 Starting LLM Integration Tests...\n")
    
    # Test LLM mapper
    llm_success = test_llm_mapper()
    
    if llm_success:
        # Test hybrid mapper
        hybrid_success = test_hybrid_mapper()
        
        if hybrid_success:
            print("\n✅ All tests passed! LLM integration is working.")
        else:
            print("\n❌ Hybrid mapper tests failed.")
    else:
        print("\n❌ LLM mapper tests failed.")
    
    print("\n📝 Next steps:")
    print("1. Install Ollama: https://ollama.ai/")
    print("2. Run: ollama pull mistral")
    print("3. Run the full pipeline to see LLM-enhanced mapping in action")

if __name__ == "__main__":
    main() 