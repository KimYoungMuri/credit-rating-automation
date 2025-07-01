#!/usr/bin/env python3
"""
Model Performance Comparison Test
Tests phi3:mini vs mistral for financial statement classification speed and accuracy.
"""

import requests
import time
import json

def test_model_performance(model_name):
    """Test a specific model's performance on financial classification"""
    print(f"\n🧪 Testing Model: {model_name}")
    print("=" * 50)
    
    test_items = [
        "SUBCHAPTER S INCOME INCOME TAX DEPOSIT",
        "Cash and cash equivalents", 
        "Accounts payable",
        "Long-term debt",
        "DEFERRED COMPENSATION PLAN INVESTMENTS",
        "Subchapter S income tax deposit obligation",
        "REVOLVING LINES OF CREDIT",
        "TOTAL COMMON SHAREHOLDERS' EQUITY"
    ]
    
    total_time = 0
    successful_calls = 0
    results = []
    
    for item in test_items:
        print(f"\nTesting: '{item[:40]}{'...' if len(item) > 40 else ''}'")
        start_time = time.time()
        
        try:
            prompt = f"""You are a financial analyst. Classify this balance sheet item into ONE section:

LINE ITEM: "{item}"

Choose EXACTLY one from:
- current_assets
- noncurrent_assets 
- current_liabilities
- noncurrent_liabilities
- equity

Respond with ONLY the section name."""

            response = requests.post(
                'http://localhost:11434/api/generate',
                json={
                    'model': model_name,
                    'prompt': prompt,
                    'stream': False,
                    'options': {
                        'temperature': 0.1,
                        'num_predict': 20
                    }
                },
                timeout=30
            )
            
            end_time = time.time()
            response_time = end_time - start_time
            total_time += response_time
            
            if response.status_code == 200:
                result = response.json()
                ollama_response = result.get('response', '').strip().lower()
                
                # Check if response is valid
                valid_sections = ['current_assets', 'noncurrent_assets', 'current_liabilities', 'noncurrent_liabilities', 'equity']
                classification = None
                for section in valid_sections:
                    if section in ollama_response:
                        classification = section
                        break
                
                if classification:
                    print(f"  ✅ {response_time:.2f}s → {classification}")
                    successful_calls += 1
                    results.append({
                        'item': item,
                        'time': response_time,
                        'classification': classification,
                        'response': ollama_response
                    })
                else:
                    print(f"  ⚠️  {response_time:.2f}s → Invalid: '{ollama_response}'")
                    results.append({
                        'item': item,
                        'time': response_time,
                        'classification': 'invalid',
                        'response': ollama_response
                    })
            else:
                print(f"  ❌ HTTP {response.status_code}")
                
        except requests.exceptions.Timeout:
            print(f"  ❌ Timeout after 30s")
        except Exception as e:
            print(f"  ❌ Error: {e}")
    
    # Summary
    print(f"\n📊 {model_name} Summary:")
    print(f"  Total time: {total_time:.2f}s")
    print(f"  Average time: {total_time/len(test_items):.2f}s per call")
    print(f"  Success rate: {successful_calls}/{len(test_items)} ({100*successful_calls/len(test_items):.1f}%)")
    
    return results, total_time, successful_calls

def compare_models():
    """Compare multiple models"""
    print("🏆 Model Performance Comparison for Financial Classification")
    print("=" * 70)
    
    models_to_test = ['phi3:mini', 'mistral:latest']
    all_results = {}
    
    for model in models_to_test:
        try:
            results, total_time, success_count = test_model_performance(model)
            all_results[model] = {
                'results': results,
                'total_time': total_time,
                'success_count': success_count,
                'avg_time': total_time / len(results) if results else 0
            }
        except Exception as e:
            print(f"❌ Failed to test {model}: {e}")
    
    # Final comparison
    print(f"\n🏁 FINAL COMPARISON:")
    print("=" * 50)
    for model, data in all_results.items():
        print(f"{model}:")
        print(f"  ⏱️  Average response time: {data['avg_time']:.2f}s")
        print(f"  ✅ Success rate: {data['success_count']}/{8} ({100*data['success_count']/8:.1f}%)")
        print()
    
    # Recommend best model
    if len(all_results) >= 2:
        fastest_model = min(all_results.keys(), key=lambda k: all_results[k]['avg_time'])
        most_accurate = max(all_results.keys(), key=lambda k: all_results[k]['success_count'])
        
        print(f"⚡ Fastest model: {fastest_model} ({all_results[fastest_model]['avg_time']:.2f}s avg)")
        print(f"🎯 Most accurate: {most_accurate} ({all_results[most_accurate]['success_count']}/8 success)")
        
        if fastest_model == most_accurate:
            print(f"🏆 Recommended: {fastest_model} (best overall)")
        else:
            phi3_time = all_results.get('phi3:mini', {}).get('avg_time', 999)
            phi3_success = all_results.get('phi3:mini', {}).get('success_count', 0)
            mistral_time = all_results.get('mistral:latest', {}).get('avg_time', 999)
            mistral_success = all_results.get('mistral:latest', {}).get('success_count', 0)
            
            if phi3_success >= mistral_success and phi3_time < mistral_time:
                print(f"🏆 Recommended: phi3:mini (same accuracy, much faster)")
            elif mistral_success > phi3_success and mistral_time < 15:
                print(f"🏆 Recommended: mistral:latest (more accurate)")
            else:
                print(f"🏆 Recommended: phi3:mini (better speed/accuracy balance)")

def test_specific_problematic_items():
    """Test the specific items that were problematic in our main script"""
    print(f"\n🔍 Testing Specific Problematic Items with phi3:mini")
    print("=" * 60)
    
    problematic_items = [
        "SUBCHAPTER S INCOME INCOME TAX DEPOSIT",
        "Subchapter S income tax deposit—current portion", 
        "Subchapter S income tax deposit obligation"
    ]
    
    for item in problematic_items:
        print(f"\nTesting: '{item}'")
        start_time = time.time()
        
        try:
            prompt = f"""You are a financial analyst. This is from a balance sheet. Classify this item:

"{item}"

Based on accounting principles:
- "deposit" usually means cash set aside (asset)
- "current portion" means due within 1 year (current)
- "obligation" usually means liability

Choose ONE:
- current_assets
- noncurrent_assets 
- current_liabilities
- noncurrent_liabilities
- equity

Answer:"""

            response = requests.post(
                'http://localhost:11434/api/generate',
                json={
                    'model': 'phi3:mini',
                    'prompt': prompt,
                    'stream': False,
                    'options': {
                        'temperature': 0.1,
                        'num_predict': 20
                    }
                },
                timeout=30
            )
            
            end_time = time.time()
            response_time = end_time - start_time
            
            if response.status_code == 200:
                result = response.json()
                ollama_response = result.get('response', '').strip()
                print(f"  ✅ {response_time:.2f}s → '{ollama_response}'")
            else:
                print(f"  ❌ HTTP {response.status_code}")
                
        except Exception as e:
            print(f"  ❌ Error: {e}")

if __name__ == "__main__":
    compare_models()
    test_specific_problematic_items() 