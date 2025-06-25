import requests
import json

def test_ollama_simple():
    """Simple test to verify Ollama is working"""
    ollama_url = "http://localhost:11434"
    model_name = "mistral:latest"
    
    print("🔍 Testing Ollama connection...")
    
    # Test 1: Check if Ollama is running
    try:
        print("1. Checking if Ollama server is running...")
        response = requests.get(f"{ollama_url}/api/tags", timeout=5)
        if response.status_code == 200:
            print("✅ Ollama server is running")
        else:
            print(f"❌ Ollama server returned status {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Cannot connect to Ollama server: {e}")
        print("   Make sure Ollama is running: ollama serve")
        return False
    
    # Test 2: Check if Mistral model is available
    try:
        print("2. Checking if Mistral model is available...")
        response = requests.get(f"{ollama_url}/api/tags", timeout=5)
        models = response.json().get("models", [])
        mistral_available = any(model["name"].startswith(model_name) for model in models)
        
        if mistral_available:
            print("✅ Mistral model is available")
        else:
            print("❌ Mistral model not found")
            print("   Available models:")
            for model in models:
                print(f"     - {model['name']}")
            print("   To install Mistral: ollama pull mistral")
            return False
    except Exception as e:
        print(f"❌ Error checking models: {e}")
        return False
    
    # Test 3: Simple generation test
    try:
        print("3. Testing simple text generation...")
        prompt = "Say 'Hello, Ollama is working!' in one sentence."
        
        payload = {
            "model": model_name,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0.1,
                "max_tokens": 50
            }
        }
        
        response = requests.post(
            f"{ollama_url}/api/generate",
            json=payload,
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            generated_text = result.get("response", "").strip()
            print(f"✅ Generation successful!")
            print(f"   Prompt: {prompt}")
            print(f"   Response: {generated_text}")
            return True
        else:
            print(f"❌ Generation failed with status {response.status_code}")
            print(f"   Response: {response.text}")
            return False
            
    except requests.exceptions.Timeout:
        print("❌ Generation timed out (30 seconds)")
        return False
    except Exception as e:
        print(f"❌ Generation error: {e}")
        return False

if __name__ == "__main__":
    success = test_ollama_simple()
    if success:
        print("\n🎉 Ollama is working correctly!")
    else:
        print("\n💥 Ollama test failed. Please check the setup.") 