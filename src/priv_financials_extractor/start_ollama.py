import subprocess
import time
import requests
import os
from pathlib import Path

def start_ollama():
    """Start Ollama if it's not already running"""
    
    # Common Ollama installation paths on Windows
    ollama_paths = [
        r"C:\Users\{}\AppData\Local\Programs\Ollama\ollama.exe".format(os.getenv('USERNAME')),
        r"C:\Program Files\Ollama\ollama.exe",
        r"C:\Program Files (x86)\Ollama\ollama.exe"
    ]
    
    # Find Ollama executable
    ollama_exe = None
    for path in ollama_paths:
        if os.path.exists(path):
            ollama_exe = path
            print(f"Found Ollama at: {ollama_exe}")
            break
    
    if not ollama_exe:
        print("❌ Ollama not found in common installation paths.")
        print("Please install Ollama from: https://ollama.ai/download")
        return False
    
    # Check if Ollama is already running
    try:
        response = requests.get("http://localhost:11434/api/tags", timeout=5)
        if response.status_code == 200:
            print("✅ Ollama is already running!")
            return True
    except:
        pass
    
    # Start Ollama
    print("🚀 Starting Ollama...")
    try:
        # Start Ollama in the background
        process = subprocess.Popen([ollama_exe, "serve"], 
                                 stdout=subprocess.PIPE, 
                                 stderr=subprocess.PIPE,
                                 creationflags=subprocess.CREATE_NEW_CONSOLE)
        
        # Wait a bit for Ollama to start
        print("⏳ Waiting for Ollama to start...")
        time.sleep(5)
        
        # Check if it's running
        for i in range(10):  # Try for 10 seconds
            try:
                response = requests.get("http://localhost:11434/api/tags", timeout=5)
                if response.status_code == 200:
                    print("✅ Ollama started successfully!")
                    return True
            except:
                time.sleep(1)
        
        print("❌ Ollama failed to start within 10 seconds")
        return False
        
    except Exception as e:
        print(f"❌ Error starting Ollama: {e}")
        return False

def check_mistral_model():
    """Check if Mistral model is available, pull if not"""
    try:
        # Check available models
        response = requests.get("http://localhost:11434/api/tags", timeout=10)
        if response.status_code == 200:
            models = response.json().get("models", [])
            mistral_models = [m for m in models if "mistral" in m["name"].lower()]
            
            if mistral_models:
                print(f"✅ Mistral model found: {mistral_models[0]['name']}")
                return True
            else:
                print("📥 Mistral model not found. Pulling...")
                # Pull Mistral model
                subprocess.run(["ollama", "pull", "mistral"], check=True)
                print("✅ Mistral model pulled successfully!")
                return True
        else:
            print("❌ Could not check for models")
            return False
    except Exception as e:
        print(f"❌ Error checking/pulling Mistral model: {e}")
        return False

def test_ollama():
    """Test Ollama with a simple prompt"""
    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": "mistral",
                "prompt": "Hello! Just say 'Ollama is working!'",
                "stream": False
            },
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ Ollama test successful!")
            print(f"Response: {result.get('response', '').strip()}")
            return True
        else:
            print(f"❌ Ollama test failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Ollama test error: {e}")
        return False

def main():
    print("🔧 Ollama Setup and Test")
    print("=" * 50)
    
    # Step 1: Start Ollama
    if not start_ollama():
        return
    
    # Step 2: Check/Pull Mistral model
    if not check_mistral_model():
        return
    
    # Step 3: Test Ollama
    if not test_ollama():
        return
    
    print("\n🎉 Ollama is ready to use!")
    print("You can now run your financial statement extraction with LLM support.")

if __name__ == "__main__":
    main() 