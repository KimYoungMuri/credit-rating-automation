import requests

payload = {
    "model": "mistral:latest",
    "prompt": "Say hello in one sentence.",
    "stream": False,
    "options": {
        "temperature": 0.1,
        "top_p": 0.9,
        "max_tokens": 32
    }
}

response = requests.post("http://localhost:11434/api/generate", json=payload, timeout=60)
print(response.status_code)
print(response.json())