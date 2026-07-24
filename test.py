# import google.generativeai as genai
# import os
# from dotenv import load_dotenv

# load_dotenv()

# genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

# models = genai.list_models()

# for m in models:
#     print("Model Name:", m.name)
#     print("Supported Methods:", m.supported_generation_methods)
#     print("-" * 40)
import requests

response = requests.post(
    "http://localhost:11434/api/generate",
    json={
        "model": "mistral",
        "prompt": "Say hello in one sentence.",
        "stream": False
    }
)

print(response.json()["response"])
