import requests

url = "https://store.google.com/us/product/nest_thermostat?hl=en-US"
try:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    response = requests.get(url, headers=headers, timeout=10)
    print(f"Status: {response.status_code}")
    print(f"Content Length: {len(response.text)}")
    print("Preview: " + response.text[:200])
except Exception as e:
    print(f"Error: {e}")
