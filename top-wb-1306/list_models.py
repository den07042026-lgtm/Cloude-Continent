import sys, requests, json
sys.stdout.reconfigure(encoding="utf-8")

key = "AIzaSyDCVRvfRaq3bswrWHTjSl7MWVl0uF4Yysk"
url = f"https://generativelanguage.googleapis.com/v1beta/models?key={key}"
r = requests.get(url, timeout=30)
print(f"HTTP {r.status_code}")
data = r.json()
models = data.get("models", [])
print(f"Найдено моделей: {len(models)}")
for m in models:
    name = m.get("name","")
    methods = m.get("supportedGenerationMethods", [])
    if "generateContent" in methods:
        print(f"  {name}  |  {m.get('displayName','')}  |  limit={m.get('outputTokenLimit','?')}")
