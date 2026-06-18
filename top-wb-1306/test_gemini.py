import sys, requests, json, time
sys.stdout.reconfigure(encoding="utf-8")

key = "AIzaSyDCVRvfRaq3bswrWHTjSl7MWVl0uF4Yysk"

models_to_try = [
    "gemini-2.0-flash-lite",
    "gemini-2.5-flash",
    "gemini-2.0-flash-lite-001",
    "gemini-2.5-flash-lite",
]

for model in models_to_try:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
    payload = {"contents": [{"parts": [{"text": "Скажи слово 'тест' и ничего больше"}]}]}
    try:
        r = requests.post(url, json=payload, timeout=30)
        print(f"{model}: HTTP {r.status_code}")
        if r.status_code == 200:
            data = r.json()
            text = data["candidates"][0]["content"]["parts"][0]["text"]
            print(f"  Ответ: {text[:100]}")
            print(f"  *** Эта модель работает! Используем её ***")
            break
        else:
            err = r.json()
            print(f"  Ошибка: {json.dumps(err.get('error', err), ensure_ascii=False)[:250]}")
    except Exception as e:
        print(f"{model}: EXCEPTION {e}")
    time.sleep(2)
