"""Извлекает куки emex.ru из браузера и сохраняет в файл."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import browser_cookie3

cookies_found = []

for browser_name, loader in [
    ('chrome', browser_cookie3.chrome),
    ('edge',   browser_cookie3.edge),
    ('firefox', browser_cookie3.firefox),
]:
    try:
        jar = loader(domain_name='emex.ru')
        items = list(jar)
        if items:
            print(f'{browser_name}: {len(items)} куки для emex.ru')
            for c in items:
                cookies_found.append({
                    'name':   c.name,
                    'value':  c.value,
                    'domain': c.domain,
                    'path':   c.path,
                })
                print(f'  {c.name} = {str(c.value)[:40]}')
        else:
            print(f'{browser_name}: нет куки для emex.ru')
    except Exception as e:
        print(f'{browser_name}: ошибка — {e}')

if cookies_found:
    with open('data/analytics/emex_cookies.json', 'w', encoding='utf-8') as f:
        json.dump(cookies_found, f, ensure_ascii=False, indent=2)
    print(f'\nСохранено {len(cookies_found)} куки в data/analytics/emex_cookies.json')
else:
    print('\nКуки не найдены')
