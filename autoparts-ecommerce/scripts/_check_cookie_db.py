"""Читает Chrome cookie DB для emex.ru."""
import sys, io, sqlite3, os, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

PROFILE  = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Google', 'Chrome', 'User Data', 'Default')
COOKIE_F = os.path.join(PROFILE, 'Network', 'Cookies')
TMP_DB   = os.path.join(os.environ['TEMP'], '_emex_ck.db')

print(f'Размер оригинала: {os.path.getsize(COOKIE_F):,} байт')

# Метод 1: os.open с O_RDONLY
try:
    fd = os.open(COOKIE_F, os.O_RDONLY | os.O_BINARY)
    with os.fdopen(fd, 'rb') as f:
        data = f.read()
    with open(TMP_DB, 'wb') as f:
        f.write(data)
    print(f'Метод 1 (os.open): скопировано {len(data):,} байт')
except Exception as e:
    print(f'Метод 1 ошибка: {e}')
    data = None

# Метод 2: shutil.copy с обходом блокировки через subprocess
if not data:
    import subprocess
    r = subprocess.run(['cmd', '/c', f'copy /Y "{COOKIE_F}" "{TMP_DB}"'], capture_output=True, text=True)
    print(f'Метод 2 (cmd copy): {r.returncode}  {r.stdout.strip()[:80]}  {r.stderr.strip()[:80]}')
    if os.path.exists(TMP_DB):
        print(f'  Скопировано: {os.path.getsize(TMP_DB):,} байт')

# Проверяем SQLite
if os.path.exists(TMP_DB) and os.path.getsize(TMP_DB) > 0:
    try:
        conn = sqlite3.connect(f'file:{TMP_DB}?mode=ro&immutable=1', uri=True)
        tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        print(f'Таблицы: {[t[0] for t in tables]}')
        if any(t[0] == 'cookies' for t in tables):
            cnt = conn.execute("SELECT COUNT(*) FROM cookies WHERE host_key LIKE '%emex%'").fetchone()[0]
            print(f'Куки emex.ru: {cnt}')
        conn.close()
    except Exception as e:
        print(f'SQLite ошибка: {e}')
    finally:
        try: os.remove(TMP_DB)
        except: pass
else:
    print('Файл пустой или не существует')
