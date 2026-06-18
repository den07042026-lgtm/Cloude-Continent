"""
Извлекает куки Chrome для emex.ru через SQLite + AES-256-GCM расшифровку.
"""
import sys, io, os, json, sqlite3, shutil, base64
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try:
    import win32crypt
    from Crypto.Cipher import AES
except ImportError as e:
    print(f'Не хватает библиотеки: {e}')
    sys.exit(1)

PROFILE  = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Google', 'Chrome', 'User Data')
STATE_F  = os.path.join(PROFILE, 'Local State')
COOKIE_F = os.path.join(PROFILE, 'Default', 'Network', 'Cookies')
TMP_DB   = os.path.join(os.environ['TEMP'], '_emex_cookies_tmp.db')

# 1. Ключ шифрования из Local State
with open(STATE_F, 'r', encoding='utf-8') as f:
    state = json.load(f)
enc_key_b64 = state['os_crypt']['encrypted_key']
enc_key = base64.b64decode(enc_key_b64)[5:]           # убираем DPAPI prefix "DPAPI"
key = win32crypt.CryptUnprotectData(enc_key, None, None, None, 0)[1]
print(f'Ключ AES получен: {len(key)} байт')

# 2. Копируем SQLite даже если Chrome держит файл открытым
import ctypes, ctypes.wintypes as wt

GENERIC_READ          = 0x80000000
FILE_SHARE_READ       = 0x00000001
FILE_SHARE_WRITE      = 0x00000002
FILE_SHARE_DELETE     = 0x00000004
OPEN_EXISTING         = 3
FILE_ATTRIBUTE_NORMAL = 0x80

k32  = ctypes.windll.kernel32
src_handle = k32.CreateFileW(
    COOKIE_F,
    GENERIC_READ,
    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
    None,
    OPEN_EXISTING,
    FILE_ATTRIBUTE_NORMAL,
    None,
)
if src_handle == wt.HANDLE(-1).value:
    raise PermissionError(f'Не удалось открыть Cookies файл, код: {k32.GetLastError()}')

try:
    with open(TMP_DB, 'wb') as dst:
        buf  = ctypes.create_string_buffer(1024 * 1024)
        read = wt.DWORD(0)
        while True:
            ok = k32.ReadFile(src_handle, buf, len(buf), ctypes.byref(read), None)
            if not ok or read.value == 0:
                break
            dst.write(buf.raw[:read.value])
finally:
    k32.CloseHandle(src_handle)

# 3. Читаем и расшифровываем куки emex.ru
def decrypt(enc_value: bytes) -> str:
    try:
        if enc_value[:3] == b'v10' or enc_value[:3] == b'v11':
            nonce  = enc_value[3:15]
            cipher = enc_value[15:-16]
            tag    = enc_value[-16:]
            aes    = AES.new(key, AES.MODE_GCM, nonce=nonce)
            return aes.decrypt_and_verify(cipher, tag).decode('utf-8', errors='replace')
        # Старый формат DPAPI
        return win32crypt.CryptUnprotectData(enc_value, None, None, None, 0)[1].decode('utf-8', errors='replace')
    except Exception as e:
        return f'[decrypt error: {e}]'

conn = sqlite3.connect(TMP_DB)
cur  = conn.cursor()
cur.execute("SELECT name, encrypted_value, host_key, path, expires_utc FROM cookies WHERE host_key LIKE '%emex.ru%'")
rows = cur.fetchall()
conn.close()
os.remove(TMP_DB)

print(f'Найдено куки emex.ru: {len(rows)}')
cookies = []
for name, enc_val, host, path, exp in rows:
    val = decrypt(enc_val)
    cookies.append({'name': name, 'value': val, 'domain': host, 'path': path})
    print(f'  {name} = {val[:60]}')

out = 'data/analytics/emex_cookies.json'
with open(out, 'w', encoding='utf-8') as f:
    json.dump(cookies, f, ensure_ascii=False, indent=2)
print(f'\nСохранено в {out}')
