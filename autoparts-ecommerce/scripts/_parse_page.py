import sys, io, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open('data/analytics/emex_AMDFA475_page.html', encoding='utf-8') as f:
    content = f.read()

# Decode escaped JSON
decoded = content.replace('\\"', '"').replace('\\/', '/').replace('\\n', ' ').replace('\\t', ' ')

# Find originals/analogs section
idx = decoded.find('"originals"')
if idx >= 0:
    print('=== originals context ===')
    print(decoded[max(0, idx-100):idx+500])
    print()

# Find makes section
idx2 = decoded.find('"makes"')
if idx2 >= 0:
    print('=== makes context ===')
    print(decoded[max(0, idx2-50):idx2+800])
    print()

# Find the details section
idx3 = decoded.find('"details":{')
if idx3 >= 0:
    print('=== details section ===')
    print(decoded[idx3:idx3+1500])
