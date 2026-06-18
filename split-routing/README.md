# SplitRouting — Pirate Route-chan ⚓

Приложение для раздельной маршрутизации трафика на Windows и Android.  
**Принцип:** российские IP идут напрямую, остальной трафик — через VPN.  
Исходники: `C:\Users\Admin\Documents\SplitRouting\`

---

## Файлы

| Файл | Описание |
|------|---------|
| `split_routing.py` | Главное приложение — GUI на Tkinter с управлением маршрутами |
| `create_icon.py` | Генерирует иконку приложения (`icon.ico`) |
| `build_android.ps1` | PowerShell скрипт сборки Android APK через Gradle |
| `debug_run.bat` | Запуск приложения в режиме отладки |
| `SplitRouting.spec` | Конфиг PyInstaller для сборки .exe |

---

## `split_routing.py` — как работает

### Интерфейс
Тёмный UI (чёрный + золото). Кнопки управления + ScrolledText с логом.

### Алгоритм
1. **Скачивает список российских CIDR** с двух источников (кешируется на 7 дней):
   - `antifilter.download/list/subnet.lst`
   - `ipdeny.com/ipblocks/data/countries/ru.zone`
2. **Определяет физический шлюз** (через `route print -4`) — IP маршрутизатора без VPN
3. **Добавляет статические маршруты** для всех RU CIDR через физический шлюз
4. **Сохраняет состояние** в `data/state.json` для восстановления после перезагрузки

### Требования
- **Запуск от администратора** — нужен для изменения таблицы маршрутизации Windows
- При запуске без прав автоматически запрашивает повышение (`runas`)
- Python 3.10+, стандартная библиотека (никаких pip-зависимостей)

### Запуск
```bash
# Через батник (сам запросит права)
debug_run.bat

# Или напрямую
python split_routing.py
```

---

## Android приложение

Аналогичный функционал для Android в папке `android/`.

```powershell
# Сборка APK
cd C:\Users\Admin\Documents\SplitRouting
.\build_android.ps1
```

Исходники на Kotlin в `android/app/src/main/java/`.

---

## `create_icon.py` — генерация иконки

Создаёт `icon.ico` для приложения программно (без внешних ресурсов).  
Запускать один раз перед сборкой .exe.

```bash
python create_icon.py
```

---

## Сборка в .exe (PyInstaller)

```bash
pip install pyinstaller
pyinstaller SplitRouting.spec
# Готовый .exe появится в dist/
```

---

## Структура данных

```
data/
├── ru_cidrs.txt    # Кеш CIDR-списка (обновляется раз в 7 дней)
└── state.json      # Текущее состояние маршрутов (включено/выключено)
```
