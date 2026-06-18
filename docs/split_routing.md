# SplitRouting — Pirate Route-chan

**Путь:** `C:\Users\Admin\Documents\SplitRouting\`  
**Язык:** Python (Tkinter UI) + Kotlin (Android)  
**Статус:** Завершён

---

## Что это

**"Pirate Route-chan"** — приложение для раздельной маршрутизации трафика на Windows.

**Принцип работы:**
- Российские IP-адреса (CIDR-блоки) → идут **напрямую** (без VPN)
- Весь остальной трафик → идёт **через VPN**

Полезно когда VPN замедляет доступ к рунету или блокирует работу с российскими сервисами.

---

## Структура папки

```
SplitRouting/
├── split_routing.py       # Главное приложение (Python + Tkinter)
├── create_icon.py         # Утилита создания иконки приложения
├── build_android.ps1      # PowerShell скрипт сборки Android APK
├── debug_run.bat          # Батник для быстрого запуска в режиме отладки
│
└── android/               # Полный Android проект (Kotlin/Java)
    ├── app/src/main/java/ # Исходники приложения
    ├── build.gradle        # Конфигурация сборки
    └── gradlew.bat         # Gradle wrapper для Windows
```

---

## Как работает `split_routing.py`

### Интерфейс
- Чёрный UI в стиле "пиратского терминала" (цвета: чёрный + золото)
- Управление через кнопки и ScrolledText

### Логика работы
1. **Скачивает список RU CIDR** с двух источников (кешируется 7 дней):
   - `antifilter.download/list/subnet.lst`
   - `ipdeny.com/ipblocks/data/countries/ru.zone`
   
2. **Находит физический шлюз** (через `route print -4`)

3. **Добавляет статические маршруты** для всех RU CIDR через физический шлюз

4. **Сохраняет состояние** в `data/state.json` для восстановления после перезагрузки

### Требования
- **Запуск от администратора** (нужен для изменения таблицы маршрутизации)
- При первом запуске запрашивает повышение прав (`ShellExecuteW` с `runas`)

---

## Как запустить

```bash
cd C:\Users\Admin\Documents\SplitRouting

# Запуск (попросит права администратора)
python split_routing.py

# Или через батник
debug_run.bat
```

---

## Android приложение

Аналогичный функционал реализован как Android-приложение.

```powershell
# Сборка APK
cd C:\Users\Admin\Documents\SplitRouting
.\build_android.ps1
```

Исходники в `android/app/src/main/java/` на Kotlin.

---

## Исходники CIDR

| URL | Описание |
|-----|---------|
| `antifilter.download/list/subnet.lst` | Актуальный список РФ подсетей |
| `ipdeny.com/ipblocks/data/countries/ru.zone` | Дублирующий источник |

Кеш обновляется раз в 7 дней автоматически.
