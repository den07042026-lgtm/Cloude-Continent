# Как загрузить на GitHub

## 1. Создать репозиторий на GitHub

1. Зайти на github.com → New repository
2. Название: `my-projects` (или любое другое)
3. Private (рекомендую — там могут быть ключи)
4. Без README (мы создадим сами)

## 2. Инициализировать git в этой папке

```bash
cd C:\Users\Admin\my-projects
git init
git add .
git commit -m "Initial: инвентарь всех проектов"
```

## 3. Привязать к GitHub и запушить

```bash
git remote add origin https://github.com/ВАШ_НИК/my-projects.git
git branch -M main
git push -u origin main
```

---

## Отдельные проекты на GitHub

Каждый большой проект (Autoparts, IFOAM, SplitRouting) лучше хранить как **отдельный репозиторий**.

### Autoparts_Ecommerce

```bash
cd C:\Users\Admin\Documents\Autoparts_Ecommerce

# Проверить что .env есть в .gitignore
cat .gitignore | findstr .env

git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/ВАШ_НИК/autoparts-ecommerce.git
git push -u origin main
```

### IFOAM Ecommerce

```bash
cd C:\Users\Admin\Documents\Ecommerce
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/ВАШ_НИК/ifoam-ecommerce.git
git push -u origin main
```

### SplitRouting

```bash
cd C:\Users\Admin\Documents\SplitRouting
git init
git add .
git commit -m "Initial commit: Pirate Route-chan"
git remote add origin https://github.com/ВАШ_НИК/split-routing.git
git push -u origin main
```

---

## ВАЖНО перед пушем

Проверить что эти файлы НЕ попадут в git:

```bash
git status
# Убедиться что .env нет в списке
```

Если `.env` попал:
```bash
echo ".env" >> .gitignore
git rm --cached .env
git commit -m "Remove .env from tracking"
```

Также удалить или заменить API ключи в файлах:
- `C:\Users\Admin\check_ozon_stocks.py` — ключи вшиты прямо в файл!
- `C:\Users\Admin\Desktop\avtoliga_images.py` — логин/пароль Автолига
