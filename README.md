# Расписание ФТУГ БНТУ

Статический сайт расписания для ФТУГ БНТУ (1 курс) на GitHub Pages.
Источник данных — официальная таблица с https://bntu.by/raspisanie/ftug (кнопка «1 курс»).

## Как это работает

- **Кнопка «1 курс»** на сайте БНТУ скачивает `.xls`: `https://files.bntu.by/s/clKyMFGYUOZhpWT/download`
- **GitHub Actions** (`cron: */8h + ручной запуск`) скачивает свежую таблицу, парсит её в `site/data.json` и деплоит на Pages
- **Сайт** — два экрана: приветствие с выбором группы → расписание (дни, недели 1/2, перерывы, совмещёнки, корпуса на карте)

## Установка одним заходом

```bash
# 1. Создай репозиторий на GitHub (например ftug-schedule)
# 2. Залей в него содержимое этой папки:
git init
git add .
git commit -m "init: расписание ФТУГ"
git branch -M main
git remote add origin https://github.com/<твой-юзернейм>/ftug-schedule.git
git push -u origin main
```

### Настройка GitHub Pages (один раз)

1. GitHub → репозиторий → **Settings → Pages**
2. **Source:** `GitHub Actions` (НЕ «Deploy from a branch» — деплоит workflow)
3. **Custom domain:** `code.vibecoder.help` → Save
4. В настройках **DNS** для `code.vibecoder.help` добавь CNAME-запись на `<твой-юзернейм>.github.io`
5. Открой Actions → запусти workflow вручную (**Run workflow**) → после зелёной галочки сайт готов

## Локальная разработка

```bash
python3 -m pip install xlrd==2.0.2
# обновить данные:
curl -fsSL -o /tmp/schedule.xls "https://files.bntu.by/s/clKyMFGYUOZhpWT/download"
python3 scripts/parse_bntu.py /tmp/schedule.xls site/data.json
# посмотреть:
python3 -m http.server 8000 --directory site
```

## Структура

```
site/                 → всё, что публикуется на Pages
  index.html          → приложение (HTML+CSS+JS, без зависимостей)
  data.json           → расписание (генерируется парсером)
  CNAME               → code.vibecoder.help
scripts/
  parse_bntu.py       → парсер xls → data.json (пути — аргументы)
.github/workflows/
  refresh.yml         → cron каждые 8ч + ручной запуск + авто-деплой
```