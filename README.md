# Расписание ААСК

- `public/` — сайт (`index.html`) и данные (`schedule.json`).
- `scripts/build_schedule.py` — скачивает файлы `DO/ДДММ.xls` с altask.ru, разбирает все группы.
- `.github/workflows/update.yml` — запускает скрипт каждые 2 часа и сохраняет изменения.
- `wrangler.jsonc` — настройка Cloudflare (имя воркера и папка с сайтом).

Ручной запуск на своём компьютере: `pip install xlrd openpyxl tzdata`, затем `python scripts/build_schedule.py`.
