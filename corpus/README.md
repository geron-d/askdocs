# Синтетический корпус SupportDesk

Корпус для демо RAG-ассистента. Все данные вымышленные, NDA-safe: продукт SupportDesk, компания, тарифы, ошибки и адреса не существуют.

## Состав

- `manual/*.pdf` — руководство пользователя, 4 части:
  - `01-nachalo-raboty.pdf` — вход, интерфейс, профиль, поиск;
  - `02-rabota-s-zayavkami.pdf` — создание, статусы, SLA, комментарии, закрытие;
  - `03-baza-znaniy.pdf` — статьи, поиск, публичный портал;
  - `04-otchety-integracii.pdf` — отчёты, email, Telegram, REST API, вебхуки.
- `admin/*.docx` — 3 документа:
  - `01-administrirovanie.docx` — пользователи, очереди, SLA, автоматизации, безопасность;
  - `02-tarify-i-oplata.docx` — тарифы, лимиты, оплата, возвраты;
  - `03-usloviya-i-bezopasnost.docx` — доступность, поддержка, данные, инциденты.
- `faq.json` — 40 пар «вопрос-ответ» с категориями.
- `manifest.json` — список файлов, число секций и FAQ-пар.
- `src/` — Markdown-исходники и `faq.json` для пересборки.
- `build.py` — генератор PDF/DOCX/FAQ.

## Пересборка

Нужны Python 3.11+ и Unicode-шрифт для PDF. Рекомендуется открытый DejaVu Sans; пример через Docker:

```bash
docker run --rm \
  -v "$PWD:/work" \
  -w /work python:3.12-slim \
  sh -c "apt-get update -qq && apt-get install -y -qq fonts-dejavu-core && pip install -q -r corpus/requirements.txt && python corpus/build.py"
```

Если DejaVu нет, скрипт попробует Arial из смонтированного каталога `/fonts` (например, `-v "C:\Windows\Fonts:/fonts:ro"`).

Генератор не скачивает данные и не вызывает LLM: весь текст лежит в `corpus/src`, поэтому корпус воспроизводим и проверяем глазами.
