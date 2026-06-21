# Brief AI — мини-приложение (Telegram Mini App)

Решает то, что не может обычный бот: принимает аудиозапись **без лимита** в 20 МБ
и 30 минут, потому что файл идёт прямым HTTP-запросом на твой сервер, а не через
`bot.get_file()` (там лимит зашит на стороне Telegram Bot API).

## Файлы

- `index.html` — само мини-приложение (чат-интерфейс), статический файл
- `core.py` — общая логика (DB, промпты, Whisper, Gemini, сборка .docx/.txt) — перенеси
  туда содержимое словаря `PROMPTS` из своего текущего `bot.py` (он не менялся, просто
  его не дублирую здесь по объёму)
- `webapp_server.py` — backend на FastAPI: `/api/status`, `/api/upload`, `/api/generate`
- `requirements.txt` — зависимости backend'а

## Шаг 1 — задеплоить backend

Отдельный Web Service на Render (не Background Worker, как обычный бот — этому нужен
публичный HTTPS-адрес):

```
uvicorn webapp_server:app --host 0.0.0.0 --port $PORT
```

Переменные окружения — те же, что у бота (`TELEGRAM_TOKEN`, `GROQ_API_KEY`,
`GEMINI_API_KEY`, `DATABASE_URL`), плюс `ALLOWED_ORIGIN` — домен, где будет жить
`index.html`.

⚠️ На Render для аудиообработки нужен `ffmpeg` в окружении — как и для обычного бота,
добавь его через `apt.txt` / Dockerfile, если ещё не настроено.

## Шаг 2 — задеплоить index.html

Любой статический хостинг с HTTPS: GitHub Pages, Render Static Site, Cloudflare Pages.
Telegram требует именно HTTPS для Mini Apps — localhost/HTTP не подойдёт.

В `index.html` поменяй:
```js
const API_BASE = "https://YOUR-BACKEND-DOMAIN.example.com";
```
на адрес backend'а из шага 1.

## Шаг 3 — подключить кнопку в боте

Через @BotFather:
```
/mybots → выбрать бота → Bot Settings → Menu Button → Edit menu button URL
```
Вставить ссылку на задеплоенный `index.html`.

Либо программно, кнопкой в самом боте (добавь в `bot.py`):
```python
from telegram import WebAppInfo, InlineKeyboardButton, InlineKeyboardMarkup

async def open_app(update, context):
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("🎙 Открыть Brief AI", web_app=WebAppInfo(url="https://your-app-url.example.com"))
    ]])
    await update.message.reply_text("Открыть мини-приложение:", reply_markup=kb)

app.add_handler(CommandHandler('app', open_app))
```

## Как это работает для длинных записей

1. Пользователь открывает мини-приложение прямо в Telegram
2. Прикрепляет файл любого размера — он грузится напрямую на твой сервер через `fetch`
3. Backend, если файл длиннее ~25 минут, режет его на куски через `ffmpeg` (`split_audio_chunks`
   в `core.py`) и прогоняет через Whisper кусками, затем склеивает текст
4. Выбранный тип документа обрабатывается тем же Gemini-промптом, что и в обычном боте
5. Готовые `.docx` и `.txt` отправляются ботом прямо в чат с пользователем — мини-приложение
   не занимается скачиванием файлов, только показывает статус

## Что стоит докрутить перед продакшеном

- `UPLOADS` в `webapp_server.py` живёт в памяти процесса — если поднимешь больше одного
  инстанса backend'а, перенеси в Redis или таблицу в Postgres
- Сейчас лимит одного аудио — 1 ГБ (`MAX_UPLOAD_BYTES`), поправь под реальные записи
- Стоит добавить прогресс-бар на реальный % (сейчас анимация декоративная, без реального
  прогресса транскрибации) — можно через `EventSource`/WebSocket, если важно
