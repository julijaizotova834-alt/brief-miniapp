"""
webapp_server.py
Backend для Telegram Mini App "Brief AI".

Главная задача: принять аудиофайл ЛЮБОГО размера и длительности напрямую через
обычный HTTP-запрос из мини-приложения (минуя ограничение Telegram Bot API
на скачивание файла в 20 МБ / 30 минут), обработать его (Whisper -> Gemini)
и отправить готовый документ пользователю обратно в чат с ботом.

Запускать отдельным процессом/веб-сервисом на Render (можно тот же проект,
но отдельный Web Service, отличный от бота на polling).

Переменные окружения (те же, что у бота, плюс):
  TELEGRAM_TOKEN   — токен бота (нужен и для отправки документов, и для проверки initData)
  GROQ_API_KEY
  GEMINI_API_KEY
  DATABASE_URL
  ALLOWED_ORIGIN   — домен, с которого грузится мини-приложение (для CORS), например
                     https://your-username.github.io

Запуск:
  uvicorn webapp_server:app --host 0.0.0.0 --port 8000
"""

import os
import hmac
import hashlib
import json
import shutil
import tempfile
import logging
import time
from urllib.parse import parse_qsl

from fastapi import FastAPI, UploadFile, File, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import telegram

import core

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("webapp_server")

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "*")

bot = telegram.Bot(token=TELEGRAM_TOKEN)

app = FastAPI(title="Brief AI Mini App backend")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[ALLOWED_ORIGIN] if ALLOWED_ORIGIN != "*" else ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

core.init_db()

# upload_id -> {"path": str, "user_id": int, "username": str, "created_at": float}
# В памяти процесса — достаточно для одного web-инстанса. Если поднимешь несколько
# инстансов backend'а (масштабирование), перенеси этот словарь в Redis или таблицу БД,
# иначе upload и generate могут попасть на разные инстансы.
UPLOADS = {}
UPLOAD_TTL_SECONDS = 60 * 60  # чистим файлы старше часа, на случай если их не забрали

MAX_UPLOAD_BYTES = 1024 * 1024 * 1024  # 1 ГБ — щедрый потолок, поправь под себя


# ---------- ПРОВЕРКА TELEGRAM initData ----------
def validate_init_data(init_data: str) -> dict:
    """
    Проверяет подпись initData, которую присылает Telegram WebApp.
    Алгоритм из официальной документации Telegram Mini Apps.
    Возвращает dict с данными пользователя или бросает HTTPException(401).
    """
    if not init_data:
        raise HTTPException(status_code=401, detail="Нет initData")

    parsed = dict(parse_qsl(init_data, strict_parsing=True))
    received_hash = parsed.pop("hash", None)
    if not received_hash:
        raise HTTPException(status_code=401, detail="Нет hash в initData")

    data_check_string = "\n".join(
        f"{k}={v}" for k, v in sorted(parsed.items())
    )

    secret_key = hmac.new(b"WebAppData", TELEGRAM_TOKEN.encode(), hashlib.sha256).digest()
    computed_hash = hmac.new(secret_key, data_check_string.encode(), hashlib.sha256).hexdigest()

    if not hmac.compare_digest(computed_hash, received_hash):
        raise HTTPException(status_code=401, detail="Неверная подпись initData")

    auth_date = int(parsed.get("auth_date", "0"))
    if time.time() - auth_date > 86400:
        raise HTTPException(status_code=401, detail="initData устарела, переоткройте мини-приложение")

    user = json.loads(parsed.get("user", "{}"))
    if not user.get("id"):
        raise HTTPException(status_code=401, detail="Нет данных пользователя")

    return user


def cleanup_old_uploads():
    now = time.time()
    expired = [uid for uid, v in UPLOADS.items() if now - v["created_at"] > UPLOAD_TTL_SECONDS]
    for uid in expired:
        try:
            os.remove(UPLOADS[uid]["path"])
        except OSError:
            pass
        UPLOADS.pop(uid, None)


# ---------- /api/status ----------
@app.get("/api/status")
def status(x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    user_id = user["id"]
    username = user.get("username", "")
    full_name = (user.get("first_name", "") + " " + user.get("last_name", "")).strip()

    core.register_user_if_new(user_id, username, full_name)
    row = core.get_user_row(user_id)

    if not row:
        return {"is_subscribed": False, "remaining": core.FREE_LIMIT, "free_limit": core.FREE_LIMIT}

    usage_count, is_subscribed, subscription_until = row
    if is_subscribed and subscription_until and subscription_until.timestamp() > time.time():
        return {
            "is_subscribed": True,
            "subscription_until": subscription_until.strftime("%d.%m.%Y"),
        }

    return {
        "is_subscribed": False,
        "remaining": max(0, core.FREE_LIMIT - usage_count),
        "free_limit": core.FREE_LIMIT,
    }


# ---------- /api/upload ----------
@app.post("/api/upload")
async def upload(file: UploadFile = File(...), x_telegram_initdata: str = Header(None)):
    cleanup_old_uploads()
    user = validate_init_data(x_telegram_initdata)
    user_id = user["id"]
    username = user.get("username", "")

    if not core.check_access(user_id, username):
        raise HTTPException(
            status_code=403,
            detail="Ліміт безкоштовних документів вичерпано. Оформіть підписку командою /subscribe в боті."
        )

    suffix = os.path.splitext(file.filename or "audio")[1] or ".audio"
    fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)

    size = 0
    with open(tmp_path, "wb") as out:
        while True:
            chunk = await file.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                out.close()
                os.remove(tmp_path)
                raise HTTPException(status_code=413, detail="Файл занадто великий.")
            out.write(chunk)

    upload_id = hashlib.sha256(f"{user_id}-{tmp_path}-{time.time()}".encode()).hexdigest()[:24]
    UPLOADS[upload_id] = {
        "path": tmp_path,
        "user_id": user_id,
        "username": username,
        "created_at": time.time(),
    }

    logger.info(f"Файл загружен: user={user_id}, size={size/1024/1024:.1f} МБ, upload_id={upload_id}")

    return {"upload_id": upload_id, "size_mb": round(size / 1024 / 1024, 1)}


# ---------- /api/generate ----------
class GenerateRequest(BaseModel):
    upload_id: str
    doc_type: str


@app.post("/api/generate")
async def generate(req: GenerateRequest, x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    user_id = user["id"]
    username = user.get("username", "")

    record = UPLOADS.get(req.upload_id)
    if not record or record["user_id"] != user_id:
        raise HTTPException(status_code=404, detail="Завантаження не знайдено. Прикріпіть файл ще раз.")

    if req.doc_type not in core.DOC_NAMES:
        raise HTTPException(status_code=400, detail="Невідомий тип документа.")

    if not core.check_access(user_id, username):
        raise HTTPException(status_code=403, detail="Ліміт безкоштовних документів вичерпано.")

    path = record["path"]
    try:
        transcription = core.transcribe_audio_file(path)
    except Exception as e:
        logger.error(f"Ошибка транскрибации: {e}")
        raise HTTPException(status_code=500, detail="Не вдалося розпізнати запис.")
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
        UPLOADS.pop(req.upload_id, None)

    if not transcription or len(transcription.strip()) < 10:
        raise HTTPException(status_code=422, detail="Мова в записі не розпізнана.")

    try:
        doc_name, clean_text = core.generate_document(req.doc_type, transcription)
    except Exception as e:
        logger.error(f"Ошибка генерации документа: {e}")
        raise HTTPException(status_code=500, detail="Не вдалося скласти документ.")

    fname = core.safe_filename(doc_name)

    docx_buf = core.build_docx(doc_name, clean_text)
    txt_buf = core.build_txt(doc_name, clean_text)

    await bot.send_document(
        chat_id=user_id,
        document=docx_buf,
        filename=f"{fname}.docx",
        caption="📎 .docx",
    )
    await bot.send_document(
        chat_id=user_id,
        document=txt_buf,
        filename=f"{fname}.txt",
        caption="📎 .txt",
    )
    await bot.send_message(
        chat_id=user_id,
        text=(
            "✨ Документ сформовано та надіслано.\n\n"
            "⚠️ Перевірте цифри, імена та правові посилання перед використанням."
        ),
    )

    core.increment_usage(user_id)

    return {"status": "ok", "doc_name": doc_name}
