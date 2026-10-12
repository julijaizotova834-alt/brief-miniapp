"""Skrepka AI Mini App backend (українська версія)."""
import os
import hmac
import hashlib
import json
import tempfile
import logging
import time
import asyncio
import threading
from urllib.parse import parse_qsl

from fastapi import FastAPI, UploadFile, File, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from telegram import Bot, ReplyKeyboardMarkup, KeyboardButton

import core

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("webapp_server")

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
if not TELEGRAM_TOKEN:
    logger.warning("TELEGRAM_TOKEN is not configured")

ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "*")
bot = Bot(token=TELEGRAM_TOKEN) if TELEGRAM_TOKEN else None

app = FastAPI(title="Skrepka AI Mini App backend")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[ALLOWED_ORIGIN] if ALLOWED_ORIGIN != "*" else ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
core.init_db()

UPLOADS = {}
UPLOAD_TTL = 3600
MAX_UPLOAD = int(os.environ.get('MAX_UPLOAD_BYTES', str(1024 * 1024 * 1024)))
MAX_PENDING_UPLOADS = int(os.environ.get('MAX_PENDING_UPLOADS', '300'))
UPLOAD_LOCK = threading.Lock()
PROCESSING = set()
MENU_KEYBOARD = ReplyKeyboardMarkup(
    [[KeyboardButton("🏠 Головне меню"), KeyboardButton("💫 Мій тариф")]],
    resize_keyboard=True,
    is_persistent=True,
)


def validate_init_data(init_data: str | None):
    if not init_data:
        raise HTTPException(status_code=401, detail="Немає initData")
    try:
        pairs = parse_qsl(init_data, strict_parsing=True, keep_blank_values=True)
        if len({k for k, _ in pairs}) != len(pairs):
            raise ValueError('duplicate initData keys')
        parsed = dict(pairs)
    except ValueError:
        raise HTTPException(status_code=401, detail="Некоректна initData")
    received_hash = parsed.pop("hash", None)
    if not received_hash:
        raise HTTPException(status_code=401, detail="Немає hash")
    if not TELEGRAM_TOKEN:
        raise HTTPException(status_code=503, detail="Сервіс тимчасово недоступний")
    data_check_string = "\n".join(f"{k}={v}" for k, v in sorted(parsed.items()))
    secret_key = hmac.new(b"WebAppData", TELEGRAM_TOKEN.encode(), hashlib.sha256).digest()
    expected_hash = hmac.new(secret_key, data_check_string.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected_hash, received_hash):
        raise HTTPException(status_code=401, detail="Невірний підпис")
    try:
        auth_date = int(parsed.get("auth_date", "0"))
    except ValueError:
        raise HTTPException(status_code=401, detail="Некоректна дата авторизації")
    if auth_date <= 0 or time.time() - auth_date > 86400 or auth_date > time.time() + 60:
        raise HTTPException(status_code=401, detail="initData застаріла")
    try:
        user = json.loads(parsed.get("user", "{}"))
    except json.JSONDecodeError:
        raise HTTPException(status_code=401, detail="Некоректні дані користувача")
    try:
        user_id = int(user.get("id", 0))
    except (TypeError, ValueError):
        user_id = 0
    if user_id <= 0:
        raise HTTPException(status_code=401, detail="Немає даних користувача")
    user["id"] = user_id
    return user


def cleanup():
    now = time.time()
    with UPLOAD_LOCK:
        expired = [upload_id for upload_id, record in UPLOADS.items()
                   if upload_id not in PROCESSING and now - record["created_at"] > UPLOAD_TTL]
        records = [UPLOADS.pop(upload_id, None) for upload_id in expired]
    for record in records:
        if record:
            try: os.remove(record["path"])
            except OSError: pass


def user_details(user):
    uid = int(user["id"])
    username = user.get("username", "")
    full_name = (user.get("first_name", "") + " " + user.get("last_name", "")).strip()
    core.register_user_if_new(uid, username, full_name)
    info = core.get_user_info(uid)
    if not info:
        info = {"usage": 0, "sub": False, "until": None, "credits": 0}
    active = core.has_active_sub(info)
    return uid, username, info, active


@app.get("/api/status")
def status(x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    _, _, info, active = user_details(user)
    result = {
        "is_subscribed": bool(active),
        "remaining": max(0, core.FREE_LIMIT - info["usage"]),
        "free_limit": core.FREE_LIMIT,
        "audit_credits": info["credits"],
        "subscription_until": info["until"].strftime("%d.%m.%Y") if active else None,
        "price_sub": core.PRICE_SUB,
        "audit_packs": core.AUDIT_PACKS,
    }
    return result


@app.post("/api/upload")
async def upload(file: UploadFile = File(...), x_telegram_initdata: str = Header(None)):
    """Завантаження дозволене до вибору типу документа. Доступ перевіряється в /generate."""
    cleanup()
    user = validate_init_data(x_telegram_initdata)
    uid = int(user["id"])
    uname = user.get("username", "")
    user_details(user)
    with UPLOAD_LOCK:
        if len(UPLOADS) >= MAX_PENDING_UPLOADS:
            raise HTTPException(status_code=503, detail="Сервіс тимчасово перевантажений. Спробуйте пізніше.")
    suffix = os.path.splitext(file.filename or "audio")[1].lower()
    if not suffix or len(suffix) > 12 or not suffix[1:].isalnum():
        suffix = ".audio"
    fd, tmp_path = tempfile.mkstemp(prefix="skrepka_", suffix=suffix)
    os.close(fd)
    size = 0
    try:
        with open(tmp_path, "wb") as out:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                size += len(chunk)
                if size > MAX_UPLOAD:
                    raise HTTPException(status_code=413, detail="Файл завеликий (максимум 1 ГБ).")
                out.write(chunk)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise
    finally:
        await file.close()
    if size == 0:
        try: os.remove(tmp_path)
        except OSError: pass
        raise HTTPException(status_code=400, detail="Файл порожній. Оберіть аудіозапис.")
    upload_id = hashlib.sha256(f"{uid}-{tmp_path}-{time.time()}".encode()).hexdigest()[:24]
    with UPLOAD_LOCK:
        UPLOADS[upload_id] = {
            "path": tmp_path,
            "user_id": uid,
            "username": uname,
            "created_at": time.time(),
        }
    return {"upload_id": upload_id, "size_mb": round(size / 1024 / 1024, 1)}


class GenerateRequest(BaseModel):
    upload_id: str
    doc_type: str


@app.post("/api/generate")
async def generate(req: GenerateRequest, x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    uid, uname, initial_info, initial_active = user_details(user)
    if req.doc_type not in core.DOC_NAMES:
        raise HTTPException(status_code=400, detail="Невідомий тип документа.")

    with UPLOAD_LOCK:
        record = UPLOADS.get(req.upload_id)
        if not record or record["user_id"] != uid:
            raise HTTPException(status_code=404, detail="Завантаження не знайдено. Прикріпіть файл ще раз.")
        if req.upload_id in PROCESSING:
            raise HTTPException(status_code=409, detail="Цей запис вже обробляється.")
        PROCESSING.add(req.upload_id)
        UPLOADS.pop(req.upload_id, None)

    path = record["path"]
    is_audit = req.doc_type == "audit"
    paid_audit = False
    free_quota_reserved = False
    delivered = False
    try:
        if is_audit:
            audit_status = core.check_audit_access(uid, uname)
            if audit_status == "need_credits":
                raise HTTPException(status_code=403, detail="Безкоштовні обробки закінчилися, а кредитів на аудит немає. Купіть пакет аудитів у Telegram-боті.")
            if audit_status == "credit":
                if not core.try_use_audit_credit(uid):
                    raise HTTPException(status_code=409, detail="Кредити на аудит закінчилися. Оновіть баланс і спробуйте знову.")
                paid_audit = True
            elif audit_status == "free_ok":
                if not core.reserve_free_usage(uid, uname):
                    raise HTTPException(status_code=403, detail="Безкоштовний ліміт вичерпано. Купіть кредити на аудит або оформіть підписку для основних функцій.")
                free_quota_reserved = not (uname and uname.strip().lstrip('@').lower() in core.FREE_USERNAMES) and not initial_active
        else:
            if not core.check_access(uid, uname):
                raise HTTPException(status_code=403, detail="Безкоштовний ліміт основних функцій вичерпано. Оформіть підписку Skrepka у Telegram-боті.")
            if not (uname and uname.strip().lstrip('@').lower() in core.FREE_USERNAMES) and not initial_active:
                if not core.reserve_free_usage(uid, uname):
                    raise HTTPException(status_code=403, detail="Безкоштовний ліміт основних функцій вичерпано. Оформіть підписку Skrepka у Telegram-боті.")
                free_quota_reserved = True

        try:
            transcription = await asyncio.to_thread(core.transcribe_audio_file, path)
        except Exception as exc:
            logger.exception("Транскрибація завершилася помилкою: %s", exc)
            raise HTTPException(status_code=500, detail="Не вдалося розпізнати запис. Спробуйте ще раз.")
        if not transcription or len(transcription.strip()) < 10:
            raise HTTPException(status_code=422, detail="Мовлення не розпізнано. Спробуйте запис з більш розбірливим мовленням.")
        try:
            doc_name, clean_text = await asyncio.to_thread(core.generate_document, req.doc_type, transcription)
            docx_file = core.build_docx(doc_name, clean_text)
        except Exception as exc:
            logger.exception("Генерація документа завершилася помилкою: %s", exc)
            raise HTTPException(status_code=500, detail="Не вдалося скласти документ. Спробуйте ще раз.")

        try:
            if not bot:
                raise RuntimeError("Telegram bot is not configured")
            filename = core.safe_filename(doc_name) or "Skrepka_document"
            await bot.send_document(chat_id=uid, document=docx_file, filename=f"{filename}.docx", caption="📎 Документ Word (.docx)")
            delivered = True
        except Exception as exc:
            logger.exception("Не вдалося надіслати результат у Telegram: %s", exc)
            raise HTTPException(status_code=502, detail="Не вдалося надіслати документ у Telegram. Спробуйте завантажити запис ще раз.")

        info = core.get_user_info(uid) or {"credits": 0, "usage": 0, "sub": False, "until": None}
        if is_audit:
            finish = ("✨ Аудит готовий!\n\nДокумент Word надіслано в чат.\n"
                      f"\n🔍 Залишилося кредитів на аудит: {info['credits']}"
                      "\n\n⚠️ Перевіряйте висновки за вихідним записом.")
        else:
            finish = "✨ Документ готовий! Документ Word (.docx) надіслано в чат.\n\n⚠️ Перевірте цифри, імена та посилання перед використанням."
        try:
            await bot.send_message(chat_id=uid, text=finish, reply_markup=MENU_KEYBOARD)
        except Exception:
            logger.exception("Не вдалося надіслати завершальне повідомлення з меню")
        info = core.get_user_info(uid) or info
        return {"status": "ok", "doc_name": doc_name, "audit_credits": info["credits"],
                "is_subscribed": core.has_active_sub(info),
                "remaining": max(0, core.FREE_LIMIT - info["usage"]), "message": finish}
    except Exception:
        if not delivered and paid_audit:
            try: core.refund_audit_credit(uid)
            except Exception: logger.exception("Не вдалося повернути кредит після помилки")
        if not delivered and free_quota_reserved:
            try: core.release_free_usage(uid)
            except Exception: logger.exception("Не вдалося повернути безкоштовну обробку після помилки")
        raise
    finally:
        try: os.remove(path)
        except OSError: pass
        with UPLOAD_LOCK:
            PROCESSING.discard(req.upload_id)
