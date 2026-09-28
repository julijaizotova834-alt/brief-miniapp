"""
webapp_server.py — Brief AI (UA) Mini App backend
"""
import os, hmac, hashlib, json, tempfile, logging, time
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

app = FastAPI(title="Brief AI (UA) Mini App backend")
app.add_middleware(CORSMiddleware, allow_origins=[ALLOWED_ORIGIN] if ALLOWED_ORIGIN != "*" else ["*"], allow_methods=["*"], allow_headers=["*"])
core.init_db()

UPLOADS = {}
UPLOAD_TTL = 3600
MAX_UPLOAD = 1024*1024*1024

def validate_init_data(init_data):
    if not init_data: raise HTTPException(401, "Немає initData")
    parsed = dict(parse_qsl(init_data, strict_parsing=True))
    h = parsed.pop("hash", None)
    if not h: raise HTTPException(401, "Немає hash")
    dcs = "\n".join(f"{k}={v}" for k, v in sorted(parsed.items()))
    sk = hmac.new(b"WebAppData", TELEGRAM_TOKEN.encode(), hashlib.sha256).digest()
    if not hmac.compare_digest(hmac.new(sk, dcs.encode(), hashlib.sha256).hexdigest(), h):
        raise HTTPException(401, "Невірний підпис")
    if time.time() - int(parsed.get("auth_date", "0")) > 86400:
        raise HTTPException(401, "initData застаріла")
    user = json.loads(parsed.get("user", "{}"))
    if not user.get("id"): raise HTTPException(401, "Немає даних користувача")
    return user

def cleanup():
    now = time.time()
    for uid in [u for u, v in UPLOADS.items() if now - v["created_at"] > UPLOAD_TTL]:
        try: os.remove(UPLOADS[uid]["path"])
        except: pass
        UPLOADS.pop(uid, None)

@app.get("/api/status")
def status(x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    uid = user["id"]; uname = user.get("username", "")
    fname = (user.get("first_name", "") + " " + user.get("last_name", "")).strip()
    core.register_user_if_new(uid, uname, fname)
    info = core.get_user_info(uid)
    if not info:
        return {"is_subscribed": False, "remaining": core.FREE_LIMIT, "free_limit": core.FREE_LIMIT, "audit_credits": 0}
    if core.has_active_sub(info):
        return {"is_subscribed": True, "subscription_until": info['until'].strftime("%d.%m.%Y"), "audit_credits": info['credits']}
    return {"is_subscribed": False, "remaining": max(0, core.FREE_LIMIT - info['usage']), "free_limit": core.FREE_LIMIT, "audit_credits": info['credits']}

@app.post("/api/upload")
async def upload(file: UploadFile = File(...), x_telegram_initdata: str = Header(None)):
    cleanup()
    user = validate_init_data(x_telegram_initdata)
    uid = user["id"]; uname = user.get("username", "")
    if not core.check_access(uid, uname):
        raise HTTPException(403, "Ліміт безкоштовних документів вичерпано. Оформіть підписку: /subscribe")
    suffix = os.path.splitext(file.filename or "audio")[1] or ".audio"
    fd, tmp = tempfile.mkstemp(suffix=suffix); os.close(fd)
    size = 0
    with open(tmp, "wb") as out:
        while True:
            chunk = await file.read(1024*1024)
            if not chunk: break
            size += len(chunk)
            if size > MAX_UPLOAD: out.close(); os.remove(tmp); raise HTTPException(413, "Файл занадто великий.")
            out.write(chunk)
    upload_id = hashlib.sha256(f"{uid}-{tmp}-{time.time()}".encode()).hexdigest()[:24]
    UPLOADS[upload_id] = {"path": tmp, "user_id": uid, "username": uname, "created_at": time.time()}
    return {"upload_id": upload_id, "size_mb": round(size/1024/1024, 1)}

class GenerateRequest(BaseModel):
    upload_id: str
    doc_type: str

@app.post("/api/generate")
async def generate(req: GenerateRequest, x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    uid = user["id"]; uname = user.get("username", "")
    record = UPLOADS.get(req.upload_id)
    if not record or record["user_id"] != uid:
        raise HTTPException(404, "Завантаження не знайдено. Прикріпіть файл ще раз.")
    if req.doc_type not in core.DOC_NAMES:
        raise HTTPException(400, "Невідомий тип документа.")
    if not core.check_access(uid, uname):
        raise HTTPException(403, "Ліміт безкоштовних документів вичерпано.")
    if req.doc_type == 'audit':
        st = core.check_audit_access(uid, uname)
        if st == 'no_sub': raise HTTPException(403, "Ліміт безкоштовних документів вичерпано. Оформіть підписку: /subscribe")
        if st == 'need_credits': raise HTTPException(403, "Кредити на аудит закінчились. Докупіть: /subscribe")
        if st == 'need_sub_for_credits':
            info = core.get_user_info(uid)
            raise HTTPException(403, f"Для аудиту потрібна активна підписка. Ваші {info['credits']} кредитів збережені.")
    path = record["path"]
    try: transcription = core.transcribe_audio_file(path)
    except Exception as e:
        logger.error(f"Транскрибація: {e}"); raise HTTPException(500, "Не вдалося розпізнати запис.")
    finally:
        try: os.remove(path)
        except: pass
        UPLOADS.pop(req.upload_id, None)
    if not transcription or len(transcription.strip()) < 10:
        raise HTTPException(422, "Мовлення не розпізнано.")
    try: doc_name, clean_text = core.generate_document(req.doc_type, transcription)
    except Exception as e:
        logger.error(f"Генерація: {e}"); raise HTTPException(500, "Не вдалося створити документ.")
    is_audit = req.doc_type == 'audit'
    if is_audit:
        a_st = core.check_audit_access(uid, uname)
        if a_st == 'ok':
            if not core.try_use_audit_credit(uid):
                raise HTTPException(403, "Кредити на аудит закінчились.")
    fname = core.safe_filename(doc_name)
    await bot.send_document(chat_id=uid, document=core.build_docx(doc_name, clean_text), filename=f"{fname}.docx", caption="📎 .docx")
    await bot.send_document(chat_id=uid, document=core.build_txt(doc_name, clean_text), filename=f"{fname}.txt", caption="📎 .txt")
    core.increment_usage(uid)
    finish = "✨ Документ сформовано та надіслано.\n\n⚠️ Перевірте цифри, імена та посилання перед використанням."
    result = {"status": "ok", "doc_name": doc_name}
    if is_audit:
        info = core.get_user_info(uid)
        cr = info['credits'] if info else 0
        finish = f"✨ Аудит готовий!\n\n🔍 Залишилось кредитів: {cr}\n\n⚠️ Перевіряйте висновки за вихідним записом."
        result["audit_credits"] = cr
    await bot.send_message(chat_id=uid, text=finish)
    return result
