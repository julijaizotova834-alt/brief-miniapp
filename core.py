"""
core.py — Brief AI (UA)
Спільна логіка для бота та міні-додатку. Українська версія.
"""
import os, re, logging, tempfile
from io import BytesIO
from datetime import datetime
import psycopg2
from groq import Groq
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

logger = logging.getLogger(__name__)
groq_client = Groq(api_key=os.environ.get('GROQ_API_KEY'))
genai.configure(api_key=os.environ.get('GEMINI_API_KEY'))
gemini_model = genai.GenerativeModel('gemini-2.5-flash')

BOT_NAME = 'brief_ua'
FREE_LIMIT = 7
GROQ_SIZE_LIMIT_MB = 24
PRICE_SUB = 750
SUB_DAYS = 30
AUDIT_PACKS = [{'size':10,'price':150},{'size':50,'price':550},{'size':100,'price':900}]
FREE_USERNAMES = {'black_physicist','juli_aleksandrova','Aliia85'}

PROMPTS = {
    'narada': """Проаналізуй транскрипцію зустрічі та підготуй короткий звіт українською мовою. Структура: 1. Резюме зустрічі (3-5 речень — суть, мета, результат) 2. Основні теми розмови (короткий список) 3. Ключові рішення (короткий список) 4. Домовленості сторін (короткий список) 5. Завдання — кожне в одному рядку за шаблоном: • [Відповідальний] — [Завдання] | Термін: [термін] | [Коментар якщо є] 6. Важливі цифри, суми та терміни 7. Відкриті питання 8. Наступні кроки. Правила: не вигадуй інформацію, якої немає в записі; при помилках розпізнавання — логічно відновлюй зміст; НЕ використовуй таблиці в жодному вигляді; лише заголовки, списки та абзаци; діловий стиль без зайвої «води».""",

    'audit': """Дій як РВП (керівник відділу продажів) та бізнес-тренер з переговорів. Проаналізуй транскрипцію розмови менеджера з клієнтом (або замовником) і сформуй повний розбір дзвінка українською мовою. ПРАВИЛА ФОРМАТУВАННЯ — виконуй суворо: - НЕ використовуй таблиці в жодному вигляді - Кожен пункт — це один короткий абзац або один рядок списку - Лише заголовки, марковані списки та короткі абзаци - Діловий, але прямий стиль — без «води» та повторів ПРАВИЛА ЗМІСТУ: - Не вигадуй інформацію, якої немає в транскрипції - Якщо даних немає — «Не виявлено» або «Не прозвучало в розмові» - При помилках розпізнавання — логічно відновлюй зміст без зміни суті - Будь чесним в оцінці менеджера — не прикрашай і не пом'якшуй СТРУКТУРА: 1. Короткий підсумок Один абзац 3-4 речення: хто з ким спілкувався, мета дзвінка, чим закінчилася розмова, загальне враження. 2. Профіль клієнта На основі того, що клієнт сказав і як себе поводив: • Тип клієнта (гарячий / теплий / холодний) • Рівень обізнаності про продукт/послугу • Ключовий запит — що саме клієнт хоче вирішити • Прихована потреба — що стоїть за запитом, якщо можна визначити • Болі клієнта — які проблеми, страхи або незручності озвучив • Критерії прийняття рішення — за якими параметрами обирає • Бюджет та очікування щодо ціни — якщо прозвучало 3. Заперечення та сумніви клієнта Кожне заперечення у форматі: • [Заперечення] — [як менеджер відпрацював або не відпрацював] Якщо заперечень не було — «Заперечень не прозвучало.» 4. Комерційні умови Якщо обговорювались ціни, терміни, обсяги — коротким списком. Якщо не обговорювались — «Комерційні умови не розглядались.» 5. Оцінка менеджера за критеріями: встановлення контакту, виявлення потреби, активне слухання, презентація рішення, робота із запереченнями, закриття, тон і манера. 6. Помилки менеджера з прикладами. 7. Сильні сторони менеджера з прикладами. 8. Етап угоди та прогноз (етап воронки, ймовірність угоди, що може завадити). 9. Завдання та наступні кроки. 10. Рекомендації менеджеру — 3-5 конкретних рекомендацій.""",

    'client_summary': """Дій як професійний асистент керівника. Проаналізуй транскрипцію зустрічі або дзвінка з клієнтом і сформуй резюме зустрічі українською мовою, яке можна надіслати клієнту одразу після розмови. ПРАВИЛА: НЕ використовуй таблиці; ввічливий діловий стиль; НЕ додавай внутрішні коментарі — клієнт це побачить; не вигадуй інформацію; НЕ включай ціни якщо вони не були остаточно узгоджені. СТРУКТУРА: Звернення, 1. Що обговорили (3-7 пунктів), 2. Домовленості, 3. Наступні кроки (з нашого боку / з вашого боку), 4. Відкриті питання (якщо є). Завершення: «Якщо я щось пропустив — буду радий зворотному зв'язку.» Підпис: «З повагою, [місце для імені]»""",

    'ideas': """Дій як помічник керівника та стратегічний консультант. Проаналізуй голосові нотатки або мозковий штурм і перетвори їх на структурований план дій українською мовою. Правила: не транскрибуй дослівно — структуруй і відбирай головне; не вигадуй цифри яких не було; НЕ використовуй таблиці; управлінський стиль без «води». Структура: 1. Головна суть і мета 2. Ключові ідеї та їхня цінність 3. Стратегічні напрямки 4. Операційні завдання (завдання | відповідальний | пріоритет | термін) 5. Термінові питання 6. Делегування 7. Ризики 8. Ресурси 9. Наступні кроки (24 год / 7 днів / 30 днів) 10. Рекомендація.""",
}

DOC_NAMES = {
    'narada':'Нарада','audit':'Аудит дзвінка менеджера',
    'client_summary':'Резюме для клієнта','ideas':'Структуровані ідеї','transcript':'Транскрипція',
}

def get_db(): return psycopg2.connect(os.environ.get('DATABASE_URL'))

def init_db():
    conn = get_db(); cur = conn.cursor()
    cur.execute('''CREATE TABLE IF NOT EXISTS users (user_id BIGINT, bot_name TEXT, username TEXT, full_name TEXT, registered_at TIMESTAMP DEFAULT NOW(), usage_count INTEGER DEFAULT 0, is_subscribed BOOLEAN DEFAULT FALSE, subscription_until TIMESTAMP, audit_credits INTEGER DEFAULT 0, last_used_at TIMESTAMP, PRIMARY KEY (user_id, bot_name))''')
    cur.execute('''CREATE TABLE IF NOT EXISTS payments (id SERIAL PRIMARY KEY, user_id BIGINT, bot_name TEXT, telegram_payment_charge_id TEXT UNIQUE, stars_amount INTEGER, payload TEXT, paid_at TIMESTAMP DEFAULT NOW())''')
    for col, ct, df in [('audit_credits','INTEGER','0')]:
        try: cur.execute(f"ALTER TABLE users ADD COLUMN {col} {ct} DEFAULT {df}")
        except: conn.rollback()
    try: cur.execute("ALTER TABLE payments ADD COLUMN payload TEXT")
    except: conn.rollback()
    try: cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_pay_charge ON payments(telegram_payment_charge_id)")
    except: conn.rollback()
    conn.commit(); cur.close(); conn.close()

def register_user_if_new(uid, uname, fname):
    conn = get_db(); cur = conn.cursor()
    cur.execute('''INSERT INTO users (user_id,bot_name,username,full_name,registered_at,usage_count,audit_credits,last_used_at) VALUES (%s,%s,%s,%s,NOW(),0,0,NOW()) ON CONFLICT (user_id,bot_name) DO UPDATE SET username=EXCLUDED.username, full_name=EXCLUDED.full_name, last_used_at=NOW()''', (uid, BOT_NAME, uname, fname))
    conn.commit(); cur.close(); conn.close()

def increment_usage(uid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('UPDATE users SET usage_count=usage_count+1, last_used_at=NOW() WHERE user_id=%s AND bot_name=%s', (uid, BOT_NAME))
    conn.commit(); cur.close(); conn.close()

def get_user_info(uid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('SELECT usage_count,is_subscribed,subscription_until,audit_credits FROM users WHERE user_id=%s AND bot_name=%s', (uid, BOT_NAME))
    r = cur.fetchone(); cur.close(); conn.close()
    if not r: return None
    return {'usage':r[0],'sub':r[1],'until':r[2],'credits':r[3] or 0}

def has_active_sub(info): return info and info['sub'] and info['until'] and info['until'] > datetime.now()
def check_access(uid, uname):
    if uname and uname.lower() in FREE_USERNAMES: return True
    info = get_user_info(uid)
    if not info: return True
    if has_active_sub(info): return True
    return info['usage'] < FREE_LIMIT

def check_audit_access(uid, uname):
    if uname and uname.lower() in FREE_USERNAMES: return 'ok'
    info = get_user_info(uid)
    if not info: return 'free_ok'
    if not has_active_sub(info) and info['usage'] < FREE_LIMIT: return 'free_ok'
    if not has_active_sub(info) and info['credits'] > 0: return 'need_sub_for_credits'
    if not has_active_sub(info): return 'no_sub'
    if info['credits'] > 0: return 'ok'
    return 'need_credits'

def try_use_audit_credit(uid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('UPDATE users SET audit_credits=audit_credits-1 WHERE user_id=%s AND bot_name=%s AND audit_credits>0', (uid, BOT_NAME))
    ok = cur.rowcount > 0; conn.commit(); cur.close(); conn.close(); return ok

def clean_markdown(text):
    if not text: return text
    text = re.sub(r'^#{1,6}\s*','',text,flags=re.MULTILINE)
    text = re.sub(r'\*\*(.+?)\*\*',r'\1',text); text = re.sub(r'__(.+?)__',r'\1',text)
    text = re.sub(r'\*(.+?)\*',r'\1',text); text = re.sub(r'_(.+?)_',r'\1',text)
    text = re.sub(r'```[a-zA-Z]*\n?','',text); text = re.sub(r'`(.+?)`',r'\1',text)
    text = re.sub(r'^\s*[\*\-]\s+','• ',text,flags=re.MULTILINE)
    return text.strip()

def build_docx(title, body):
    doc = Document(); s = doc.styles['Normal']; s.font.name='Calibri'; s.font.size=Pt(11)
    h = doc.add_paragraph(); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = h.add_run(title.upper()); run.bold=True; run.font.size=Pt(16); run.font.color.rgb=RGBColor(0x1F,0x3A,0x5F)
    doc.add_paragraph()
    for raw in body.split('\n'):
        line = raw.rstrip()
        if not line.strip(): doc.add_paragraph(); continue
        is_h = bool(re.match(r'^\d+\.\s+\S', line))
        ltrs = [c for c in line if c.isalpha()]
        is_caps = len(ltrs) >= 3 and all(c.isupper() for c in ltrs) and len(line) < 80
        if is_h or is_caps:
            p = doc.add_paragraph(); r = p.add_run(line); r.bold=True; r.font.size=Pt(12); r.font.color.rgb=RGBColor(0x1F,0x3A,0x5F)
        elif line.lstrip().startswith('•'): doc.add_paragraph(line.lstrip()[1:].strip(), style='List Bullet')
        else: doc.add_paragraph(line)
    buf = BytesIO(); doc.save(buf); buf.seek(0); return buf

def build_txt(title, body):
    buf = BytesIO(f"{title.upper()}\n{'='*len(title)}\n\n{body}".encode('utf-8')); buf.seek(0); return buf

def safe_filename(n): return re.sub(r'[^\w\-]','',n.replace(' ','_'))

def compress_audio(p):
    import subprocess; o = p+'.compressed.mp3'
    try:
        subprocess.run(['ffmpeg','-y','-i',p,'-ac','1','-ar','16000','-b:a','32k',o],check=True,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=600)
        if os.path.exists(o) and os.path.getsize(o)>0: return o
    except: pass
    return p

def split_audio_chunks(input_path, chunk_minutes=25):
    import subprocess; out_dir = tempfile.mkdtemp()
    pattern = os.path.join(out_dir,"chunk_%03d.mp3")
    subprocess.run(['ffmpeg','-y','-i',input_path,'-f','segment','-segment_time',str(chunk_minutes*60),'-ac','1','-ar','16000','-b:a','32k',pattern],check=True,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=1800)
    return sorted(os.path.join(out_dir,f) for f in os.listdir(out_dir) if f.startswith("chunk_"))

def transcribe_audio_file(path):
    size_mb = os.path.getsize(path)/(1024*1024)
    chunk_paths = [path]; cleanup = []
    if size_mb > GROQ_SIZE_LIMIT_MB:
        try: chunk_paths = split_audio_chunks(path); cleanup = chunk_paths
        except:
            c = compress_audio(path); chunk_paths = [c]
            if c != path: cleanup = [c]
    parts = []
    try:
        for p in chunk_paths:
            with open(p,'rb') as f:
                t = groq_client.audio.transcriptions.create(file=(os.path.basename(p),f.read()),model='whisper-large-v3-turbo',language='uk')
            parts.append(t.text)
    finally:
        for p in cleanup:
            try: os.remove(p)
            except: pass
    return "\n\n".join(parts)

def generate_document(doc_type, transcription):
    doc_name = DOC_NAMES[doc_type]
    if doc_type == 'transcript': return doc_name, transcription.strip()
    resp = gemini_model.generate_content(f"{PROMPTS[doc_type]}\n\nОсь транскрипція запису:\n\n{transcription}")
    return doc_name, clean_markdown(resp.text)
