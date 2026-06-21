"""
core.py
Общая логика Brief AI: база данных, промпты, генерация документов, транскрибация.
Используется и обычным ботом (bot.py), и backend'ом мини-приложения (webapp_server.py),
чтобы не дублировать промпты/DB-функции в двух местах.

Перенеси сюда содержимое из своего текущего bot.py (DB-функции, PROMPTS, DOC_NAMES,
build_docx, build_txt, clean_markdown, safe_filename, compress_audio) — структура и
сигнатуры функций ниже совпадают с тем, что у тебя уже написано, только добавлена
transcribe_audio_file(path), которая работает с обычным путём к файлу на диске,
а не с Telegram File — это и даёт мини-приложению возможность принимать файлы
любого размера (выгрузка идёт напрямую на сервер, а не через Bot API).
"""

import os
import re
import logging
import tempfile
from io import BytesIO
from datetime import datetime

import psycopg2
from groq import Groq
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

logger = logging.getLogger(__name__)

GROQ_API_KEY = os.environ.get('GROQ_API_KEY')
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY')

groq_client = Groq(api_key=GROQ_API_KEY)
genai.configure(api_key=GEMINI_API_KEY)
gemini_model = genai.GenerativeModel('gemini-2.5-flash')

BOT_NAME = 'brief_ua'
SUBSCRIPTION_PRICE_STARS = 620
SUBSCRIPTION_DAYS = 30
FREE_LIMIT = 7
GROQ_SIZE_LIMIT_MB = 24

FREE_USERNAMES = {
    'black_physicist',
    'juli_aleksandrova',
}

# ---------- ПРОМПТИ ----------
PROMPTS = {

    'narada': """Проаналізуй транскрипцію зустрічі та підготуй стислий звіт українською мовою. Структура: 1. Резюме зустрічі (3-5 речень - суть, мета, результат) 2. Основні теми розмови (короткий список) 3. Ключові рішення (короткий список) 4. Домовленості сторін (короткий список) 5. Завдання - кожне в одному рядку за шаблоном: • [Відповідальний] - [Завдання] | Термін: [термін] | [Коментар якщо є] 6. Важливі цифри, суми та строки 7. Відкриті питання 8. Наступні кроки. Правила: не вигадуй інформацію якої немає у записі; при помилках розпізнавання - логічно віднови зміст; НЕ використовуй таблиці у жодному вигляді; лише заголовки, списки та абзаци; діловий стиль без зайвої «води».""",

    'protocol': """Дій як бізнес-консультант та керівник проєктів. Проаналізуй транскрипцію ділової зустрічі та сформуй управлінський протокол українською мовою. ПРАВИЛА ФОРМАТУВАННЯ - виконуй суворо: - НЕ використовуй таблиці у жодному вигляді - НЕ роби вкладені підпункти «Тема:», «Рішення:», «Домовленості:» - це робить текст розляпистим - Кожен пункт - це один короткий абзац або один рядок списку - Лише заголовки, марковані списки та короткі абзаци - Діловий стиль без «води» та повторів ПРАВИЛА ЗМІСТУ: - Не вигадуй інформацію якої немає у транскрипції - Якщо дані відсутні - «Не вказано» або «Не визначено» - При помилках розпізнавання - логічно віднови зміст без зміни суті СТРУКТУРА: 1. Підсумок зустрічі Один абзац 3-4 речення: хто зустрівся, з якою метою, який головний результат. 2. Ключові теми та рішення Для кожної теми - один рядок списку у форматі: • [Назва теми]: [суть рішення або домовленості] Приклад: • Взаємодія з Радобанком: готується претензія та позов до суду. НЕ розбивай одну тему на окремі рядки «Тема / Рішення / Домовленості». 3. Фінанси та умови Якщо фінансові дані були - коротким списком: • [Що]: [сума / строк / умова] Якщо фінансових даних не було - напиши одним рядком: «Фінансові умови не обговорювались.» 4. Ризики та відкриті питання Ризики - короткий список без підзаголовків: • [Ризик одним реченням] Відкриті питання - короткий список: • [Питання одним реченням] В кінці одним рядком: Рекомендація: [запускати / допрацювати / відкласти / відмовитися] - [одне речення чому]. 5. Реєстр завдань Кожне завдання - ОДИН рядок: • [Відповідальний] - [Завдання] | Термін: [термін] | [Статус якщо є] Якщо відповідальний невідомий - «Уточнити». Якщо термін невідомий - «Не визначено». 6. Наступні кроки 2-3 пункти списком: що, хто, коли. 7. Шаблон комерційної пропозиції Тільки якщо обговорювалися продажі або партнерство. Якщо зустріч внутрішня - цей розділ не включай взагалі.""",

    'memorandum': """Дій як корпоративний юрист та комерційний директор. Проаналізуй транскрипцію переговорів та сформуй Меморандум про наміри українською мовою. ПРАВИЛА ФОРМАТУВАННЯ - виконуй суворо: - НЕ використовуй таблиці у жодному вигляді - НЕ роби великих вертикальних списків із одного слова або коротких фраз - НЕ роби розриви між кожним пунктом - Використовуй компактний діловий формат: один пункт = один рядок або короткий абзац - Лише заголовки, короткі списки та абзаци - Офіційно-діловий юридично нейтральний стиль без «води». ПРАВИЛА ЗМІСТУ: - Не вигадуй цифри, строки, штрафи або домовленості яких не було - Якщо дані відсутні - «Не визначено» або «Потребує погодження» - Чітко розділяй Сторону 1 та Сторону 2 - Якщо не погоджені штрафи, строки або юрисдикція - «Підлягає окремому юридичному узгодженню». СТРУКТУРА: 1. Загальна інформація - дата, учасники, компанії, посади. 2. Предмет та мета переговорів - один компактний абзац 3-4 речення. 3. Потреби та інтереси сторін: • Сторона 1 - [коротко в один блок] • Сторона 2 - [коротко в один блок]. 4. Ризики та обмеження - короткий компактний список без великих розривів. 5. Попередні комерційні умови - подавай тільки компактними рядками у форматі: • Бюджет: ... • Оплата: ... • Строки: ... • KPI: ... • Модель співпраці: ... • Інші умови: ... НЕ роби окремі підпункти або великі списки. 6. Розподіл відповідальності - тільки компактний формат у рядок: • Сторона 1: [перелік обов'язків через крапку з комою] • Сторона 2: [перелік обов'язків через крапку з комою]. НЕ розписуй кожну людину окремим блоком. 7. Наступні кроки та терміни - тільки компактний список у форматі: • [Дія] - [відповідальний] - [термін]. НЕ групуй на «терміново», «середньостроково» тощо. 8. Врегулювання спорів - один короткий абзац. 9. Висновок - рекомендація: запуск / підготовка пропозиції / продовження / пауза / відмова + одне коротке пояснення. 10. Статус документа - один короткий абзац про попередній характер меморандуму. 11. Реквізити сторін - компанія, ЄДРПОУ, представник, посада, підпис, дата.""",

    'letter': """Дій як корпоративний юрист. Проаналізуй транскрипцію та сформуй ГОТОВИЙ офіційний діловий лист українською мовою для негайного надсилання адресату. ПРАВИЛА ФОРМАТУВАННЯ - виконуй суворо: - НЕ пиши вступи типу «Діючи як корпоративний юрист», «Я проаналізував транскрипцію», «Тип листа» або будь-які пояснення AI - НЕ додавай заголовок «ЮРИДИЧНИЙ ДІЛОВИЙ ЛИСТ» - НЕ використовуй загальні назви документів - Автоматично визнач тему та тип листа і, якщо доречно, використовуй короткий реальний заголовок за змістом, наприклад: «Щодо затримки поставки обладнання», «Претензія щодо оплати», «Повідомлення про зміну реквізитів» або «Відповідь на звернення» - НЕ використовуй нумерацію розділів - Лист повинен виглядати як реальний готовий документ для відправки - Починай одразу з реквізитів або звернення - НЕ використовуй таблиці - Лише абзаци, короткі блоки та списки якщо доречно - Офіційно-діловий стиль без емоцій, агресії або «води». ПРАВИЛА ЗМІСТУ: - Не вигадуй факти, суми, договори або порушення яких немає у транскрипції - Якщо даних немає - «Не зазначено» або пропусти пункт якщо це не критично - Якщо доречно - ОБОВ'ЯЗКОВО коротко посилайся на релевантні норми законодавства України (Цивільний кодекс України, Господарський кодекс України, КЗпП, Податковий кодекс України, Закон України «Про захист прав споживачів» або інші релевантні норми) та ВКАЗУЙ конкретні статті й кодекси, наприклад: «відповідно до ст. 617 Цивільного кодексу України» або «згідно зі ст. 218 Господарського кодексу України»; НЕ цитуй закони повністю та НЕ вигадуй правові підстави - Автоматично визнач тип листа та адаптуй тон: претензія, повідомлення, запит, вимога, нагадування, відповідь або комерційний лист. СТРУКТУРА: Відправник - ПІБ, посада, компанія, контакти. Одержувач - ПІБ або компанія, посада, адреса. Дата. Тема листа - коротко і по суті. Звернення. Основний текст - суть питання, обставини, прохання, вимога або позиція сторони; суми та строки якщо були. Правове обґрунтування - коротко і тільки якщо доречно, з конкретними статтями законодавства України. Завершення - прохання надати відповідь або виконати зобов'язання. Підпис - «З повагою», ПІБ, посада, компанія, дата. НЕ додавай жодних пояснень перед листом або після нього.""",

    'ideas': """Дій як помічник керівника та стратегічний консультант. Проаналізуй голосові нотатки або мозковий штурм та перетвори їх у структурований план дій українською мовою. Правила: не транскрибуй текст дослівно - структуруй та відбирай головне; не вигадуй цифри, бюджети або терміни яких не було; якщо даних немає - «Не визначено»; НЕ використовуй таблиці у жодному вигляді - лише заголовки, марковані списки та абзаци; управлінський стиль без «води»; відповідай компактно та без повторів. Структура: 1. Головна суть та ціль (3-4 речення) 2. Ключові ідеї та їхня цінність (список) 3. Стратегічні напрямки та перспективи 4. Операційні задачі - кожна в одному рядку: • [Задача] | Відповідальний: [хто] | Пріоритет: [високий/середній/низький] | Термін: [коли] 5. Термінові питання та критичні блокери 6. Що делегувати та кому (ролі або спеціалісти) 7. Ризики - фінансові, юридичні, технічні або операційні (коротко) 8. Ресурси - люди, інструменти, бюджет 9. Наступні кроки - план на 24 год / 7 днів / 30 днів 10. Рекомендація - що найперспективніше та що перевірити першочергово. Якщо у записі є ідея продукту або бізнесу - додатково: цільова аудиторія, основна цінність, мінімальна версія, монетизація, ключові функції, складність запуску.""",

}

DOC_NAMES = {
    'narada': 'Нарада',
    'protocol': 'Протокол переговорів',
    'memorandum': 'Меморандум про наміри',
    'letter': 'Юридичний діловий лист',
    'ideas': 'Структуровані ідеї',
    'transcript': 'Транскрипція',
}


# ---------- БАЗА ДАНИХ ----------
def get_db():
    return psycopg2.connect(os.environ.get('DATABASE_URL'))


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id BIGINT,
            bot_name TEXT,
            username TEXT,
            full_name TEXT,
            registered_at TIMESTAMP DEFAULT NOW(),
            usage_count INTEGER DEFAULT 0,
            is_subscribed BOOLEAN DEFAULT FALSE,
            subscription_until TIMESTAMP,
            last_used_at TIMESTAMP,
            PRIMARY KEY (user_id, bot_name)
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS payments (
            id SERIAL PRIMARY KEY,
            user_id BIGINT,
            bot_name TEXT,
            telegram_payment_charge_id TEXT,
            stars_amount INTEGER,
            paid_at TIMESTAMP DEFAULT NOW()
        )
    ''')
    conn.commit()
    cur.close()
    conn.close()


def register_user_if_new(user_id: int, username: str, full_name: str):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        INSERT INTO users (user_id, bot_name, username, full_name, registered_at, usage_count, last_used_at)
        VALUES (%s, %s, %s, %s, NOW(), 0, NOW())
        ON CONFLICT (user_id, bot_name) DO UPDATE
        SET username = EXCLUDED.username,
            full_name = EXCLUDED.full_name,
            last_used_at = NOW()
    ''', (user_id, BOT_NAME, username, full_name))
    conn.commit()
    cur.close()
    conn.close()


def increment_usage(user_id: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        UPDATE users SET usage_count = usage_count + 1, last_used_at = NOW()
        WHERE user_id = %s AND bot_name = %s
    ''', (user_id, BOT_NAME))
    conn.commit()
    cur.close()
    conn.close()


def get_user_row(user_id: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT usage_count, is_subscribed, subscription_until
        FROM users WHERE user_id = %s AND bot_name = %s
    ''', (user_id, BOT_NAME))
    row = cur.fetchone()
    cur.close()
    conn.close()
    return row


def check_access(user_id: int, username: str) -> bool:
    if username and username.lower() in FREE_USERNAMES:
        return True
    row = get_user_row(user_id)
    if not row:
        return True
    usage_count, is_subscribed, subscription_until = row
    if is_subscribed and subscription_until and subscription_until > datetime.now():
        return True
    return usage_count < FREE_LIMIT


def activate_subscription(user_id: int, days: int, charge_id: str, stars_amount: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        UPDATE users
        SET is_subscribed = TRUE,
            subscription_until = NOW() + INTERVAL '%s days'
        WHERE user_id = %s AND bot_name = %s
    ''', (days, user_id, BOT_NAME))
    cur.execute('''
        INSERT INTO payments (user_id, bot_name, telegram_payment_charge_id, stars_amount, paid_at)
        VALUES (%s, %s, %s, %s, NOW())
    ''', (user_id, BOT_NAME, charge_id, stars_amount))
    conn.commit()
    cur.close()
    conn.close()


# ---------- ТЕКСТ / ДОКУМЕНТИ ----------
def clean_markdown(text: str) -> str:
    if not text:
        return text
    text = re.sub(r'^#{1,6}\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'_(.+?)_', r'\1', text)
    text = re.sub(r'```[a-zA-Z]*\n?', '', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'^\s*[\*\-]\s+', '• ', text, flags=re.MULTILINE)
    return text.strip()


def build_docx(title: str, body: str) -> BytesIO:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = h.add_run(title.upper())
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    doc.add_paragraph()
    for raw_line in body.split('\n'):
        line = raw_line.rstrip()
        if not line.strip():
            doc.add_paragraph()
            continue
        is_numbered_heading = bool(re.match(r'^\d+\.\s+\S', line))
        letters = [c for c in line if c.isalpha()]
        is_caps_heading = (
            len(letters) >= 3
            and all(c.isupper() for c in letters)
            and len(line) < 80
        )
        if is_numbered_heading or is_caps_heading:
            p = doc.add_paragraph()
            r = p.add_run(line)
            r.bold = True
            r.font.size = Pt(12)
            r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
        elif line.lstrip().startswith('•'):
            doc.add_paragraph(line.lstrip()[1:].strip(), style='List Bullet')
        else:
            doc.add_paragraph(line)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_txt(title: str, body: str) -> BytesIO:
    content = f"{title.upper()}\n{'=' * len(title)}\n\n{body}"
    buf = BytesIO(content.encode('utf-8'))
    buf.seek(0)
    return buf


def safe_filename(name: str) -> str:
    name = name.replace(' ', '_')
    return re.sub(r'[^\w\-]', '', name)


def compress_audio(input_path: str) -> str:
    import subprocess
    output_path = input_path + '.compressed.mp3'
    try:
        subprocess.run(
            ['ffmpeg', '-y', '-i', input_path, '-ac', '1', '-ar', '16000', '-b:a', '32k', output_path],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600,
        )
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            return output_path
    except FileNotFoundError:
        logger.warning("ffmpeg не знайдено - надсилаю файл без стиснення.")
    except Exception as e:
        logger.warning(f"Стиснення не вдалося ({e}) - надсилаю файл без стиснення.")
    return input_path


def split_audio_chunks(input_path: str, chunk_minutes: int = 25) -> list[str]:
    """
    Для очень длинных записей (мини-приложение не ограничено 30 минутами, в отличие
    от обычного бота) — режем аудио на куски по N минут перед отправкой в Whisper,
    т.к. у Groq есть лимит размера запроса (~24-25 МБ).
    Возвращает список путей к временным файлам-кускам.
    """
    import subprocess
    out_dir = tempfile.mkdtemp()
    pattern = os.path.join(out_dir, "chunk_%03d.mp3")
    subprocess.run(
        [
            'ffmpeg', '-y', '-i', input_path,
            '-f', 'segment', '-segment_time', str(chunk_minutes * 60),
            '-ac', '1', '-ar', '16000', '-b:a', '32k',
            pattern,
        ],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=1800,
    )
    chunks = sorted(
        os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.startswith("chunk_")
    )
    return chunks


def transcribe_audio_file(path: str) -> str:
    """
    Транскрибация файла по локальному пути (не Telegram File!).
    Если файл длинный — режем на куски по 25 минут и склеиваем текст.
    Используется backend'ом мини-приложения.
    """
    size_mb = os.path.getsize(path) / (1024 * 1024)

    chunk_paths = [path]
    cleanup = []

    if size_mb > GROQ_SIZE_LIMIT_MB:
        try:
            chunk_paths = split_audio_chunks(path, chunk_minutes=25)
            cleanup = chunk_paths
        except Exception as e:
            logger.warning(f"Не удалось порезать аудио на части ({e}), пробую сжать целиком.")
            compressed = compress_audio(path)
            chunk_paths = [compressed]
            if compressed != path:
                cleanup = [compressed]

    full_text_parts = []
    try:
        for p in chunk_paths:
            with open(p, 'rb') as f:
                transcription = groq_client.audio.transcriptions.create(
                    file=(os.path.basename(p), f.read()),
                    model='whisper-large-v3-turbo',
                    language='uk',
                )
            full_text_parts.append(transcription.text)
    finally:
        for p in cleanup:
            try:
                os.remove(p)
            except OSError:
                pass

    return "\n\n".join(full_text_parts)


def generate_document(doc_type: str, transcription: str) -> tuple[str, str]:
    """
    Возвращает (имя_документа, очищенный_текст).
    Для 'transcript' — без обращения к Gemini, просто чистый текст.
    """
    doc_name = DOC_NAMES[doc_type]
    if doc_type == 'transcript':
        return doc_name, transcription.strip()

    full_prompt = f"{PROMPTS[doc_type]}\n\nОсь транскрипція запису:\n\n{transcription}"
    response = gemini_model.generate_content(full_prompt)
    return doc_name, clean_markdown(response.text)
