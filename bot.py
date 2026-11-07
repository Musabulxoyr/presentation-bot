"""
📊 AI Slayd Bot (TEKIN versiya)
Mavzu -> slayd soni -> til -> rang -> maket -> rasm (ha/yo'q) -> tayyor .pptx
Har 10 ta taqdimotdan ~3 tasidan keyin foydalanuvchidan baho/fikr so'raydi va adminga yuboradi.

Hammasi bepul:
  • AI: Google Gemini (bepul kalit: aistudio.google.com/apikey)
  • Rasm: Pexels (bepul kalit, ixtiyoriy) yoki Wikimedia Commons (kalitsiz)

O'rnatish:  pip install aiogram python-pptx aiohttp pillow
"""
import asyncio
import html
import io
import json
import logging
import os
import random
import re
import sqlite3
import tempfile
import time
from contextlib import closing

import aiohttp
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (CallbackQuery, FSInputFile, KeyboardButton, LabeledPrice,
                           Message, PreCheckoutQuery, ReplyKeyboardMarkup)
from aiogram.utils.keyboard import InlineKeyboardBuilder
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# ======================= SOZLAMALAR (shu yerni to'ldiring) =======================
BOT_TOKEN = os.getenv("BOT_TOKEN") or "BOT_TOKEN_NI_SHU_YERGA"
GEMINI_KEY = os.getenv("GEMINI_API_KEY") or "GEMINI_KALITNI_SHU_YERGA"
PEXELS_KEY = os.getenv("PEXELS_API_KEY") or ""  # ixtiyoriy; bo'sh bo'lsa Wikimedia ishlatiladi
ADMIN_ID = int(os.getenv("ADMIN_ID") or 0)      # sizning Telegram ID raqamingiz (@userinfobot dan oling)
CARD_NUMBER = os.getenv("CARD_NUMBER") or ""      # karta raqami, masalan 8600 1234 5678 9012
CARD_OWNER = os.getenv("CARD_OWNER") or ""        # karta egasi ismi
DB_PATH = os.getenv("DB_PATH") or "bot.db"        # Railway'da Volume ulasangiz: /data/bot.db

FREE_LIMIT = 3          # har bir yangi foydalanuvchiga bepul taqdimotlar soni
FEEDBACK_RATE = 0.3     # har 10 ta taqdimotdan ~3 tasidan keyin fikr so'raladi (sinash uchun 1.0 qiling)
# Balansni to'ldirish paketlari: (taqdimotlar soni, Stars narxi, so'm narxi)
PACKS = [
    (1, 15, 3000),
    (5, 60, 12000),
    (10, 100, 20000),
]

DEBUG = True  # xato sababini chatda ko'rsatadi; hammasi ishlab ketsa False qiling
# =================================================================================

MODELS = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-flash-latest", "gemini-2.5-flash"]  # biri ishlamasa keyingisi
_discovered: list | None = None
UA = {"User-Agent": "AISlaydBot/1.0 (telegram bot)"}

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("slide-bot")

SLIDE_COUNTS = [5, 8, 10, 12, 15, 18, 20]

LANGS = {
    "uz": ("🇺🇿 O'zbekcha", "Uzbek (Latin script)"),
    "ru": ("🇷🇺 Русский", "Russian"),
    "en": ("🇬🇧 English", "English"),
}
THANKS = {
    "uz": ("E'tiboringiz uchun rahmat!", "Savollaringiz bormi?"),
    "ru": ("Спасибо за внимание!", "Есть вопросы?"),
    "en": ("Thank you for your attention!", "Any questions?"),
}

# nomi, to'q rang, aksent, och fon, matn rangi
THEMES = {
    "blue":   ("🔵 Ko'k",       "0B3D91", "2F80ED", "F4F8FF", "1F2937"),
    "green":  ("🟢 Yashil",     "064E3B", "10B981", "F0FDF4", "1F2937"),
    "purple": ("🟣 Binafsha",   "3B0764", "A855F7", "FAF5FF", "1F2937"),
    "orange": ("🟠 To'q sariq", "7C2D12", "F97316", "FFF7ED", "1F2937"),
    "red":    ("🔴 Qizil",      "7F1D1D", "EF4444", "FEF2F2", "1F2937"),
    "dark":   ("⚫️ Qora",       "0F172A", "38BDF8", "1E293B", "F1F5F9"),
}

# Slayd maketlari (joylashuv). Rangni THEMES beradi, joylashuvni STYLES.
STYLES = {
    "classic":  "📐 Klassik",
    "side":     "📑 Yon panel",
    "minimal":  "✨ Minimal",
    "cards":    "🗂 Kartochkali",
    "split":    "🖼 Ikki ustunli",
    "bignum":   "🔢 Katta raqam",
    "timeline": "🛤 Chiziqli",
    "gradient": "🌈 Gradient",
    "magazine": "📰 Jurnal",
    "random":   "🎲 Tasodifiy",
}
STYLE_KEYS = [k for k in STYLES if k != "random"]

router = Router()
busy_users: set[int] = set()


class Form(StatesGroup):
    topic = State()
    receipt = State()
    comment = State()


# ================================ BAZA (limit) ================================
def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def prepare_db_path():
    """/data ochilmasa (Volume ulanmagan bo'lsa) bot yiqilmasin: vaqtincha bot.db ishlatiladi."""
    global DB_PATH
    try:
        folder = os.path.dirname(DB_PATH)
        if folder:
            os.makedirs(folder, exist_ok=True)
        with open(DB_PATH, "a"):
            pass
    except OSError as e:
        log.warning("%s ochilmadi (%s). Vaqtincha bot.db ishlatiladi: Volume ulanganini tekshiring!",
                    DB_PATH, e)
        DB_PATH = "bot.db"


def init_db():
    prepare_db_path()
    with closing(db()) as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(
            user_id INTEGER PRIMARY KEY, name TEXT, credits INTEGER NOT NULL,
            made INTEGER NOT NULL DEFAULT 0, created INTEGER);
        CREATE TABLE IF NOT EXISTS payments(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, kind TEXT,
            amount INTEGER, status TEXT, receipt TEXT, ts INTEGER, credits INTEGER);
        CREATE TABLE IF NOT EXISTS feedback(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, name TEXT, topic TEXT,
            theme TEXT, style TEXT, score INTEGER, comment TEXT, ts INTEGER);
        """)
        c.commit()
        try:  # eski bazalar uchun
            c.execute("ALTER TABLE payments ADD COLUMN credits INTEGER")
            c.commit()
        except sqlite3.OperationalError:
            pass


def get_credits(uid: int, name: str = "") -> int:
    if uid == ADMIN_ID:
        return 10**6  # admin cheksiz
    with closing(db()) as c:
        c.execute("INSERT OR IGNORE INTO users(user_id,name,credits,created) VALUES(?,?,?,?)",
                  (uid, name, FREE_LIMIT, int(time.time())))
        c.commit()
        return c.execute("SELECT credits FROM users WHERE user_id=?", (uid,)).fetchone()[0]


def add_credits(uid: int, n: int):
    get_credits(uid)
    with closing(db()) as c:
        c.execute("UPDATE users SET credits=credits+? WHERE user_id=?", (n, uid))
        c.commit()


def consume(uid: int):
    if uid == ADMIN_ID:
        return
    with closing(db()) as c:
        c.execute("UPDATE users SET credits=MAX(credits-1,0), made=made+1 WHERE user_id=?", (uid,))
        c.commit()


def new_payment(uid, kind, amount, status, receipt="", credits=1) -> int:
    with closing(db()) as c:
        cur = c.execute("INSERT INTO payments(user_id,kind,amount,status,receipt,ts,credits) "
                        "VALUES(?,?,?,?,?,?,?)",
                        (uid, kind, amount, status, receipt, int(time.time()), credits))
        c.commit()
        return cur.lastrowid


def finish_payment(pid: int, status: str):
    """Faqat 'pending' to'lovni o'zgartiradi (ikki marta tasdiqlanmasligi uchun)."""
    with closing(db()) as c:
        cur = c.execute("UPDATE payments SET status=? WHERE id=? AND status='pending'", (status, pid))
        c.commit()
        if cur.rowcount == 0:
            return None
        return c.execute("SELECT * FROM payments WHERE id=?", (pid,)).fetchone()


def new_feedback(uid, name, topic, theme, style) -> int:
    with closing(db()) as c:
        cur = c.execute("INSERT INTO feedback(user_id,name,topic,theme,style,ts) VALUES(?,?,?,?,?,?)",
                        (uid, name, topic, theme, style, int(time.time())))
        c.commit()
        return cur.lastrowid


def set_score(fid: int, uid: int, score: int):
    """Bahoni faqat bir marta qabul qiladi."""
    with closing(db()) as c:
        cur = c.execute("UPDATE feedback SET score=? WHERE id=? AND user_id=? AND score IS NULL",
                        (score, fid, uid))
        c.commit()
        if cur.rowcount == 0:
            return None
        return c.execute("SELECT * FROM feedback WHERE id=?", (fid,)).fetchone()


def set_comment(fid: int, uid: int, comment: str):
    with closing(db()) as c:
        c.execute("UPDATE feedback SET comment=? WHERE id=? AND user_id=?", (comment, fid, uid))
        c.commit()
        return c.execute("SELECT * FROM feedback WHERE id=?", (fid,)).fetchone()


def som(n: int) -> str:
    return f"{n:,}".replace(",", " ")


# ================================ KLAVIATURALAR ================================
BTN_NEW = "📊 Yangi taqdimot"
BTN_BAL = "💎 Balans"
BTN_BUY = "➕ Balansni to'ldirish"
BTN_HELP = "ℹ️ Yordam"


def main_kb():
    """Doimiy pastki menyu."""
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=BTN_NEW), KeyboardButton(text=BTN_BAL)],
                  [KeyboardButton(text=BTN_BUY), KeyboardButton(text=BTN_HELP)]],
        resize_keyboard=True, is_persistent=True,
        input_field_placeholder="Mavzuni yozing yoki tugmani bosing")


def count_kb():
    kb = InlineKeyboardBuilder()
    for n in SLIDE_COUNTS:
        kb.button(text=f"📊 {n} slayd", callback_data=f"count:{n}")
    kb.adjust(2, 2, 2, 1)
    return kb.as_markup()


def lang_kb():
    kb = InlineKeyboardBuilder()
    for key, (name, _) in LANGS.items():
        kb.button(text=name, callback_data=f"lang:{key}")
    kb.adjust(1)
    return kb.as_markup()


def theme_kb():
    kb = InlineKeyboardBuilder()
    for key, (name, *_r) in THEMES.items():
        kb.button(text=name, callback_data=f"theme:{key}")
    kb.adjust(2)
    return kb.as_markup()


def style_kb():
    kb = InlineKeyboardBuilder()
    for key, name in STYLES.items():
        kb.button(text=name, callback_data=f"style:{key}")
    kb.adjust(2)
    return kb.as_markup()


def feedback_kb(fid: int):
    kb = InlineKeyboardBuilder()
    for n in range(1, 6):
        kb.button(text=f"{n}⭐️", callback_data=f"fb:{n}:{fid}")
    kb.adjust(5)
    return kb.as_markup()


def card_on() -> bool:
    return bool(CARD_NUMBER and ADMIN_ID)


def packs_kb():
    kb = InlineKeyboardBuilder()
    for i, (n, stars, s) in enumerate(PACKS):
        price = f"⭐️{stars} / {som(s)} so'm" if card_on() else f"⭐️{stars}"
        kb.button(text=f"💎 {n} ta taqdimot — {price}", callback_data=f"pack:{i}")
    kb.adjust(1)
    return kb.as_markup()


pay_kb = packs_kb  # eski nom


def images_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="🖼 Ha, rasm qo'shilsin", callback_data="img:yes")
    kb.button(text="📄 Yo'q, rasmsiz", callback_data="img:no")
    kb.adjust(1)
    return kb.as_markup()


# ==================================== AI ====================================
async def discover_models(session) -> list:
    """Google'dan hozir mavjud flash modellarni so'raydi (nomlar o'zgarib turadi)."""
    global _discovered
    if _discovered is not None:
        return _discovered
    out = []
    try:
        async with session.get(
            "https://generativelanguage.googleapis.com/v1beta/models",
            params={"pageSize": 200}, headers={"x-goog-api-key": GEMINI_KEY},
        ) as r:
            j = await r.json()
        bad = ("image", "tts", "live", "audio", "embedding", "robotics", "computer", "native")
        for m in j.get("models", []):
            name = m.get("name", "").split("/")[-1]
            if ("generateContent" in m.get("supportedGenerationMethods", [])
                    and "flash" in name and not any(x in name for x in bad)):
                out.append(name)
    except Exception as e:
        log.warning("Modellar ro'yxati olinmadi: %s", e)
    out.sort(reverse=True)
    _discovered = out
    log.info("Mavjud modellar: %s", out)
    return out


async def gemini(prompt: str) -> str:
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0.7},
    }
    errors = []
    timeout = aiohttp.ClientTimeout(total=150)
    async with aiohttp.ClientSession(timeout=timeout) as session:
        order = []
        for m in [os.getenv("AI_MODEL", "")] + MODELS + await discover_models(session):
            if m and m not in order:
                order.append(m)
        for model in order[:8]:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
            for attempt in range(2):
                try:
                    async with session.post(url, json=body,
                                            headers={"x-goog-api-key": GEMINI_KEY}) as r:
                        j = await r.json()
                        if r.status == 200:
                            parts = j["candidates"][0]["content"]["parts"]
                            txt = "".join(p.get("text", "") for p in parts)
                            if txt.strip():
                                return txt
                            errors.append(f"{model}: bo'sh javob")
                            break
                        msg = j.get("error", {}).get("message", str(j))
                        errors.append(f"{model} [{r.status}]: {msg[:120]}")
                        if r.status in (401, 403) or (r.status == 400 and "API key" in msg):
                            raise RuntimeError(errors[-1])  # kalit xato: boshqa model yordam bermaydi
                        if r.status in (400, 404):
                            break  # bu model yaroqsiz, keyingisiga o't
                except (aiohttp.ClientError, asyncio.TimeoutError) as e:
                    errors.append(f"{model}: {type(e).__name__}")
                await asyncio.sleep(3 * (attempt + 1))
    raise RuntimeError(" | ".join(errors[-3:]) or "noma'lum xato")


def extract_json(text: str) -> dict:
    text = re.sub(r"```(?:json)?", "", text)
    a, b = text.find("{"), text.rfind("}")
    if a == -1 or b == -1:
        raise ValueError("AI JSON qaytarmadi")
    return json.loads(text[a : b + 1])


async def generate_content(topic: str, total: int, lang: str) -> dict:
    n = total - 2  # titul va yakuniy slaydlar alohida
    prompt = f"""Topic: "{topic}"

Create content for a professional presentation. Write ALL text in {LANGS[lang][1]}.
Exactly {n} main slides.

Reply with ONLY this JSON, nothing else:
{{
  "title": "Presentation title",
  "subtitle": "Short subtitle",
  "slides": [
    {{"title": "Slide title",
      "bullets": ["point 1", "point 2", "point 3", "point 4"],
      "notes": "2-3 sentences the presenter can say aloud for this slide",
      "image_query": "2-3 ENGLISH words for a stock photo search"}}
  ]
}}

Rules:
- Each slide has 3-5 bullets, each 8-18 words, concrete and informative (facts, numbers, examples).
- Logical flow: introduction, main parts, examples/statistics, conclusion.
- Slide titles are short (max 6 words).
- image_query must be a concrete visual subject, always in English."""
    last = None
    for _ in range(2):
        try:
            data = extract_json(await gemini(prompt))
            if data.get("slides"):
                return data
            last = ValueError("Slaydlar bo'sh")
        except (ValueError, KeyError, json.JSONDecodeError) as e:
            last = e
    raise last


# =================================== RASMLAR ===================================
IMG_W, IMG_H = 4.3, 4.8  # dyuym


def crop_image(raw: bytes) -> bytes:
    img = Image.open(io.BytesIO(raw)).convert("RGB")
    target = IMG_W / IMG_H
    w, h = img.size
    if w / h > target:
        nw = int(h * target)
        img = img.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
    else:
        nh = int(w / target)
        img = img.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
    img.thumbnail((1000, 1100))
    out = io.BytesIO()
    img.save(out, "JPEG", quality=85)
    return out.getvalue()


async def url_pexels(session, q):
    async with session.get(
        "https://api.pexels.com/v1/search",
        params={"query": q, "per_page": 1, "orientation": "portrait"},
        headers={"Authorization": PEXELS_KEY},
    ) as r:
        photos = (await r.json()).get("photos") or []
    return photos[0]["src"]["large"] if photos else None


async def url_commons(session, q):
    params = {
        "action": "query", "format": "json", "generator": "search",
        "gsrsearch": f"filetype:bitmap {q}", "gsrnamespace": 6, "gsrlimit": 8,
        "prop": "imageinfo", "iiprop": "url|mime|size", "iiurlwidth": 1000,
    }
    async with session.get("https://commons.wikimedia.org/w/api.php", params=params) as r:
        pages = ((await r.json()).get("query") or {}).get("pages") or {}
    for p in sorted(pages.values(), key=lambda x: x.get("index", 99)):
        info = (p.get("imageinfo") or [{}])[0]
        if info.get("mime") == "image/jpeg" and info.get("width", 0) >= 700:
            return info.get("thumburl") or info.get("url")
    return None


async def fetch_image(session, query):
    if not query:
        return None
    try:
        url = None
        if PEXELS_KEY:
            url = await url_pexels(session, query)
        if not url:
            url = await url_commons(session, query)
        if not url:
            return None
        async with session.get(url) as r:
            if r.status != 200:
                return None
            raw = await r.read()
        return await asyncio.to_thread(crop_image, raw)
    except Exception as e:
        log.warning("Rasm olinmadi (%s): %s", query, e)
        return None


async def fetch_all_images(slides):
    timeout = aiohttp.ClientTimeout(total=25)
    async with aiohttp.ClientSession(timeout=timeout, headers=UA) as session:
        sem = asyncio.Semaphore(4)

        async def one(sl):
            async with sem:
                return await fetch_image(session, sl.get("image_query", ""))

        return await asyncio.gather(*[one(sl) for sl in slides])


# ==================================== PPTX ====================================
def rgb(h):
    return RGBColor.from_string(h)


def lum(h):
    """Rang yorqinligi: 0 (qora) ... 1 (oq)."""
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


def mix(a, b, t):
    """a va b ranglarini aralashtiradi (t=0 -> a, t=1 -> b)."""
    ca = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def set_grad(slide, bottom, top):
    """Gradient fon (pastdan tepaga). Ishlamasa oddiy rangga qaytadi."""
    try:
        f = slide.background.fill
        f.gradient()
        f.gradient_angle = 90
        f.gradient_stops[0].color.rgb = rgb(bottom)
        f.gradient_stops[1].color.rgb = rgb(top)
    except Exception:
        set_bg(slide, top)


def rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(color)
    s.line.fill.background()
    return s


def text(slide, x, y, w, h, t, size, color, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = t
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = "Calibri"
    r.font.color.rgb = rgb(color)
    return tb


def badge(slide, x, y, d, label, color, size=16, fg="FFFFFF"):
    """Ichida raqami bor aylana."""
    b = rect(slide, x, y, d, d, color, MSO_SHAPE.OVAL)
    tf = b.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(label)
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.name = "Calibri"
    r.font.color.rgb = rgb(fg)
    return b


def fit_size(bullets, w, h, start):
    """Matn berilgan joyga sig'adigan eng katta shriftni tanlaydi (taxminan)."""
    for size in range(start, 11, -1):
        cpl = max(1, int((w - 0.25) * 72 / (size * 0.52)))  # bir qatordagi belgilar soni
        if all(-(-len(b) // cpl) * size * 1.25 / 72 <= h for b in bullets):
            return size
    return 12


def bullets_plain(s, bullets, dot_x, text_x, text_w, top, bottom, start, color, accent):
    n = len(bullets)
    step = (bottom - top) / n
    size = fit_size(bullets, text_w, step - 0.05, start)
    for j, b in enumerate(bullets):
        y = top + step * j
        rect(s, Inches(dot_x), Inches(y + 0.13), Inches(0.2), Inches(0.2), accent, MSO_SHAPE.OVAL)
        text(s, Inches(text_x), Inches(y), Inches(text_w), Inches(step), b, size, color)


def bullets_cards(s, bullets, x, w, top, bottom, cols, start, color, accent, card):
    n = len(bullets)
    rows = -(-n // cols)
    gap = 0.25
    cw = (w - gap * (cols - 1)) / cols
    ch = ((bottom - top) - gap * (rows - 1)) / rows
    size = fit_size(bullets, cw - 0.4, ch - 0.1, start)
    for j, b in enumerate(bullets):
        r_, c_ = divmod(j, cols)
        cx = x + c_ * (cw + gap)
        cy = top + r_ * (ch + gap)
        rect(s, Inches(cx), Inches(cy), Inches(cw), Inches(ch), card)
        rect(s, Inches(cx), Inches(cy), Inches(0.1), Inches(ch), accent)
        text(s, Inches(cx + 0.25), Inches(cy), Inches(cw - 0.4), Inches(ch), b, size, color,
             anchor=MSO_ANCHOR.MIDDLE)


def bullets_timeline(s, bullets, line_x, text_x, text_w, top, bottom, start, color, accent):
    n = len(bullets)
    step = (bottom - top) / n
    size = fit_size(bullets, text_w, step - 0.05, start)
    if n > 1:
        rect(s, Inches(line_x - 0.03), Inches(top + 0.25), Inches(0.06),
             Inches(step * (n - 1)), accent)
    for j, b in enumerate(bullets):
        y = top + step * j
        badge(s, Inches(line_x - 0.25), Inches(y), Inches(0.5), j + 1, accent, 16)
        text(s, Inches(text_x), Inches(y), Inches(text_w), Inches(step), b, size, color)


def bullets_numbered(s, bullets, x, text_x, text_w, top, bottom, start, color, accent):
    n = len(bullets)
    step = (bottom - top) / n
    size = fit_size(bullets, text_w, step - 0.05, start)
    for j, b in enumerate(bullets):
        y = top + step * j
        text(s, Inches(x), Inches(y), Inches(0.9), Inches(0.6), f"{j + 1:02d}", 20, accent, True)
        text(s, Inches(text_x), Inches(y), Inches(text_w), Inches(step), b, size, color)


def put_image(s, img, x, y, w, h, frame, accent):
    if frame:
        rect(s, Inches(x), Inches(y), Inches(w + 0.2), Inches(h + 0.2), accent)
        x, y = x + 0.1, y + 0.1
    s.shapes.add_picture(io.BytesIO(img), Inches(x), Inches(y), Inches(w), Inches(h))


def build_pptx(data, theme_key, style, lang, path, images=None):
    _, dark, accent, light, txt = THEMES[theme_key]
    if style not in STYLE_KEYS:
        style = "classic"
    on_dark = "FFFFFF"
    soft = "D1D5DB"
    bg_dark = lum(light) < 0.4                       # fon to'q bo'lsa (Qora tema)
    head_c = "FFFFFF" if bg_dark else dark           # och fondagi sarlavha rangi
    card = mix(light, "FFFFFF", 0.1) if bg_dark else "FFFFFF"
    g_bottom, g_top = mix(dark, accent, 0.4), dark   # gradient ranglari
    g_txt = "F1F5F9"

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    W, H = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]
    slides = data["slides"]
    total = len(slides) + 2
    title_t, sub_t = data.get("title", ""), data.get("subtitle", "")

    # ------------------------------ TITUL ------------------------------
    s = prs.slides.add_slide(blank)
    if style == "classic":
        set_bg(s, dark)
        rect(s, 0, 0, Inches(0.35), H, accent)
        rect(s, W - Inches(3.2), Inches(-1.2), Inches(4.2), Inches(4.2), accent, MSO_SHAPE.OVAL)
        rect(s, W - Inches(2.2), Inches(5.3), Inches(3.0), Inches(3.0), accent, MSO_SHAPE.OVAL)
        rect(s, Inches(1.2), Inches(4.4), Inches(2.2), Inches(0.08), accent)
        text(s, Inches(1.2), Inches(1.7), Inches(9.0), Inches(2.6), title_t, 46, on_dark, True,
             anchor=MSO_ANCHOR.BOTTOM)
        text(s, Inches(1.2), Inches(4.65), Inches(9.0), Inches(1.2), sub_t, 22, soft)
    elif style == "side":
        set_bg(s, dark)
        rect(s, 0, 0, Inches(4.5), H, accent)
        rect(s, Inches(-1.0), Inches(4.5), Inches(3.5), Inches(3.5), dark, MSO_SHAPE.OVAL)
        rect(s, Inches(5.2), Inches(4.1), Inches(2.2), Inches(0.08), accent)
        text(s, Inches(5.2), Inches(1.4), Inches(7.4), Inches(2.6), title_t, 42, on_dark, True,
             anchor=MSO_ANCHOR.BOTTOM)
        text(s, Inches(5.2), Inches(4.4), Inches(7.4), Inches(1.5), sub_t, 22, soft)
    elif style == "minimal":
        set_bg(s, light)
        rect(s, Inches(1.2), Inches(1.3), Inches(1.4), Inches(0.12), accent)
        text(s, Inches(1.2), Inches(1.7), Inches(10.8), Inches(2.8), title_t, 50, head_c, True,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(1.2), Inches(4.7), Inches(10.8), Inches(1.2), sub_t, 22, txt)
        rect(s, 0, H - Inches(0.5), W, Inches(0.5), dark)
    elif style == "cards":
        set_bg(s, light)
        rect(s, Inches(1.0), Inches(1.0), W - Inches(2.0), Inches(5.3), card)
        rect(s, Inches(1.0), Inches(1.0), Inches(0.2), Inches(5.3), accent)
        text(s, Inches(1.7), Inches(1.3), Inches(9.9), Inches(2.7), title_t, 44, head_c, True,
             anchor=MSO_ANCHOR.BOTTOM)
        rect(s, Inches(1.7), Inches(4.25), Inches(2.0), Inches(0.08), accent)
        text(s, Inches(1.7), Inches(4.5), Inches(9.9), Inches(1.4), sub_t, 22, txt)
        rect(s, 0, H - Inches(0.5), W, Inches(0.5), dark)
    elif style == "split":
        set_bg(s, dark)
        rect(s, W - Inches(5.3), 0, Inches(5.3), H, accent)
        rect(s, W - Inches(3.8), Inches(1.2), Inches(3.0), Inches(3.0), dark, MSO_SHAPE.OVAL)
        rect(s, W - Inches(2.8), Inches(4.6), Inches(1.6), Inches(1.6),
             mix(accent, "FFFFFF", 0.5), MSO_SHAPE.OVAL)
        rect(s, Inches(0.9), Inches(4.5), Inches(2.2), Inches(0.08), accent)
        text(s, Inches(0.9), Inches(1.5), Inches(6.8), Inches(2.7), title_t, 42, on_dark, True,
             anchor=MSO_ANCHOR.BOTTOM)
        text(s, Inches(0.9), Inches(4.75), Inches(6.8), Inches(1.5), sub_t, 22, soft)
    elif style == "bignum":
        set_bg(s, light)
        rect(s, W - Inches(5.2), Inches(0.8), Inches(6.4), Inches(6.4), accent, MSO_SHAPE.OVAL)
        rect(s, W - Inches(4.0), Inches(2.0), Inches(4.0), Inches(4.0),
             mix(accent, light, 0.55), MSO_SHAPE.OVAL)
        rect(s, W - Inches(2.9), Inches(3.1), Inches(1.8), Inches(1.8), dark, MSO_SHAPE.OVAL)
        rect(s, Inches(1.0), Inches(4.7), Inches(2.2), Inches(0.1), accent)
        text(s, Inches(1.0), Inches(1.7), Inches(7.2), Inches(2.8), title_t, 50, head_c, True,
             anchor=MSO_ANCHOR.BOTTOM)
        text(s, Inches(1.0), Inches(4.95), Inches(7.2), Inches(1.4), sub_t, 22, txt)
    elif style == "timeline":
        set_bg(s, light)
        rect(s, Inches(1.4), Inches(0.8), Inches(0.06), H - Inches(1.6), accent)
        badge(s, Inches(1.28), Inches(0.65), Inches(0.3), "", accent)
        badge(s, Inches(1.28), Inches(5.5), Inches(0.3), "", accent)
        badge(s, Inches(1.08), Inches(2.95), Inches(0.7), "", dark)
        text(s, Inches(2.4), Inches(1.5), Inches(9.5), Inches(2.8), title_t, 44, head_c, True,
             anchor=MSO_ANCHOR.BOTTOM)
        rect(s, Inches(2.4), Inches(4.5), Inches(2.0), Inches(0.08), accent)
        text(s, Inches(2.4), Inches(4.75), Inches(9.5), Inches(1.4), sub_t, 22, txt)
    elif style == "gradient":
        set_grad(s, g_bottom, g_top)
        rect(s, W - Inches(4.0), Inches(3.8), Inches(6.0), Inches(6.0),
             mix(dark, accent, 0.3), MSO_SHAPE.OVAL)
        rect(s, 0, 0, W, Inches(0.25), accent)
        rect(s, Inches(1.2), Inches(4.4), Inches(2.2), Inches(0.08), accent)
        text(s, Inches(1.2), Inches(1.7), Inches(10.0), Inches(2.6), title_t, 48, on_dark, True,
             anchor=MSO_ANCHOR.BOTTOM)
        text(s, Inches(1.2), Inches(4.65), Inches(10.0), Inches(1.3), sub_t, 22, soft)
    else:  # magazine
        set_bg(s, "FFFFFF")
        rect(s, Inches(0.8), Inches(4.7), Inches(11.7), Inches(0.05), "111111")
        rect(s, Inches(0.8), Inches(4.65), Inches(2.2), Inches(0.15), accent)
        text(s, Inches(0.8), Inches(0.8), Inches(11.7), Inches(3.7), title_t, 58, "111111", True,
             anchor=MSO_ANCHOR.BOTTOM)
        text(s, Inches(0.8), Inches(5.0), Inches(11.7), Inches(1.3), sub_t, 24, "4B5563")
        rect(s, 0, H - Inches(0.35), W, Inches(0.35), "111111")

    # --------------------------- ASOSIY SLAYDLAR ---------------------------
    for i, sl in enumerate(slides, start=2):
        s = prs.slides.add_slide(blank)
        title = sl.get("title", "")
        num = i - 1
        img = images[i - 2] if images and images[i - 2] else None
        bullets = [b for b in sl.get("bullets", []) if b][:6] or [""]
        n = len(bullets)
        start = 20 if img else 22
        body = txt
        ix, iy, iw, ih, frame = 8.4, 1.8, IMG_W, IMG_H, True   # rasm joyi
        pn_x, pn_align = 11.6, PP_ALIGN.RIGHT                  # sahifa raqami joyi

        if style == "classic":
            set_bg(s, light)
            rect(s, 0, 0, W, Inches(1.35), dark)
            rect(s, 0, Inches(1.35), W, Inches(0.07), accent)
            badge(s, Inches(0.55), Inches(0.35), Inches(0.65), num, accent, 18)
            text(s, Inches(1.45), Inches(0.1), Inches(11.4), Inches(1.15), title, 30, on_dark,
                 True, anchor=MSO_ANCHOR.MIDDLE)
            kind = "plain"
            P = dict(dot_x=0.8, text_x=1.2, text_w=6.9 if img else 11.5, top=1.85, bottom=6.7)

        elif style == "side":
            set_bg(s, light)
            rect(s, 0, 0, Inches(3.1), H, dark)
            rect(s, Inches(3.1), 0, Inches(0.07), H, accent)
            text(s, Inches(0.35), Inches(0.5), Inches(2.5), Inches(1.2), str(num), 60, accent, True)
            text(s, Inches(0.35), Inches(1.9), Inches(2.5), Inches(4.5), title, 26, on_dark, True)
            kind = "plain"
            P = dict(dot_x=3.6, text_x=4.0, text_w=4.1 if img else 8.9, top=1.0, bottom=6.7)

        elif style == "minimal":
            set_bg(s, light)
            text(s, Inches(0.8), Inches(0.3), Inches(1.5), Inches(0.6), f"{num:02d}", 22, accent, True)
            text(s, Inches(0.8), Inches(0.8), Inches(11.7), Inches(1.0), title, 32, head_c, True,
                 anchor=MSO_ANCHOR.MIDDLE)
            rect(s, Inches(0.8), Inches(1.85), Inches(1.2), Inches(0.07), accent)
            kind = "plain"
            P = dict(dot_x=0.8, text_x=1.2, text_w=6.9 if img else 11.5, top=2.2, bottom=6.7)
            iy = 1.95

        elif style == "cards":
            set_bg(s, light)
            badge(s, Inches(0.8), Inches(0.45), Inches(0.7), num, accent, 20)
            text(s, Inches(1.75), Inches(0.3), Inches(11.0), Inches(1.0), title, 30, head_c, True,
                 anchor=MSO_ANCHOR.MIDDLE)
            rect(s, Inches(0.8), Inches(1.45), Inches(1.2), Inches(0.07), accent)
            kind = "cards"
            P = dict(x=0.8, w=7.3 if img else 11.7, top=1.85, bottom=6.7,
                     cols=1 if (img or n <= 2) else 2)

        elif style == "split":
            set_bg(s, light)
            rect(s, 0, 0, Inches(5.3), H, dark)
            rect(s, Inches(5.3), 0, Inches(0.07), H, accent)
            if img:
                badge(s, Inches(0.5), Inches(0.45), Inches(0.65), num, accent, 18)
                text(s, Inches(1.35), Inches(0.3), Inches(3.8), Inches(1.6), title, 24, on_dark,
                     True, anchor=MSO_ANCHOR.MIDDLE)
                ix, iy = 0.45, 2.1
            else:
                text(s, Inches(0.5), Inches(0.8), Inches(4.3), Inches(2.2), f"{num:02d}", 100,
                     accent, True, anchor=MSO_ANCHOR.MIDDLE)
                rect(s, Inches(0.55), Inches(3.2), Inches(1.6), Inches(0.08), accent)
                text(s, Inches(0.5), Inches(3.45), Inches(4.4), Inches(3.2), title, 28, on_dark, True)
            kind = "plain"
            P = dict(dot_x=5.9, text_x=6.3, text_w=6.5, top=1.0, bottom=6.7)
            start = min(start, 20)

        elif style == "bignum":
            set_bg(s, light)
            text(s, Inches(0.7), Inches(0.3), Inches(3.6), Inches(2.3), f"{num:02d}", 110, accent,
                 True, anchor=MSO_ANCHOR.MIDDLE)
            text(s, Inches(4.3), Inches(0.5), Inches(4.0 if img else 8.4), Inches(2.0), title,
                 26 if img else 34, head_c, True, anchor=MSO_ANCHOR.MIDDLE)
            rect(s, Inches(0.8), Inches(2.75), Inches(7.3 if img else 11.7), Inches(0.04), accent)
            kind = "plain"
            P = dict(dot_x=0.9, text_x=1.3, text_w=6.8 if img else 11.4, top=3.1, bottom=6.7)

        elif style == "timeline":
            set_bg(s, light)
            text(s, Inches(0.8), Inches(0.3), Inches(10.4), Inches(1.1), title, 30, head_c, True,
                 anchor=MSO_ANCHOR.MIDDLE)
            text(s, Inches(11.3), Inches(0.3), Inches(1.7), Inches(1.1), f"{num:02d}", 36, accent,
                 True, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
            rect(s, Inches(0.8), Inches(1.45), Inches(1.2), Inches(0.07), accent)
            kind = "timeline"
            P = dict(line_x=1.4, text_x=2.1, text_w=6.0 if img else 10.6, top=2.0, bottom=6.7)

        elif style == "gradient":
            set_grad(s, g_bottom, g_top)
            rect(s, W - Inches(3.0), Inches(5.2), Inches(4.0), Inches(4.0),
                 mix(dark, accent, 0.3), MSO_SHAPE.OVAL)
            rect(s, 0, 0, W, Inches(0.25), accent)
            badge(s, Inches(0.7), Inches(0.55), Inches(0.75), num, accent, 20)
            text(s, Inches(1.7), Inches(0.4), Inches(11.0), Inches(1.1), title, 30, on_dark, True,
                 anchor=MSO_ANCHOR.MIDDLE)
            kind = "plain"
            body = g_txt
            P = dict(dot_x=0.8, text_x=1.2, text_w=6.9 if img else 11.5, top=1.9, bottom=6.7)

        else:  # magazine
            set_bg(s, "FFFFFF")
            text(s, Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.9), title, 40, "111111", True,
                 anchor=MSO_ANCHOR.BOTTOM)
            rect(s, Inches(0.8), Inches(2.45), Inches(11.7), Inches(0.04), "111111")
            rect(s, Inches(0.8), Inches(2.4), Inches(1.6), Inches(0.14), accent)
            kind = "numbered"
            body = "1F2937"
            P = dict(x=0.8, text_x=1.7, text_w=6.6 if img else 10.8, top=2.85, bottom=6.8)
            ix, iy, iw, frame = 8.9, 2.75, 3.9, False
            ih = iw * IMG_H / IMG_W
            pn_x, pn_align = 0.8, PP_ALIGN.LEFT

        if img:
            put_image(s, img, ix, iy, iw, ih, frame, accent)

        common = dict(start=start, color=body, accent=accent)
        if kind == "plain":
            bullets_plain(s, bullets, **P, **common)
        elif kind == "cards":
            bullets_cards(s, bullets, card=card, **P, **common)
        elif kind == "timeline":
            bullets_timeline(s, bullets, **P, **common)
        else:
            bullets_numbered(s, bullets, **P, **common)

        text(s, Inches(pn_x), Inches(6.95), Inches(1.4), Inches(0.4), f"{i} / {total}", 12,
             "9CA3AF", align=pn_align)
        if sl.get("notes"):
            s.notes_slide.notes_text_frame.text = sl["notes"]

    # ------------------------------ YAKUNIY ------------------------------
    s = prs.slides.add_slide(blank)
    t1, t2 = THANKS[lang]
    if style in ("classic", "side", "split"):
        set_bg(s, dark)
        if style == "side":
            rect(s, 0, 0, Inches(0.35), H, accent)
        else:
            rect(s, W - Inches(0.35), 0, Inches(0.35), H, accent)
        rect(s, Inches(-1.5), Inches(4.8), Inches(4.0), Inches(4.0), accent, MSO_SHAPE.OVAL)
        c1, c2 = on_dark, soft
    elif style == "gradient":
        set_grad(s, g_bottom, g_top)
        rect(s, 0, 0, W, Inches(0.25), accent)
        rect(s, Inches(-1.5), Inches(4.8), Inches(4.0), Inches(4.0),
             mix(dark, accent, 0.3), MSO_SHAPE.OVAL)
        c1, c2 = on_dark, soft
    elif style == "magazine":
        set_bg(s, "FFFFFF")
        rect(s, (W - Inches(2.2)) // 2, Inches(4.1), Inches(2.2), Inches(0.1), accent)
        rect(s, 0, H - Inches(0.35), W, Inches(0.35), "111111")
        c1, c2 = "111111", "4B5563"
    else:  # minimal, cards, bignum, timeline
        set_bg(s, light)
        rect(s, (W - Inches(2.2)) // 2, Inches(4.1), Inches(2.2), Inches(0.08), accent)
        rect(s, 0, H - Inches(0.5), W, Inches(0.5), dark)
        c1, c2 = head_c, txt
    text(s, 0, Inches(2.5), W, Inches(1.6), t1, 44, c1, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, 0, Inches(4.3), W, Inches(0.8), t2, 22, c2, align=PP_ALIGN.CENTER)

    prs.save(path)


# =================================== HANDLERS ===================================
async def ask_topic(m: Message, state: FSMContext):
    await state.clear()
    await state.set_state(Form.topic)
    await m.answer(
        "👋 Salom! Men sizga professional taqdimot tayyorlab beraman.\n\n"
        "✍️ Taqdimot mavzusini yozing, masalan:\n"
        "• <i>Global isish</i>\n• <i>Sun'iy intellekt kelajagi</i>\n• <i>Amir Temur</i>",
        parse_mode="HTML", reply_markup=main_kb(),
    )


@router.message(CommandStart())
@router.message(Command("new"))
async def cmd_start(m: Message, state: FSMContext):
    await ask_topic(m, state)


@router.message(Command("cancel"))
async def cmd_cancel(m: Message, state: FSMContext):
    await state.clear()
    busy_users.discard(m.from_user.id)
    await m.answer("❌ Bekor qilindi. Yangi taqdimot uchun /new ni bosing.", reply_markup=main_kb())


@router.message(Command("skip"))
async def cmd_skip(m: Message, state: FSMContext):
    await state.clear()
    await m.answer("👍 Yaxshi. Yangi taqdimot uchun /new ni bosing.", reply_markup=main_kb())


@router.message(Command("help"))
async def cmd_help(m: Message):
    await m.answer(
        "ℹ️ <b>Qanday ishlaydi?</b>\n"
        "1️⃣ Mavzuni yozasiz\n2️⃣ Slayd sonini tanlaysiz\n3️⃣ Tilni tanlaysiz\n"
        "4️⃣ Rangni tanlaysiz\n5️⃣ Maketni tanlaysiz\n6️⃣ Rasm kerakmi — tanlaysiz\n"
        "7️⃣ Tayyor .pptx faylni olasiz\n\n"
        "/new — yangi taqdimot\n/cancel — bekor qilish",
        parse_mode="HTML",
    )


@router.message(F.text == BTN_NEW)
async def btn_new(m: Message, state: FSMContext):
    await ask_topic(m, state)


@router.message(F.text == BTN_BAL)
async def btn_balance(m: Message):
    await cmd_balance(m)


@router.message(F.text == BTN_BUY)
async def btn_buy(m: Message):
    await m.answer("➕ <b>Balansni to'ldirish</b>\n\nPaketni tanlang:",
                   parse_mode="HTML", reply_markup=packs_kb())


@router.message(F.text == BTN_HELP)
async def btn_help(m: Message):
    await cmd_help(m)


async def show_paywall(m: Message):
    await m.answer(
        "😔 <b>Bepul limitingiz tugadi.</b>\n\n"
        "➕ Davom etish uchun balansni to'ldiring, paketni tanlang:",
        parse_mode="HTML", reply_markup=packs_kb())


async def take_topic(m: Message, state: FSMContext):
    if get_credits(m.from_user.id, m.from_user.full_name) <= 0:
        await state.clear()
        await show_paywall(m)
        return
    topic = (m.text or "").strip()
    if len(topic) < 3 or len(topic) > 200:
        await m.answer("⚠️ Mavzu 3 dan 200 belgigacha bo'lishi kerak. Qayta yozing.")
        return
    await state.update_data(topic=topic)
    await m.answer(f"📝 Mavzu: <b>{html.escape(topic)}</b>\n\n📊 Nechta slayd kerak?",
                   parse_mode="HTML", reply_markup=count_kb())


@router.message(Form.topic, F.text, ~F.text.startswith("/"))
async def got_topic(m: Message, state: FSMContext):
    await take_topic(m, state)


@router.callback_query(F.data.startswith("count:"))
async def got_count(c: CallbackQuery, state: FSMContext):
    if "topic" not in await state.get_data():
        await c.answer("Avval mavzuni yozing: /new", show_alert=True)
        return
    n = int(c.data.split(":")[1])
    await state.update_data(count=n)
    await c.message.edit_text(f"📊 Slaydlar soni: <b>{n}</b>\n\n🌐 Taqdimot tilini tanlang:",
                              parse_mode="HTML", reply_markup=lang_kb())
    await c.answer()


@router.callback_query(F.data.startswith("lang:"))
async def got_lang(c: CallbackQuery, state: FSMContext):
    if "count" not in await state.get_data():
        await c.answer("Avval /new ni bosing", show_alert=True)
        return
    lang = c.data.split(":")[1]
    await state.update_data(lang=lang)
    await c.message.edit_text(f"🌐 Til: <b>{LANGS[lang][0]}</b>\n\n🎨 Rangni tanlang:",
                              parse_mode="HTML", reply_markup=theme_kb())
    await c.answer()


@router.callback_query(F.data.startswith("theme:"))
async def got_theme(c: CallbackQuery, state: FSMContext):
    if "lang" not in await state.get_data():
        await c.answer("Avval /new ni bosing", show_alert=True)
        return
    theme = c.data.split(":")[1]
    await state.update_data(theme=theme)
    await c.message.edit_text(
        f"🎨 Rang: <b>{THEMES[theme][0]}</b>\n\n📐 Slayd maketini tanlang:",
        parse_mode="HTML", reply_markup=style_kb())
    await c.answer()


@router.callback_query(F.data.startswith("style:"))
async def got_style(c: CallbackQuery, state: FSMContext):
    if "theme" not in await state.get_data():
        await c.answer("Avval /new ni bosing", show_alert=True)
        return
    style = c.data.split(":")[1]
    if style not in STYLES:
        await c.answer("Noma'lum maket", show_alert=True)
        return
    await state.update_data(style=style)
    await c.message.edit_text(
        f"📐 Maket: <b>{STYLES[style]}</b>\n\n🖼 Slaydlarga rasm qo'shilsinmi?",
        parse_mode="HTML", reply_markup=images_kb())
    await c.answer()


@router.callback_query(F.data.startswith("img:"))
async def got_images(c: CallbackQuery, state: FSMContext):
    uid = c.from_user.id
    d = await state.get_data()
    if not all(k in d for k in ("topic", "count", "lang", "theme", "style")):
        await c.answer("Avval /new ni bosing", show_alert=True)
        return
    if uid in busy_users:
        await c.answer("⏳ Oldingi taqdimot tayyorlanmoqda, kuting...", show_alert=True)
        return

    if get_credits(uid, c.from_user.full_name) <= 0:
        await c.answer()
        await state.clear()
        await show_paywall(c.message)
        return

    want_images = c.data.endswith("yes")
    topic, count, lang, theme, style = d["topic"], d["count"], d["lang"], d["theme"], d["style"]
    if style == "random":
        style = random.choice(STYLE_KEYS)
    busy_users.add(uid)
    await c.answer()
    status = await c.message.edit_text(
        f"⏳ <b>{html.escape(topic)}</b> bo'yicha {count} slaydli taqdimot tayyorlanmoqda...\n"
        "Bu 20-90 soniya olishi mumkin.", parse_mode="HTML")

    path = None
    try:
        content = await generate_content(topic, count, lang)

        images = None
        if want_images:
            await status.edit_text("🖼 Rasmlar qidirilmoqda...")
            images = await fetch_all_images(content["slides"])
        found = sum(1 for x in (images or []) if x)

        await status.edit_text("🛠 Fayl yig'ilmoqda...")
        fd, path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)
        await asyncio.to_thread(build_pptx, content, theme, style, lang, path, images)

        safe = re.sub(r"[^\w\-]+", "_", topic)[:40] or "taqdimot"
        extra = f"\n🎨 {THEMES[theme][0]} · {STYLES[style]}"
        if want_images:
            extra += f"\n🖼 Rasmlar: {found} ta"
        if uid != ADMIN_ID:
            extra += f"\n💎 Qolgan limit: {max(get_credits(uid) - 1, 0)} ta"
        await c.message.answer_document(
            FSInputFile(path, filename=f"{safe}.pptx"),
            caption=(f"✅ Tayyor! 📊 {count} slayd\n🎯 Mavzu: {topic}{extra}\n"
                     "🗣 Har slaydda ma'ruzachi uchun izohlar (Notes) bor.\n\n"
                     "Yangi taqdimot: /new"),
            reply_markup=main_kb(),
        )
        consume(uid)  # limit faqat muvaffaqiyatli yuborilgandan keyin kamayadi

        # Har 10 ta taqdimotdan ~3 tasidan keyin fikr so'raymiz
        if random.random() < FEEDBACK_RATE:
            try:
                fid = new_feedback(uid, c.from_user.full_name, topic, theme, style)
                await c.message.answer(
                    "💬 Slaydlar dizayni sizga qanday ko'rinsa, baholang:",
                    reply_markup=feedback_kb(fid))
            except Exception as e:
                log.warning("Fikr so'rovi yuborilmadi: %s", e)
        await status.delete()
    except Exception as e:
        log.exception("Generatsiya xatosi")
        msg = "😔 Xatolik bo'ldi. Qayta urinib ko'ring: /new"
        if DEBUG:
            msg += f"\n\n🐞 {type(e).__name__}: {str(e)[:300]}"
        await status.edit_text(msg)
    finally:
        busy_users.discard(uid)
        await state.clear()
        if path and os.path.exists(path):
            os.remove(path)


# ================================== FIKR (baho) ==================================
def feedback_text(row, comment=None) -> str:
    theme_name = THEMES.get(row["theme"], ("?",))[0]
    style_name = STYLES.get(row["style"], row["style"])
    t = (f"💬 <b>Fikr #{row['id']}</b>\n"
         f"👤 {html.escape(row['name'] or '')} (id {row['user_id']})\n"
         f"⭐️ Baho: {row['score']}/5\n"
         f"🎯 Mavzu: {html.escape(row['topic'] or '')}\n"
         f"🎨 {theme_name} · {style_name}")
    if comment:
        t += f"\n\n✍️ {html.escape(comment)}"
    return t


async def tell_admin(bot: Bot, msg: str):
    if not ADMIN_ID:
        return
    try:
        await bot.send_message(ADMIN_ID, msg, parse_mode="HTML")
    except Exception as e:
        log.warning("Adminga xabar ketmadi: %s", e)


@router.callback_query(F.data.startswith("fb:"))
async def got_rating(c: CallbackQuery, state: FSMContext):
    _, score, fid = c.data.split(":")
    row = set_score(int(fid), c.from_user.id, int(score))
    if not row:
        await c.answer("Bu baho allaqachon qabul qilingan", show_alert=True)
        return
    await state.set_state(Form.comment)
    await state.update_data(fb=int(fid))
    await c.message.edit_text(
        f"🙏 Rahmat! Bahoyingiz: {'⭐️' * int(score)}\n\n"
        "✍️ Xohlasangiz, fikringizni yozing (nima yoqdi, nima yoqmadi).\n"
        "Yozmasangiz — /skip ni bosing.")
    await c.answer()
    await tell_admin(c.bot, feedback_text(row))


@router.message(Form.comment, F.text, ~F.text.startswith("/"))
async def got_comment(m: Message, state: FSMContext):
    fid = (await state.get_data()).get("fb")
    await state.clear()
    if fid:
        row = set_comment(fid, m.from_user.id, m.text[:1000])
        if row:
            await tell_admin(m.bot, feedback_text(row, comment=m.text[:1000]))
    await m.answer("✅ Fikringiz uchun rahmat! Yangi taqdimot: /new", reply_markup=main_kb())


# ================================= TO'LOVLAR =================================
@router.message(Command("balance"))
async def cmd_balance(m: Message):
    n = get_credits(m.from_user.id, m.from_user.full_name)
    if m.from_user.id == ADMIN_ID:
        await m.answer("👑 Siz adminsiz: cheksiz.")
    else:
        await m.answer(f"💎 Sizda <b>{n}</b> ta taqdimot qoldi.\n\n"
                       "➕ Balansni to'ldirish uchun paketni tanlang:",
                       parse_mode="HTML", reply_markup=packs_kb())


@router.callback_query(F.data == "buy:menu")
async def buy_menu(c: CallbackQuery):
    await c.message.edit_text("➕ <b>Balansni to'ldirish</b>\n\nPaketni tanlang:",
                              parse_mode="HTML", reply_markup=packs_kb())
    await c.answer()


@router.callback_query(F.data.startswith("pack:"))
async def choose_pack(c: CallbackQuery):
    i = int(c.data.split(":")[1])
    n, stars, s = PACKS[i]
    kb = InlineKeyboardBuilder()
    kb.button(text=f"⭐️ {stars} Stars bilan to'lash", callback_data=f"pay:stars:{i}")
    if card_on():
        kb.button(text=f"💳 {som(s)} so'm (karta)", callback_data=f"pay:card:{i}")
    kb.button(text="⬅️ Orqaga", callback_data="buy:menu")
    kb.adjust(1)
    await c.message.edit_text(f"📦 Paket: <b>{n} ta taqdimot</b>\n\nTo'lov usulini tanlang:",
                              parse_mode="HTML", reply_markup=kb.as_markup())
    await c.answer()


@router.callback_query(F.data.startswith("pay:stars:"))
async def pay_stars(c: CallbackQuery):
    i = int(c.data.split(":")[2])
    n, stars, _ = PACKS[i]
    await c.answer()
    await c.message.answer_invoice(
        title="📊 Taqdimot paketi",
        description=f"{n} ta taqdimot yaratish",
        payload=f"pack:{i}",
        currency="XTR",  # Telegram Stars
        prices=[LabeledPrice(label=f"{n} ta taqdimot", amount=stars)],
    )


@router.pre_checkout_query()
async def pre_checkout(q: PreCheckoutQuery):
    await q.answer(ok=True)


@router.message(F.successful_payment)
async def stars_paid(m: Message):
    sp = m.successful_payment
    uid = m.from_user.id
    try:
        n = PACKS[int(sp.invoice_payload.split(":")[1])][0]
    except Exception:
        n = 1
    new_payment(uid, "stars", sp.total_amount, "paid", sp.telegram_payment_charge_id, n)
    add_credits(uid, n)
    await m.answer(f"✅ To'lov qabul qilindi! +{n} ta taqdimot.\n"
                   f"💎 Jami: {get_credits(uid)} ta\n\n📊 Boshlash: /new",
                   reply_markup=main_kb())
    if ADMIN_ID:
        try:
            await m.bot.send_message(ADMIN_ID, f"⭐️ {sp.total_amount} Stars to'landi (+{n}): "
                                               f"{m.from_user.full_name} (id {uid})")
        except Exception:
            pass


@router.callback_query(F.data.startswith("pay:card:"))
async def pay_card(c: CallbackQuery, state: FSMContext):
    if not card_on():
        await c.answer("Karta orqali to'lov hozircha yo'q", show_alert=True)
        return
    i = int(c.data.split(":")[2])
    n, _, s = PACKS[i]
    await c.answer()
    await state.update_data(pack=i)
    await state.set_state(Form.receipt)
    await c.message.answer(
        f"💳 <b>{som(s)} so'm</b> ({n} ta taqdimot) ni shu kartaga o'tkazing:\n\n"
        f"<code>{CARD_NUMBER}</code>\n👤 {html.escape(CARD_OWNER)}\n\n"
        "O'tkazgach <b>chek rasmini (skrinshot)</b> shu yerga yuboring.\n"
        "🙏 Kechirasiz, karta to'lovini admin qo'lda tekshiradi, shuning uchun "
        "tasdiqlanishi biroz vaqt olishi mumkin.\n\nBekor qilish: /cancel",
        parse_mode="HTML")


@router.message(Form.receipt, F.photo)
async def got_receipt(m: Message, state: FSMContext):
    uid = m.from_user.id
    i = (await state.get_data()).get("pack", 0)
    n, _, s = PACKS[i]
    file_id = m.photo[-1].file_id
    pid = new_payment(uid, "card", s, "pending", file_id, n)
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Tasdiqlash", callback_data=f"adm:ok:{pid}")
    kb.button(text="❌ Rad etish", callback_data=f"adm:no:{pid}")
    kb.adjust(2)
    await m.bot.send_photo(
        ADMIN_ID, file_id, reply_markup=kb.as_markup(),
        caption=f"💳 To'lov so'rovi #{pid}\n👤 {m.from_user.full_name} (id {uid})\n"
                f"💰 {som(s)} so'm → {n} ta taqdimot")
    await state.clear()
    await m.answer("🙏 Kechirasiz, chekingizni admin tekshirishi kerak.\n"
                   "Iltimos, biroz kuting. Tasdiqlangach balansingiz to'ldiriladi "
                   "va sizga xabar keladi.",
                   reply_markup=main_kb())


@router.message(Form.receipt, ~F.text.startswith("/"))
async def receipt_not_photo(m: Message):
    await m.answer("📸 Iltimos, chek <b>rasmini</b> yuboring yoki /cancel bosing.",
                   parse_mode="HTML")


@router.callback_query(F.data.startswith("adm:"))
async def admin_decide(c: CallbackQuery):
    if c.from_user.id != ADMIN_ID:
        await c.answer("Ruxsat yo'q", show_alert=True)
        return
    _, action, pid = c.data.split(":")
    ok = action == "ok"
    row = finish_payment(int(pid), "approved" if ok else "rejected")
    if not row:
        await c.answer("Bu so'rov allaqachon ko'rib chiqilgan", show_alert=True)
        return
    uid = row["user_id"]
    n = row["credits"] or 1
    try:
        if ok:
            add_credits(uid, n)
            await c.bot.send_message(uid, f"✅ To'lov tasdiqlandi! +{n} ta taqdimot.\n"
                                          f"💎 Jami: {get_credits(uid)} ta\n\n📊 Boshlash: /new")
        else:
            await c.bot.send_message(uid, "❌ To'lovingiz tasdiqlanmadi. "
                                          "Chek noto'g'ri bo'lsa, qayta yuboring: ➕ Balansni to'ldirish")
    except Exception as e:
        log.warning("Foydalanuvchiga xabar ketmadi: %s", e)
    await c.message.edit_caption(
        caption=(c.message.caption or "") + ("\n\n✅ Tasdiqlandi" if ok else "\n\n❌ Rad etildi"),
        reply_markup=None)
    await c.answer("Bajarildi")


@router.message(Command("stats"))
async def cmd_stats(m: Message):
    if m.from_user.id != ADMIN_ID:
        return
    with closing(db()) as c:
        users = c.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        made = c.execute("SELECT COALESCE(SUM(made),0) FROM users").fetchone()[0]
        stars = c.execute("SELECT COUNT(*), COALESCE(SUM(amount),0) FROM payments "
                          "WHERE kind='stars'").fetchone()
        card = c.execute("SELECT COUNT(*), COALESCE(SUM(amount),0) FROM payments "
                         "WHERE kind='card' AND status='approved'").fetchone()
        pend = c.execute("SELECT COUNT(*) FROM payments WHERE status='pending'").fetchone()[0]
        fb = c.execute("SELECT COUNT(score), COALESCE(AVG(score),0) FROM feedback "
                       "WHERE score IS NOT NULL").fetchone()
        by_style = c.execute("SELECT style, COUNT(score) n, AVG(score) a FROM feedback "
                             "WHERE score IS NOT NULL GROUP BY style ORDER BY a DESC").fetchall()
    lines = [f"  {STYLES.get(r['style'], r['style'])}: {r['a']:.1f} ⭐️ ({r['n']} ta)" for r in by_style]
    await m.answer(
        f"📈 <b>Statistika</b>\n👥 Foydalanuvchilar: {users}\n📊 Yaratilgan taqdimotlar: {made}\n"
        f"⭐️ Stars: {stars[0]} ta, jami {stars[1]} ⭐️\n"
        f"💳 Karta: {card[0]} ta, jami {som(card[1])} so'm\n⏳ Kutilayotgan: {pend}\n\n"
        f"💬 Baholar: {fb[0]} ta, o'rtacha {fb[1]:.1f}/5"
        + ("\n🎨 Maketlar bo'yicha:\n" + "\n".join(lines) if lines else ""),
        parse_mode="HTML")


@router.message(Command("give"))
async def cmd_give(m: Message):
    """Admin: /give <user_id> <son>  - qo'lda limit berish"""
    if m.from_user.id != ADMIN_ID:
        return
    try:
        _, uid, n = m.text.split()
        add_credits(int(uid), int(n))
        await m.answer(f"✅ {uid} ga +{n} ta berildi. Jami: {get_credits(int(uid))}")
    except Exception:
        await m.answer("Format: /give 123456789 5")


@router.message(F.text, ~F.text.startswith("/"))
async def fallback(m: Message, state: FSMContext):
    """Holat yo'qolgan bo'lsa (bot qayta ishga tushgan), matnni yangi mavzu deb olamiz."""
    await take_topic(m, state)


async def notify_missed(bot: Bot):
    """Bot o'chiq paytda yozgan foydalanuvchilarni topib, 'tuzatildi' deb xabar beradi.
    Telegram bunday xabarlarni ~24 soat saqlaydi."""
    users = {}
    offset = None
    for _ in range(30):
        try:
            ups = await bot.get_updates(offset=offset, limit=100, timeout=0,
                                        allowed_updates=["message", "callback_query"])
        except Exception as e:
            log.warning("Kutilayotgan xabarlarni o'qib bo'lmadi: %s", e)
            break
        if not ups:
            break
        for u in ups:
            if u.message and u.message.from_user:
                user, chat = u.message.from_user, u.message.chat.id
            elif u.callback_query and u.callback_query.message:
                user, chat = u.callback_query.from_user, u.callback_query.message.chat.id
            else:
                continue
            if not user.is_bot:
                users[chat] = user
        offset = ups[-1].update_id + 1  # keyingi so'rov bularni "o'qilgan" qilib qo'yadi

    sent = 0
    for chat_id in users:
        try:
            await bot.send_message(
                chat_id,
                "✅ Kechirasiz, bot vaqtincha ishlamay turgan edi.\n"
                "Muammo tuzatildi, endi hammasi ishlayapti!\n\n"
                "📊 Taqdimot yaratish uchun /new ni bosing.",
            )
            sent += 1
        except Exception as e:  # foydalanuvchi botni bloklagan bo'lishi mumkin
            log.warning("Xabar yuborilmadi (%s): %s", chat_id, e)
        await asyncio.sleep(0.05)
    log.info("O'tkazib yuborilgan foydalanuvchilarga xabar: %s ta", sent)


async def main():
    if BOT_TOKEN.startswith("BOT_TOKEN_NI") or GEMINI_KEY.startswith("GEMINI_KALIT"):
        raise SystemExit("❗ BOT_TOKEN va GEMINI_API_KEY ni kiriting (faylning yuqori qismida)")
    init_db()
    bot = Bot(BOT_TOKEN)
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    await bot.delete_webhook(drop_pending_updates=False)  # eski xabarlar saqlansin
    await notify_missed(bot)
    log.info("Bot ishga tushdi")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
