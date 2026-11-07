"""
📊 AI Slayd Bot (TEKIN versiya) - KENGAYTIRILGAN

Mavzu (yoki PDF/DOCX/TXT fayl) -> slayd soni -> til -> rang -> maket (namuna rasm bilan)
-> rasm (ha/yo'q) -> REJA tasdiqlash -> tayyor .pptx (+ PDF) -> tahrirlash / dizayn almashtirish.

Imkoniyatlar:
  - Reja tasdiqlash, grafik va jadval slaydlari, har slaydga izoh (Notes)
  - Fayldan taqdimot (PDF, DOCX, TXT), muallif ismi va logo
  - Slaydni tahrirlash (qayta yoz / qisqartir / misol / grafik / jadval / o'z ko'rsatmang)
  - Dizaynni almashtirish (matnni qayta yaratmasdan), PDF variant, /history
  - Referal tizimi, /broadcast, /stats, /user, /give
  - Holat SQLite'da saqlanadi (bot qayta ishga tushsa yo'qolmaydi)
  - Bir nechta Gemini kalit (GEMINI_API_KEYS=kalit1,kalit2), generatsiya navbati
  - Yakuniy slayd matnini AI mavzuga qarab o'zi yozadi (closing_title / closing_sub)
  - YANGI: slayd soni tugmalari + "Boshqa" (o'zi xohlagan sonni yozadi)
  - YANGI: SPAM himoyasi (juda tez bosgan/yozganni vaqtincha to'xtatadi)
  - YANGI: avtomatik ZAXIRA (bazani adminga yuboradi) + /backup
  - YANGI: kunlik BONUS (150 so'm) + "Kunlik bonus" tugmasi va /bonus
  - YANGI: karta to'lovida chek skrinshot YOKI PDF ko'rinishida qabul qilinadi

O'rnatish:  pip install -r requirements.txt
PDF uchun serverda LibreOffice kerak (nixpacks.toml ga qarang); bo'lmasa PDF tugmasi chiqmaydi.
"""
import asyncio
import glob
import gzip
import html
import io
import json
import logging
import os
import pickle
import random
import re
import shutil
import sqlite3
import tempfile
import time
from contextlib import asynccontextmanager, closing

import aiohttp
from aiogram import BaseMiddleware, Bot, Dispatcher, F, Router
from aiogram.exceptions import TelegramRetryAfter
from aiogram.filters import Command, CommandObject, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.base import BaseStorage, StorageKey
from aiogram.types import (BotCommand, BufferedInputFile, CallbackQuery, FSInputFile,
                           KeyboardButton, LabeledPrice, Message, PreCheckoutQuery,
                           ReplyKeyboardMarkup)
from aiogram.utils.keyboard import InlineKeyboardBuilder
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# ======================= SOZLAMALAR (shu yerni to'ldiring) =======================
BOT_TOKEN = os.getenv("BOT_TOKEN") or "BOT_TOKEN_NI_SHU_YERGA"
# Bitta kalit yoki vergul bilan bir nechta: GEMINI_API_KEYS=kalit1,kalit2,kalit3
GEMINI_KEYS = [k.strip() for k in (os.getenv("GEMINI_API_KEYS") or os.getenv("GEMINI_API_KEY") or "").split(",")
               if k.strip()] or ["GEMINI_KALITNI_SHU_YERGA"]
PEXELS_KEY = os.getenv("PEXELS_API_KEY") or ""  # ixtiyoriy; bo'sh bo'lsa Wikimedia ishlatiladi
ADMIN_ID = int(os.getenv("ADMIN_ID") or 0)      # sizning Telegram ID raqamingiz (@userinfobot dan oling)
CARD_NUMBER = os.getenv("CARD_NUMBER") or ""      # karta raqami
CARD_OWNER = os.getenv("CARD_OWNER") or ""        # karta egasi ismi
# Railway'da Volume ulansa, RAILWAY_VOLUME_MOUNT_PATH avtomatik beriladi: baza shu yerda saqlanadi
_VOL = os.getenv("RAILWAY_VOLUME_MOUNT_PATH") or ""
DB_PATH = os.getenv("DB_PATH") or (os.path.join(_VOL, "bot.db") if _VOL else "bot.db")  # qo'lda: /data/bot.db
PERSISTENT_DB = bool(os.getenv("DB_PATH") or _VOL)  # False bo'lsa baza deployda o'chib ketadi!

FREE_LIMIT = 3          # har bir yangi foydalanuvchiga bepul taqdimotlar soni
FEEDBACK_RATE = 0.3     # har 10 ta taqdimotdan ~3 tasidan keyin fikr so'raladi (sinash uchun 1.0 qiling)
REF_BONUS_SOM = 1000    # do'st birinchi taqdimotini yaratganda taklif qilganga balansiga tushadigan summa (so'm)
                        # bu pul faqat bot ichida ishlatiladi, yechib bo'lmaydi
REF_MAX = 20            # bitta foydalanuvchi maksimal nechta do'st uchun bonus oladi
EDIT_LIMIT = 5          # bitta taqdimotni bepul tahrirlash soni
KEEP_DECKS = 10         # /history da saqlanadigan taqdimotlar soni
KEEP_IMG = 3            # shulardan nechtasining rasmlari bazada saqlanadi (qolganlari kerak bo'lsa qayta qidiriladi)
SOURCE_MAX = 20000      # fayldan olinadigan matn uzunligi (belgi)
GEN_CONCURRENCY = int(os.getenv("GEN_CONCURRENCY") or 3)  # bir vaqtda nechta generatsiya

BONUS_SOM = 150         # kunlik bonus (so'm); har kuni bir marta, faqat bot ichida ishlatiladi
TZ_OFFSET = 5 * 3600    # "kun" Toshkent vaqti (UTC+5) bo'yicha yangilanadi

MIN_SLIDES = 3          # "Boshqa" tugmasi orqali yozish mumkin bo'lgan eng kam slayd soni
MAX_SLIDES = 30         # ... va eng ko'p soni

SPAM_WINDOW = 5         # soniya: shu vaqt ichida
SPAM_MAX = 8            # ... shuncha so'rovdan ko'p bo'lsa spam hisoblanadi
SPAM_BLOCK = 30         # soniya: spamchi shuncha vaqt to'xtatiladi

BACKUP_HOURS = float(os.getenv("BACKUP_HOURS") or 24)  # zaxira necha soatda bir marta
BACKUP_KEEP = 5         # serverda oxirgi nechta zaxira fayl saqlansin

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
UA = {"User-Agent": "AISlaydBot/2.0 (telegram bot)"}
PDF_OK = bool(shutil.which("soffice") or shutil.which("libreoffice"))

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("slide-bot")

SLIDE_COUNTS = [5, 8, 10, 12, 15, 18, 20]

LANGS = {
    "uz": ("🇺🇿 O'zbekcha", "Uzbek (Latin script)"),
    "ru": ("🇷🇺 Русский", "Russian"),
    "en": ("🇬🇧 English", "English"),
}
# Zaxira matn: AI yakuniy matn bermasa (yoki eski taqdimotlarda) shu chiqadi
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
STYLE_LABELS = {  # namuna rasmidagi yozuvlar (emoji'siz)
    "classic": "Klassik", "side": "Yon panel", "minimal": "Minimal", "cards": "Kartochkali",
    "split": "Ikki ustunli", "bignum": "Katta raqam", "timeline": "Chiziqli",
    "gradient": "Gradient", "magazine": "Jurnal",
}

# Tahrirlash amallari: kalit -> (tugma nomi, AI uchun ko'rsatma)
EDIT_ACTIONS = {
    "rewrite": ("🔄 Qayta yoz", "Rewrite this slide with fresh wording and different facts and examples. Same topic."),
    "short":   ("✂️ Qisqartir", "Make this slide shorter: exactly 3 concise bullets, each at most 12 words."),
    "more":    ("➕ Ko'proq misol", "Expand this slide with more concrete examples, numbers and facts (5 bullets)."),
    "chart":   ("📊 Grafikka aylantir", "Convert this slide into a chart slide (type chart) with realistic numbers relevant to its topic. Add a caption."),
    "table":   ("🧮 Jadvalga aylantir", "Convert this slide into a table slide (type table) comparing relevant items. Add a caption."),
}

router = Router()
busy_users: set[int] = set()
bg_tasks: set = set()
BOT_USERNAME = ""
gen_sem: asyncio.Semaphore | None = None
pdf_sem: asyncio.Semaphore | None = None
_waiting = 0


class Form(StatesGroup):
    topic = State()
    count_input = State()
    receipt = State()
    comment = State()
    edit_text = State()
    author = State()
    logo = State()
    broadcast = State()


# ================================ BAZA ================================
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


def add_column(c, table: str, coldef: str):
    try:
        c.execute(f"ALTER TABLE {table} ADD COLUMN {coldef}")
        c.commit()
    except sqlite3.OperationalError:
        pass  # ustun allaqachon bor


def init_db():
    prepare_db_path()
    with closing(db()) as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(
            user_id INTEGER PRIMARY KEY, name TEXT, credits INTEGER NOT NULL,
            made INTEGER NOT NULL DEFAULT 0, created INTEGER,
            referrer INTEGER, ref_paid INTEGER NOT NULL DEFAULT 0, author TEXT, logo BLOB,
            balance INTEGER NOT NULL DEFAULT 0, bonus_day TEXT);
        CREATE TABLE IF NOT EXISTS payments(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, kind TEXT,
            amount INTEGER, status TEXT, receipt TEXT, ts INTEGER, credits INTEGER);
        CREATE TABLE IF NOT EXISTS feedback(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, name TEXT, topic TEXT,
            theme TEXT, style TEXT, score INTEGER, comment TEXT, ts INTEGER);
        CREATE TABLE IF NOT EXISTS decks(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, topic TEXT, content TEXT,
            theme TEXT, style TEXT, lang TEXT, want_images INTEGER, images BLOB,
            edits INTEGER NOT NULL DEFAULT 0, ts INTEGER);
        CREATE TABLE IF NOT EXISTS fsm(
            key TEXT PRIMARY KEY, state TEXT, data TEXT NOT NULL DEFAULT '{}');
        CREATE TABLE IF NOT EXISTS meta(
            key TEXT PRIMARY KEY, value TEXT);
        """)
        c.commit()
        # eski bazalar uchun
        add_column(c, "payments", "credits INTEGER")
        add_column(c, "users", "referrer INTEGER")
        add_column(c, "users", "ref_paid INTEGER NOT NULL DEFAULT 0")
        add_column(c, "users", "author TEXT")
        add_column(c, "users", "logo BLOB")
        add_column(c, "users", "balance INTEGER NOT NULL DEFAULT 0")
        add_column(c, "users", "bonus_day TEXT")
        add_column(c, "users", "last_seen INTEGER")


def meta_get(key: str, default: str = ""):
    with closing(db()) as c:
        r = c.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return r["value"] if r else default


def meta_set(key: str, value: str):
    with closing(db()) as c:
        c.execute("INSERT INTO meta(key,value) VALUES(?,?) "
                  "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))
        c.commit()


# ---------------- FSM holati SQLite'da (qayta ishga tushsa yo'qolmaydi) ----------------
class SQLiteStorage(BaseStorage):
    @staticmethod
    def _k(key: StorageKey) -> str:
        return f"{key.bot_id}:{key.chat_id}:{key.user_id}:{key.thread_id}:{key.destiny}"

    async def set_state(self, key: StorageKey, state=None) -> None:
        s = state.state if isinstance(state, State) else state
        with closing(db()) as c:
            c.execute("INSERT INTO fsm(key,state,data) VALUES(?,?,'{}') "
                      "ON CONFLICT(key) DO UPDATE SET state=excluded.state", (self._k(key), s))
            c.commit()

    async def get_state(self, key: StorageKey):
        with closing(db()) as c:
            r = c.execute("SELECT state FROM fsm WHERE key=?", (self._k(key),)).fetchone()
        return r["state"] if r else None

    async def set_data(self, key: StorageKey, data) -> None:
        with closing(db()) as c:
            c.execute("INSERT INTO fsm(key,state,data) VALUES(?,NULL,?) "
                      "ON CONFLICT(key) DO UPDATE SET data=excluded.data",
                      (self._k(key), json.dumps(data, ensure_ascii=False)))
            c.commit()

    async def get_data(self, key: StorageKey) -> dict:
        with closing(db()) as c:
            r = c.execute("SELECT data FROM fsm WHERE key=?", (self._k(key),)).fetchone()
        try:
            return json.loads(r["data"]) if r else {}
        except ValueError:
            return {}

    async def close(self) -> None:
        pass


# ---------------- Foydalanuvchilar / limit ----------------
def ensure_user(uid: int, name: str = ""):
    with closing(db()) as c:
        c.execute("INSERT OR IGNORE INTO users(user_id,name,credits,created) VALUES(?,?,?,?)",
                  (uid, name, FREE_LIMIT, int(time.time())))
        c.commit()


def user_exists(uid: int) -> bool:
    with closing(db()) as c:
        return c.execute("SELECT 1 FROM users WHERE user_id=?", (uid,)).fetchone() is not None


def get_credits(uid: int, name: str = "") -> int:
    ensure_user(uid, name)
    if uid == ADMIN_ID:
        return 10**6  # admin cheksiz
    with closing(db()) as c:
        return c.execute("SELECT credits FROM users WHERE user_id=?", (uid,)).fetchone()[0]


def add_credits(uid: int, n: int):
    ensure_user(uid)
    with closing(db()) as c:
        c.execute("UPDATE users SET credits=credits+? WHERE user_id=?", (n, uid))
        c.commit()


def reserve_credit(uid: int) -> bool:
    """Generatsiya BOSHLANISHIDAN OLDIN 1 ta limitni atomik yechadi. Yetmasa False."""
    if uid == ADMIN_ID:
        return True
    with closing(db()) as c:
        cur = c.execute("UPDATE users SET credits=credits-1 WHERE user_id=? AND credits>0", (uid,))
        c.commit()
        ok = cur.rowcount == 1
    log.info("Limit yechildi: uid=%s ok=%s", uid, ok)
    return ok


def refund_credit(uid: int):
    """Generatsiya muvaffaqiyatsiz bo'lsa, yechilgan limitni qaytaradi."""
    if uid == ADMIN_ID:
        return
    add_credits(uid, 1)
    log.info("Limit qaytarildi: uid=%s", uid)


def get_money(uid: int) -> int:
    """Foydalanuvchining bot ichidagi so'm balansi (faqat bot ichida ishlatiladi)."""
    ensure_user(uid)
    with closing(db()) as c:
        return c.execute("SELECT balance FROM users WHERE user_id=?", (uid,)).fetchone()[0] or 0


def add_money(uid: int, n: int):
    ensure_user(uid)
    with closing(db()) as c:
        c.execute("UPDATE users SET balance=balance+? WHERE user_id=?", (n, uid))
        c.commit()


def spend_money(uid: int, n: int) -> bool:
    """Balansdan n so'm yechadi. Yetmasa hech narsa o'zgarmaydi va False qaytaradi."""
    with closing(db()) as c:
        cur = c.execute("UPDATE users SET balance=balance-? WHERE user_id=? AND balance>=?", (n, uid, n))
        c.commit()
        return cur.rowcount == 1


# ---------------- Kunlik bonus ----------------
def today_key() -> str:
    """Bugungi sana (Toshkent vaqti bilan)."""
    return time.strftime("%Y-%m-%d", time.gmtime(time.time() + TZ_OFFSET))


def claim_daily(uid: int) -> bool:
    """Kunlik bonusni beradi. Bugun olingan bo'lsa False (ikki marta olib bo'lmaydi)."""
    ensure_user(uid)
    day = today_key()
    with closing(db()) as c:
        cur = c.execute("UPDATE users SET balance=balance+?, bonus_day=? "
                        "WHERE user_id=? AND (bonus_day IS NULL OR bonus_day<>?)",
                        (BONUS_SOM, day, uid, day))
        c.commit()
        return cur.rowcount == 1


def touch_user(uid: int):
    """Foydalanuvchining oxirgi faolligini yozib qo'yadi (statistika uchun)."""
    try:
        with closing(db()) as c:
            c.execute("UPDATE users SET last_seen=? WHERE user_id=?", (int(time.time()), uid))
            c.commit()
    except Exception:
        pass


def bonus_wait_text() -> str:
    """Keyingi bonusgacha qancha vaqt qolgani (Toshkent yarim kechasigacha)."""
    left = 86400 - ((time.time() + TZ_OFFSET) % 86400)
    h, mnt = int(left // 3600), int(left % 3600 // 60)
    return f"{h} soat {mnt} daqiqa"


def consume(uid: int):
    """Limitni kamaytiradi. Agar bu foydalanuvchining birinchi taqdimoti bo'lsa va uni
    kimdir taklif qilgan bo'lsa, taklif qilganga bonus beradi va uning ID sini qaytaradi."""
    if uid == ADMIN_ID:
        return None
    with closing(db()) as c:
        c.execute("UPDATE users SET made=made+1 WHERE user_id=?", (uid,))  # limit oldindan yechilgan (reserve_credit)
        row = c.execute("SELECT referrer, ref_paid FROM users WHERE user_id=?", (uid,)).fetchone()
        c.commit()
    if row and row["referrer"] and not row["ref_paid"]:
        ref = row["referrer"]
        with closing(db()) as c:
            paid = c.execute("SELECT COUNT(*) FROM users WHERE referrer=? AND ref_paid=1", (ref,)).fetchone()[0]
            c.execute("UPDATE users SET ref_paid=1 WHERE user_id=?", (uid,))
            c.commit()
        if paid < REF_MAX:
            add_money(ref, REF_BONUS_SOM)
            return ref
    return None


def set_referrer(uid: int, ref: int):
    with closing(db()) as c:
        c.execute("UPDATE users SET referrer=? WHERE user_id=? AND referrer IS NULL", (ref, uid))
        c.commit()


def ref_stats(uid: int):
    with closing(db()) as c:
        a = c.execute("SELECT COUNT(*) FROM users WHERE referrer=?", (uid,)).fetchone()[0]
        b = c.execute("SELECT COUNT(*) FROM users WHERE referrer=? AND ref_paid=1", (uid,)).fetchone()[0]
    return a, b


def get_profile(uid: int):
    with closing(db()) as c:
        r = c.execute("SELECT author, logo FROM users WHERE user_id=?", (uid,)).fetchone()
    return ((r["author"] or ""), r["logo"]) if r else ("", None)


def set_author(uid: int, name: str):
    ensure_user(uid)
    with closing(db()) as c:
        c.execute("UPDATE users SET author=? WHERE user_id=?", (name or None, uid))
        c.commit()


def set_logo(uid: int, data):
    ensure_user(uid)
    with closing(db()) as c:
        c.execute("UPDATE users SET logo=? WHERE user_id=?", (data, uid))
        c.commit()


# ---------------- To'lovlar / fikr ----------------
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


# ---------------- Taqdimotlar tarixi ----------------
def save_deck(uid, topic, content, theme, style, lang, want_images, images) -> int:
    blob = pickle.dumps(images) if (want_images and images) else None
    with closing(db()) as c:
        cur = c.execute("INSERT INTO decks(user_id,topic,content,theme,style,lang,want_images,images,ts) "
                        "VALUES(?,?,?,?,?,?,?,?,?)",
                        (uid, topic, json.dumps(content, ensure_ascii=False), theme, style, lang,
                         1 if want_images else 0, blob, int(time.time())))
        did = cur.lastrowid
        c.execute("DELETE FROM decks WHERE user_id=? AND id NOT IN "
                  "(SELECT id FROM decks WHERE user_id=? ORDER BY id DESC LIMIT ?)", (uid, uid, KEEP_DECKS))
        c.execute("UPDATE decks SET images=NULL WHERE user_id=? AND id NOT IN "
                  "(SELECT id FROM decks WHERE user_id=? ORDER BY id DESC LIMIT ?)", (uid, uid, KEEP_IMG))
        c.commit()
    return did


def get_deck(did: int, uid: int):
    with closing(db()) as c:
        return c.execute("SELECT * FROM decks WHERE id=? AND user_id=?", (did, uid)).fetchone()


def list_decks(uid: int, limit: int = KEEP_DECKS):
    with closing(db()) as c:
        return c.execute("SELECT id, topic, ts FROM decks WHERE user_id=? ORDER BY id DESC LIMIT ?",
                         (uid, limit)).fetchall()


def update_deck(did: int, content, images, edits: int):
    with closing(db()) as c:
        c.execute("UPDATE decks SET content=?, images=?, edits=? WHERE id=?",
                  (json.dumps(content, ensure_ascii=False), pickle.dumps(images) if images else None,
                   edits, did))
        c.commit()


def set_deck_design(did: int, theme: str, style: str):
    with closing(db()) as c:
        c.execute("UPDATE decks SET theme=?, style=? WHERE id=?", (theme, style, did))
        c.commit()


def set_deck_images(did: int, images):
    with closing(db()) as c:
        c.execute("UPDATE decks SET images=? WHERE id=?", (pickle.dumps(images) if images else None, did))
        c.commit()


def som(n: int) -> str:
    return f"{n:,}".replace(",", " ")


# ============================ AVTOMATIK ZAXIRA ============================
def _gzip_file(src: str, dst: str):
    with open(src, "rb") as f_in, gzip.open(dst, "wb", compresslevel=6) as f_out:
        shutil.copyfileobj(f_in, f_out)


def make_backup() -> str:
    """Bazaning to'liq nusxasini (gzip) yaratadi. Fayl yo'lini qaytaradi.
    Telegram 50 MB dan katta faylni qabul qilmaydi: katta bo'lsa rasmlarsiz yengil nusxa olinadi."""
    folder = os.path.join(os.path.dirname(os.path.abspath(DB_PATH)), "backups")
    os.makedirs(folder, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    tmp = os.path.join(folder, f"bot_{stamp}.db")
    gz = tmp + ".gz"
    src = sqlite3.connect(DB_PATH)
    dst = sqlite3.connect(tmp)
    try:
        src.backup(dst)  # ishlab turgan bazadan xavfsiz nusxa
    finally:
        dst.close()
        src.close()
    _gzip_file(tmp, gz)
    if os.path.getsize(gz) > 48 * 1024 * 1024:
        lite = sqlite3.connect(tmp)
        try:
            lite.execute("UPDATE decks SET images=NULL")
            lite.execute("DELETE FROM fsm")
            lite.commit()
            lite.execute("VACUUM")
        finally:
            lite.close()
        _gzip_file(tmp, gz)
    rm(tmp)
    # eski zaxiralarni tozalash
    olds = sorted(glob.glob(os.path.join(folder, "bot_*.db.gz")))
    for p in olds[:-BACKUP_KEEP]:
        rm(p)
    return gz


async def run_backup(bot: Bot) -> bool:
    try:
        path = await asyncio.to_thread(make_backup)
        meta_set("last_backup", str(int(time.time())))
        kb = max(os.path.getsize(path) // 1024, 1)
        if ADMIN_ID:
            await bot.send_document(
                ADMIN_ID, FSInputFile(path, filename=os.path.basename(path)),
                caption=(f"💾 Bazaning zaxira nusxasi ({kb} KB)\n"
                         f"🕐 {time.strftime('%d.%m.%Y %H:%M', time.gmtime(time.time() + TZ_OFFSET))}\n\n"
                         "Tiklash: faylni .gz dan chiqaring (gunzip), bot.db nomi bilan "
                         "DB_PATH joyiga qo'ying va botni qayta ishga tushiring."))
        return True
    except Exception as e:
        log.exception("Zaxira xatosi")
        if ADMIN_ID:
            try:
                await bot.send_message(ADMIN_ID, f"⚠️ Zaxira olishda xato: {type(e).__name__}: {str(e)[:200]}")
            except Exception:
                pass
        return False


async def backup_loop(bot: Bot):
    """Har BACKUP_HOURS soatda bazani adminga yuboradi. Oxirgi zaxira vaqti bazada saqlanadi,
    shuning uchun bot tez-tez qayta ishga tushsa ham xabarlar ko'payib ketmaydi."""
    interval = max(int(BACKUP_HOURS * 3600), 600)
    while True:
        try:
            last = int(meta_get("last_backup", "0") or 0)
        except ValueError:
            last = 0
        wait = max(60, last + interval - int(time.time()))
        await asyncio.sleep(wait)
        ok = await run_backup(bot)
        if not ok:
            await asyncio.sleep(3600)  # xato bo'lsa 1 soatdan keyin qayta uriniladi


# ============================ SPAM HIMOYASI ============================
class AntiSpam(BaseMiddleware):
    """Juda tez yozgan/bosgan foydalanuvchini SPAM_BLOCK soniyaga to'xtatadi. Admin uchun ishlamaydi."""

    def __init__(self):
        self.hits: dict[int, list] = {}
        self.until: dict[int, float] = {}
        self.seen: dict[int, float] = {}

    def _prune(self, now: float):
        for uid in list(self.hits):
            h = self.hits[uid]
            if (not h or now - h[-1] > SPAM_WINDOW) and self.until.get(uid, 0) <= now:
                self.hits.pop(uid, None)
                self.until.pop(uid, None)

    async def __call__(self, handler, event, data):
        user = data.get("event_from_user")
        if user is None or is_admin(user.id):
            return await handler(event, data)
        # To'lov xabarini hech qachon tashlab yubormaymiz
        if isinstance(event, Message) and event.successful_payment:
            return await handler(event, data)
        now = time.monotonic()
        if self.until.get(user.id, 0) > now:
            if isinstance(event, CallbackQuery):
                try:
                    await event.answer("⏳ Juda tez! Biroz kuting.")
                except Exception:
                    pass
            return None
        hits = [t for t in self.hits.get(user.id, []) if now - t < SPAM_WINDOW]
        hits.append(now)
        self.hits[user.id] = hits
        if len(hits) > SPAM_MAX:
            self.until[user.id] = now + SPAM_BLOCK
            self.hits[user.id] = []
            warn = f"⛔ Juda tez-tez so'rov yuboryapsiz. {SPAM_BLOCK} soniya kuting."
            try:
                if isinstance(event, CallbackQuery):
                    await event.answer(warn, show_alert=True)
                elif isinstance(event, Message):
                    await event.answer(warn)
            except Exception:
                pass
            return None
        if len(self.hits) > 5000:
            self._prune(now)
        if now - self.seen.get(user.id, 0) > 60:  # faollikni daqiqada bir marta yozamiz
            self.seen[user.id] = now
            touch_user(user.id)
        return await handler(event, data)


# ================================ KLAVIATURALAR ================================
BTN_NEW = "📊 Yangi taqdimot"
BTN_HIST = "🗂 Tarix"
BTN_BAL = "💎 Balans"
BTN_BUY = "➕ Balansni to'ldirish"
BTN_REF = "🎁 Do'st taklif"
BTN_SET = "⚙️ Sozlamalar"
BTN_HELP = "ℹ️ Yordam"
BTN_BONUS = "🎉 Kunlik bonus"


def main_kb():
    """Doimiy pastki menyu."""
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=BTN_NEW), KeyboardButton(text=BTN_HIST)],
                  [KeyboardButton(text=BTN_BAL), KeyboardButton(text=BTN_BUY)],
                  [KeyboardButton(text=BTN_REF), KeyboardButton(text=BTN_SET)],
                  [KeyboardButton(text=BTN_HELP)],
                  [KeyboardButton(text=BTN_BONUS)]],
        resize_keyboard=True, is_persistent=True,
        input_field_placeholder="Mavzuni yozing, fayl yuboring yoki tugmani bosing")


def count_kb():
    kb = InlineKeyboardBuilder()
    for n in SLIDE_COUNTS:
        kb.button(text=f"📊 {n} slayd", callback_data=f"count:{n}")
    kb.button(text="✏️ Boshqa (o'zim yozaman)", callback_data="count:other")
    kb.adjust(2, 2, 2, 1, 1)
    return kb.as_markup()


def lang_kb():
    kb = InlineKeyboardBuilder()
    for key, (name, _) in LANGS.items():
        kb.button(text=name, callback_data=f"lang:{key}")
    kb.adjust(1)
    return kb.as_markup()


def theme_kb(prefix: str = "theme:"):
    kb = InlineKeyboardBuilder()
    for key, (name, *_r) in THEMES.items():
        kb.button(text=name, callback_data=f"{prefix}{key}")
    kb.adjust(2)
    return kb.as_markup()


def style_kb(prefix: str = "style:"):
    kb = InlineKeyboardBuilder()
    for n, (key, name) in enumerate(STYLES.items(), start=1):
        label = f"{n}. {name}" if key != "random" else name
        kb.button(text=label, callback_data=f"{prefix}{key}")
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


def images_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="🖼 Ha, rasm qo'shilsin", callback_data="img:yes")
    kb.button(text="📄 Yo'q, rasmsiz", callback_data="img:no")
    kb.adjust(1)
    return kb.as_markup()


def plan_kb(regen: int):
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Ma'qul, boshlash", callback_data="plan:ok")
    if regen < 2:
        kb.button(text="🔄 Boshqa reja", callback_data="plan:redo")
    kb.button(text="❌ Bekor qilish", callback_data="plan:no")
    kb.adjust(1)
    return kb.as_markup()


def deck_kb(did: int):
    kb = InlineKeyboardBuilder()
    kb.button(text="✏️ Tahrirlash", callback_data=f"ed:menu:{did}")
    kb.button(text="🎨 Dizayn", callback_data=f"rd:menu:{did}")
    if PDF_OK:
        kb.button(text="📄 PDF", callback_data=f"pdf:{did}")
    kb.adjust(2)
    return kb.as_markup()


def settings_kb(has_author: bool, has_logo: bool):
    kb = InlineKeyboardBuilder()
    kb.button(text="👤 Muallif ismini o'zgartirish", callback_data="set:author")
    kb.button(text="🖼 Logo yuklash", callback_data="set:logo")
    if has_author:
        kb.button(text="🗑 Muallifni o'chirish", callback_data="set:author_off")
    if has_logo:
        kb.button(text="🗑 Logoni o'chirish", callback_data="set:logo_off")
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
            params={"pageSize": 200}, headers={"x-goog-api-key": GEMINI_KEYS[0]},
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


_rr = 0


def rotated_keys() -> list:
    """Kalitlarni navbatma-navbat aylantiradi (yuk teng taqsimlansin)."""
    global _rr
    n = len(GEMINI_KEYS)
    _rr = (_rr + 1) % n
    return GEMINI_KEYS[_rr:] + GEMINI_KEYS[:_rr]


async def call_gemini(session, model: str, key: str, body: dict):
    """Natija: ('ok', matn) | ('key', xato) | ('quota', xato) | ('model', xato) | ('retry', xato)"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    try:
        async with session.post(url, json=body, headers={"x-goog-api-key": key}) as r:
            j = await r.json(content_type=None)
            status = r.status
    except (aiohttp.ClientError, asyncio.TimeoutError, ValueError) as e:
        return "retry", f"{model}: {type(e).__name__}"
    if status == 200:
        try:
            parts = j["candidates"][0]["content"]["parts"]
            out = "".join(p.get("text", "") for p in parts)
        except (KeyError, IndexError, TypeError):
            return "model", f"{model}: javob bo'sh yoki bloklangan"
        return ("ok", out) if out.strip() else ("model", f"{model}: bo'sh javob")
    msg = (j.get("error", {}) or {}).get("message", str(j)) if isinstance(j, dict) else str(j)
    err = f"{model} [{status}]: {msg[:120]}"
    if status in (401, 403) or (status == 400 and "API key" in msg):
        return "key", err
    if status == 429:
        return "quota", err
    if status in (400, 404):
        return "model", err
    return "retry", err


async def gemini(prompt: str) -> str:
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0.7,
                             "maxOutputTokens": 16384,  # katta (20-30 slaydli) taqdimotlar uchun
                             "thinkingConfig": {"thinkingBudget": 0}},  # "o'ylash"ni o'chiramiz: ancha tez
    }
    errors = []
    dead = set()
    keys = rotated_keys()
    timeout = aiohttp.ClientTimeout(total=240)
    async with aiohttp.ClientSession(timeout=timeout) as session:
        order = []
        for m in [os.getenv("AI_MODEL", "")] + MODELS + await discover_models(session):
            if m and m not in order:
                order.append(m)
        for model in order[:8]:
            bad_model = False
            mbody = {**body, "generationConfig": dict(body["generationConfig"])}
            for key in keys:
                if key in dead:
                    continue
                for attempt in range(2):
                    kind, val = await call_gemini(session, model, key, mbody)
                    if kind == "ok":
                        return val
                    errors.append(val)
                    if (kind == "model" and "thinking" in val.lower()
                            and "thinkingConfig" in mbody["generationConfig"]):
                        mbody["generationConfig"].pop("thinkingConfig")  # bu model o'chirishni qabul qilmadi
                        continue
                    if kind == "key":
                        dead.add(key)  # kalit xato: boshqa kalitga o't
                        break
                    if kind == "quota":
                        break  # limit tugadi: boshqa kalitga o't
                    if kind == "model":
                        bad_model = True
                        break
                    await asyncio.sleep(3 * (attempt + 1))
                if bad_model:
                    break
            if len(dead) == len(keys):
                raise RuntimeError(errors[-1] if errors else "API kalit xato")
    raise RuntimeError(" | ".join(errors[-3:]) or "noma'lum xato")


def extract_json(raw: str) -> dict:
    raw = re.sub(r"```(?:json)?", "", raw)
    a, b = raw.find("{"), raw.rfind("}")
    if a == -1 or b == -1:
        raise ValueError("AI JSON qaytarmadi")
    return json.loads(raw[a : b + 1])


# ---------------- Slayd ma'lumotini tekshirish ----------------
CHART_KINDS = ("column", "bar", "line", "pie")

SLIDE_SCHEMA = """Slide object formats (pick "type" for each slide):
1) {"type":"bullets","title":"Slide title","bullets":["point 1","point 2","point 3","point 4"],"notes":"2-3 sentences the presenter can say aloud","image_query":"2-3 ENGLISH words for a stock photo search"}
2) {"type":"chart","title":"Slide title","chart":{"kind":"column|bar|line|pie","series":"What is measured (with unit)","categories":["A","B","C"],"values":[10,20,30]},"caption":"one-sentence takeaway","bullets":["fallback point 1","fallback point 2","fallback point 3"],"notes":"..."}
3) {"type":"table","title":"Slide title","table":{"headers":["Col 1","Col 2","Col 3"],"rows":[["a","b","c"],["d","e","f"]]},"caption":"one-sentence takeaway","bullets":["fallback point 1","fallback point 2","fallback point 3"],"notes":"..."}"""


def _num(x):
    try:
        if isinstance(x, str):
            x = x.replace(" ", "").replace(",", ".").rstrip("%")
        return float(x)
    except (TypeError, ValueError):
        return None


def norm_slide(sl: dict) -> dict:
    """AI qaytargan slaydni tekshiradi; grafik/jadval noto'g'ri bo'lsa oddiy slaydga aylantiradi."""
    out = {
        "type": "bullets",
        "title": str(sl.get("title", "")).strip()[:90],
        "bullets": [str(b).strip() for b in (sl.get("bullets") or []) if str(b).strip()][:6],
        "notes": str(sl.get("notes", "")).strip(),
        "image_query": str(sl.get("image_query", "")).strip(),
        "caption": str(sl.get("caption", "")).strip()[:160],
    }
    kind = str(sl.get("type", "bullets")).lower()
    if kind == "chart" and isinstance(sl.get("chart"), dict):
        ch = sl["chart"]
        cats = [str(x).strip()[:30] for x in (ch.get("categories") or [])]
        vals = [_num(v) for v in (ch.get("values") or [])]
        pairs = [(a, b) for a, b in zip(cats, vals) if a and b is not None][:8]
        ck = str(ch.get("kind", "column")).lower()
        if ck not in CHART_KINDS:
            ck = "column"
        pie_bad = ck == "pie" and (any(v < 0 for _, v in pairs) or sum(v for _, v in pairs) <= 0)
        if len(pairs) >= 2 and not pie_bad:
            out["type"] = "chart"
            out["chart"] = {"kind": ck, "series": str(ch.get("series") or out["title"])[:40],
                            "categories": [a for a, _ in pairs], "values": [b for _, b in pairs]}
    elif kind == "table" and isinstance(sl.get("table"), dict):
        tb = sl["table"]
        headers = [str(h).strip()[:30] for h in (tb.get("headers") or [])][:5]
        if len(headers) >= 2:
            rows = []
            for r in (tb.get("rows") or [])[:7]:
                if isinstance(r, (list, tuple)):
                    cells = [str(x).strip()[:60] for x in r][:len(headers)]
                    cells += [""] * (len(headers) - len(cells))
                    rows.append(cells)
            if len(rows) >= 2:
                out["type"] = "table"
                out["table"] = {"headers": headers, "rows": rows}
    if out["type"] == "bullets" and not out["bullets"]:
        out["bullets"] = [out["caption"] or out["title"] or "-"]
    return out


async def generate_content(topic: str, total: int, lang: str, source: str = "") -> dict:
    n = total - 2  # titul va yakuniy slaydlar alohida
    src = ""
    if source:
        src = ('\n\nSource material. Base the presentation ONLY on this text and do not invent '
               'facts that contradict it:\n"""\n' + source + '\n"""')
    prompt = f"""Topic: "{topic}"{src}

Create content for a professional presentation. Write ALL text in {LANGS[lang][1]}.
Exactly {n} main slides.

Reply with ONLY this JSON, nothing else:
{{
  "title": "Presentation title",
  "subtitle": "Short subtitle",
  "closing_title": "Short memorable closing line for the final slide, max 8 words",
  "closing_sub": "One-sentence key takeaway or call to action for the final slide, max 15 words",
  "slides": [ ...{n} slide objects... ]
}}

{SLIDE_SCHEMA}

Rules:
- Most slides are type "bullets": 3-5 bullets, each 8-18 words, concrete and informative (facts, numbers, examples).
- Use at most 2 slides of type "chart" or "table", and ONLY if the topic has real numeric or comparative data (otherwise use only "bullets"). Chart: 3-6 categories, values are plain numbers (put the unit in "series"). Table: 3-5 columns, 3-6 rows. Use well-known approximate figures and never fabricate precise statistics.
- Logical flow: introduction, main parts, examples/statistics, conclusion.
- Slide titles are short (max 6 words).
- closing_title and closing_sub must be specific to this topic (NOT a generic "Thank you" or "Any questions?").
- image_query must be a concrete visual subject, always in English."""
    last = None
    best = None  # kam slayd qaytsa ham eng yaxshi natijani saqlab qo'yamiz
    for _ in range(2):
        try:
            data = extract_json(await gemini(prompt))
            raw = data.get("slides")
            if isinstance(raw, list):
                slides = [norm_slide(s) for s in raw if isinstance(s, dict)][:n]
                if slides:
                    res = {"title": str(data.get("title", "")).strip(),
                           "subtitle": str(data.get("subtitle", "")).strip(),
                           "closing_title": str(data.get("closing_title", "")).strip()[:70],
                           "closing_sub": str(data.get("closing_sub", "")).strip()[:120],
                           "slides": slides}
                    if len(slides) >= n:
                        return res
                    if best is None or len(slides) > len(best["slides"]):
                        best = res
                    last = ValueError("Slaydlar yetarli emas")
                    continue
            last = ValueError("Slaydlar bo'sh")
        except (ValueError, KeyError, json.JSONDecodeError) as e:
            last = e
    if best:
        return best
    raise last


async def edit_slide(topic: str, lang: str, slide: dict, instruction: str) -> dict:
    cur = {k: v for k, v in slide.items() if k != "image_query" or v}
    prompt = f"""Presentation topic: "{topic}". Write ALL text in {LANGS[lang][1]}.

Current slide (JSON):
{json.dumps(cur, ensure_ascii=False)}

Task: {instruction}

{SLIDE_SCHEMA}

Reply with ONLY the JSON object of the new slide. Keep it concrete and informative."""
    last = None
    for _ in range(2):
        try:
            data = extract_json(await gemini(prompt))
            new = norm_slide(data)
            if not new["title"]:
                new["title"] = slide.get("title", "")
            if not new["notes"]:
                new["notes"] = slide.get("notes", "")
            if not new["image_query"]:
                new["image_query"] = slide.get("image_query", "")
            return new
        except (ValueError, KeyError, json.JSONDecodeError) as e:
            last = e
    raise last


# ---------------- Fayldan matn olish ----------------
def extract_text(path: str, ext: str) -> str:
    if ext == "pdf":
        from pypdf import PdfReader
        reader = PdfReader(path)
        out = "\n".join((p.extract_text() or "") for p in reader.pages[:60])
    elif ext == "docx":
        import docx
        d = docx.Document(path)
        out = "\n".join(p.text for p in d.paragraphs)
        for t in d.tables:
            for row in t.rows:
                out += "\n" + " | ".join(cell.text for cell in row.cells)
    else:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            out = f.read()
    out = re.sub(r"[ \t]+", " ", out)
    out = re.sub(r"\n\s*\n+", "\n", out).strip()
    return out[:SOURCE_MAX]


# ---------------- Navbat ----------------
@asynccontextmanager
async def gen_slot(status: Message | None = None):
    """Bir vaqtda GEN_CONCURRENCY tadan ko'p generatsiya ketmasin; qolganlar navbatda turadi."""
    global _waiting
    queued = gen_sem.locked()
    if queued:
        _waiting += 1
        if status is not None:
            try:
                await status.edit_text(f"🕐 Navbatdasiz: {_waiting}-o'rin. Biroz kuting, tez orada boshlanadi...")
            except Exception:
                pass
    try:
        await gen_sem.acquire()
    finally:
        if queued:
            _waiting -= 1
    try:
        yield
    finally:
        gen_sem.release()


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
    img.thumbnail((900, 1000))
    out = io.BytesIO()
    img.save(out, "JPEG", quality=80)
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
    """Faqat oddiy (bullets) slaydlar uchun rasm qidiradi; grafik/jadval slaydlarga rasm kerak emas."""
    timeout = aiohttp.ClientTimeout(total=12)
    async with aiohttp.ClientSession(timeout=timeout, headers=UA) as session:
        sem = asyncio.Semaphore(8)

        async def one(sl):
            if sl.get("type", "bullets") != "bullets":
                return None
            async with sem:
                return await fetch_image(session, sl.get("image_query", ""))

        return await asyncio.gather(*[one(sl) for sl in slides])


img_prefetch: dict = {}  # uid -> (slaydlar, vazifa): reja ko'rilayotganda rasmlar fonda yuklanadi


def cancel_prefetch(uid: int):
    item = img_prefetch.pop(uid, None)
    if item and not item[1].done():
        item[1].cancel()


async def take_prefetched(uid: int, slides):
    """Fonda yuklangan rasmlarni oladi; reja o'zgargan bo'lsa qaytadan qidiradi."""
    t0 = time.monotonic()
    item = img_prefetch.pop(uid, None)
    res = None
    if item and item[0] == slides:
        try:
            res = await item[1]
        except Exception:
            res = None
    elif item and not item[1].done():
        item[1].cancel()
    if res is None:
        res = await fetch_all_images(slides)
    log.info("Rasmlar tayyor: %.1f soniya", time.monotonic() - t0)
    return res


def prep_logo(raw: bytes) -> bytes:
    """Logoni PNG ga o'tkazadi (shaffoflik saqlanadi) va kichraytiradi."""
    img = Image.open(io.BytesIO(raw)).convert("RGBA")
    img.thumbnail((500, 500))
    out = io.BytesIO()
    img.save(out, "PNG")
    return out.getvalue()


# ============================ MAKET NAMUNALARI (preview) ============================
def draw_tile(d, x0, y0, w, h, style, dark, accent, light, txt):
    """Bitta maketning soddalashtirilgan rasmini chizadi."""
    ln = mix(txt, light, 0.5)
    head = "FFFFFF" if lum(light) < 0.4 else dark
    white = "FFFFFF" if lum(light) >= 0.4 else mix(light, "FFFFFF", 0.1)

    def R(a, b, c, e, col):
        d.rectangle([x0 + a * w, y0 + b * h, x0 + c * w, y0 + e * h], fill="#" + col)

    def O(cx, cy, r, col):
        d.ellipse([x0 + cx * w - r, y0 + cy * h - r, x0 + cx * w + r, y0 + cy * h + r], fill="#" + col)

    def lines(x1, y1, x2, n, gap=0.13, col=None, dots=None):
        for k in range(n):
            yy = y1 + k * gap
            if dots:
                O(x1 - 0.04, yy + 0.017, 5, dots)
            R(x1, yy, x2 - (0.12 if k % 2 else 0), yy + 0.035, col or ln)

    if style == "gradient":
        gb, gt = mix(dark, accent, 0.4), dark
        for yy in range(int(h)):
            d.line([x0, y0 + yy, x0 + w, y0 + yy], fill="#" + mix(gt, gb, yy / h))
        R(0, 0, 1, .04, accent)
        O(.08, .13, 10, accent)
        R(.15, .11, .62, .15, "FFFFFF")
        lines(.12, .34, .85, 4, col="D1D5DB", dots=accent)
        return
    bg = "FFFFFF" if style == "magazine" else light
    R(0, 0, 1, 1, bg)
    if style == "classic":
        R(0, 0, 1, .2, dark)
        R(0, .2, 1, .225, accent)
        O(.07, .1, 10, accent)
        R(.14, .085, .6, .12, "FFFFFF")
        lines(.12, .36, .85, 4, dots=accent)
    elif style == "side":
        R(0, 0, .3, 1, dark)
        R(.3, 0, .312, 1, accent)
        R(.05, .12, .14, .22, accent)
        R(.05, .32, .24, .36, "FFFFFF")
        R(.05, .4, .2, .44, "FFFFFF")
        lines(.4, .15, .93, 5, gap=.14, dots=accent)
    elif style == "minimal":
        R(.08, .08, .16, .13, accent)
        R(.08, .2, .7, .27, head)
        R(.08, .31, .2, .325, accent)
        lines(.1, .45, .9, 4, dots=accent)
    elif style == "cards":
        O(.09, .12, 10, accent)
        R(.16, .1, .6, .14, head)
        for r_ in range(2):
            for c_ in range(2):
                x, y = .08 + c_ * .45, .3 + r_ * .3
                R(x, y, x + .42, y + .26, white)
                R(x, y, x + .02, y + .26, accent)
                R(x + .05, y + .08, x + .35, y + .1, ln)
                R(x + .05, y + .15, x + .28, y + .17, ln)
    elif style == "split":
        R(0, 0, .38, 1, dark)
        R(.38, 0, .392, 1, accent)
        R(.06, .25, .3, .5, accent)
        R(.06, .6, .3, .66, "FFFFFF")
        lines(.47, .15, .94, 5, gap=.14, dots=accent)
    elif style == "bignum":
        R(.06, .06, .32, .3, accent)
        R(.4, .12, .9, .2, head)
        R(.06, .38, .94, .395, accent)
        lines(.1, .5, .9, 4, gap=.12, dots=accent)
    elif style == "timeline":
        R(.08, .2, .9, .23, bg)
        R(.1, .22, .112, .86, accent)
        R(.08, .06, .6, .12, head)
        for k in range(4):
            O(.106, .24 + k * .2, 8, accent)
            R(.17, .22 + k * .2, .82 - .1 * (k % 2), .255 + k * .2, ln)
    else:  # magazine
        R(.06, .08, .8, .18, "111111")
        R(.06, .28, .94, .285, "111111")
        R(.06, .275, .25, .295, accent)
        for k in range(4):
            R(.06, .4 + k * .14, .12, .46 + k * .14, accent)
            R(.16, .41 + k * .14, .82 - .1 * (k % 2), .445 + k * .14, "9CA3AF")


def style_preview(theme_key: str) -> bytes:
    """Barcha 9 maketning tanlangan rangdagi namunasi (3x3 jadval)."""
    _, dark, accent, light, txt = THEMES[theme_key]
    tw, th, g = 400, 225, 14
    cols, rows = 3, 3
    im = Image.new("RGB", (cols * tw + (cols + 1) * g, rows * th + (rows + 1) * g), "#E5E7EB")
    d = ImageDraw.Draw(im)
    try:
        font = ImageFont.load_default(size=20)
    except TypeError:
        font = ImageFont.load_default()
    for n, key in enumerate(STYLE_KEYS):
        r_, c_ = divmod(n, cols)
        x0, y0 = g + c_ * (tw + g), g + r_ * (th + g)
        draw_tile(d, x0, y0, tw, th, key, dark, accent, light, txt)
        d.rectangle([x0, y0, x0 + tw - 1, y0 + th - 1], outline="#9CA3AF")
        label = f"{n + 1}. {STYLE_LABELS[key]}"
        lw = d.textlength(label, font=font)
        d.rectangle([x0 + 8, y0 + th - 36, x0 + 24 + lw, y0 + th - 8], fill="#111827")
        d.text((x0 + 16, y0 + th - 34), label, fill="#FFFFFF", font=font)
    out = io.BytesIO()
    im.save(out, "JPEG", quality=85)
    return out.getvalue()


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


def add_logo(s, logo: bytes):
    """Titul slaydning o'ng yuqori burchagiga oq fonli logo qo'yadi."""
    try:
        iw, ih = Image.open(io.BytesIO(logo)).size
    except Exception:
        return
    bw, bh = 1.9, 1.2
    k = min(bw / iw, bh / ih)
    w, h = iw * k, ih * k
    bx, by = 13.333 - 0.7 - bw, 0.45
    rect(s, Inches(bx - 0.12), Inches(by - 0.12), Inches(bw + 0.24), Inches(bh + 0.24), "FFFFFF")
    s.shapes.add_picture(io.BytesIO(logo), Inches(bx + (bw - w) / 2), Inches(by + (bh - h) / 2),
                         Inches(w), Inches(h))


# Grafik/jadval slaydlari uchun joy: (x, y, eni, bo'yi) dyuymda - maketga qarab
AREA = {
    "classic": (0.8, 1.85, 11.7, 4.85), "side": (3.6, 1.0, 9.1, 5.7),
    "minimal": (0.8, 2.2, 11.7, 4.5), "cards": (0.8, 1.85, 11.7, 4.85),
    "split": (5.9, 1.0, 6.9, 5.7), "bignum": (0.8, 3.1, 11.7, 3.65),
    "timeline": (1.0, 2.0, 11.7, 4.7), "gradient": (0.8, 1.9, 11.7, 4.8),
    "magazine": (0.8, 2.85, 11.7, 3.95),
}
PIE_COLORS = ["2F80ED", "10B981", "F59E0B", "EF4444", "A855F7", "14B8A6", "F97316", "64748B"]
AUTHOR_X = {"classic": 1.2, "side": 5.2, "minimal": 1.2, "cards": 1.7, "split": 0.9,
            "bignum": 1.0, "timeline": 2.4, "gradient": 1.2, "magazine": 0.8}


def draw_chart(s, sl, x, y, w, h, body, accent):
    ch = sl["chart"]
    kind = ch["kind"]
    cd = CategoryChartData()
    cd.categories = ch["categories"]
    cd.add_series(ch["series"], ch["values"])
    ctype = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
             "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE}[kind]
    chart = s.shapes.add_chart(ctype, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    chart.font.size = Pt(14)
    chart.font.name = "Calibri"
    chart.font.color.rgb = rgb(body)
    chart.has_title = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(14)
    dl.font.color.rgb = rgb(body)
    if kind == "pie":
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(15)
        chart.legend.font.color.rgb = rgb(body)
        dl.show_percentage = True
        dl.show_value = False
        dl.number_format = "0%"
        dl.number_format_is_linked = False
        dl.font.bold = True
        dl.font.color.rgb = rgb("FFFFFF")
        palette = [accent] + PIE_COLORS[1:]
        ser = plot.series[0]
        for i in range(len(ch["values"])):
            pt = ser.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(palette[i % len(palette)])
        return
    chart.has_legend = False
    dl.show_value = True
    ser = plot.series[0]
    if kind == "line":
        dl.position = XL_LABEL_POSITION.ABOVE
        ser.smooth = False
        ser.format.line.color.rgb = rgb(accent)
        ser.format.line.width = Pt(3)
    else:
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        plot.gap_width = 60
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(accent)
    for ax in (chart.category_axis, chart.value_axis):
        ax.tick_labels.font.size = Pt(14)
        ax.tick_labels.font.color.rgb = rgb(body)
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = rgb("9CA3AF")


def draw_table(s, sl, x, y, w, h, accent, dark, dark_now):
    tb = sl["table"]
    rows = [tb["headers"]] + tb["rows"]
    nr, nc = len(rows), len(tb["headers"])
    tbl = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h)).table
    for r in range(nr):
        tbl.rows[r].height = Inches(h / nr)
    size = 12 if (nc >= 4 and nr > 5) else 14 if (nc >= 4 or nr > 5) else 16
    hdr_txt = "FFFFFF" if lum(accent) < 0.6 else "0F172A"
    if dark_now:
        fills, body_txt = (mix(dark, "FFFFFF", 0.10), mix(dark, "FFFFFF", 0.18)), "F1F5F9"
    else:
        fills, body_txt = ("FFFFFF", "F3F4F6"), "1F2937"
    for r, row in enumerate(rows):
        for c_, val in enumerate(row):
            cell = tbl.cell(r, c_)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(accent if r == 0 else fills[r % 2])
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.12)
            cell.margin_right = Inches(0.12)
            tf = cell.text_frame
            tf.word_wrap = True
            run = tf.paragraphs[0].add_run()
            run.text = val
            run.font.size = Pt(size)
            run.font.name = "Calibri"
            run.font.bold = (r == 0 or c_ == 0)
            run.font.color.rgb = rgb(hdr_txt if r == 0 else body_txt)


def draw_rich(s, sl, area, body, accent, dark, dark_now):
    """Grafik yoki jadval slaydining asosiy qismi (+ pastda bir qatorli xulosa)."""
    x, y, w, h = area
    cap = sl.get("caption", "")
    ch_h = h - (0.8 if cap else 0)
    if sl["type"] == "chart":
        draw_chart(s, sl, x, y, w, ch_h, body, accent)
    else:
        draw_table(s, sl, x, y, w, ch_h, accent, dark, dark_now)
    if cap:
        text(s, Inches(x), Inches(y + ch_h + 0.05), Inches(w), Inches(0.7), cap, 16, body,
             anchor=MSO_ANCHOR.MIDDLE)


def build_pptx(data, theme_key, style, lang, path, images=None, author="", logo=None):
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
    dark_now = style == "gradient" or (bg_dark and style != "magazine")  # slayd foni to'qmi

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

    if author:  # muallif ismi
        acol = soft if style in ("classic", "side", "split", "gradient") else (
            "4B5563" if style == "magazine" else txt)
        text(s, Inches(AUTHOR_X[style]), Inches(6.4), Inches(7.5), Inches(0.5), author, 18, acol, True)
    if logo:
        add_logo(s, logo)

    # --------------------------- ASOSIY SLAYDLAR ---------------------------
    for i, sl in enumerate(slides, start=2):
        s = prs.slides.add_slide(blank)
        title = sl.get("title", "")
        num = i - 1
        rich = sl.get("type") in ("chart", "table")
        img = None if rich else (images[i - 2] if images and images[i - 2] else None)
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

        if rich:
            draw_rich(s, sl, AREA[style], body, accent, dark, dark_now)
        else:
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
    # AI mavzuga qarab yozgan matn; bo'lmasa (eski taqdimot) standart "Rahmat" matni
    t1 = data.get("closing_title") or THANKS[lang][0]
    t2 = data.get("closing_sub") or THANKS[lang][1]
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
    # Sarlavha uzun bo'lsa shrift kichraytiriladi (sig'ishi uchun)
    t1_size = 44 if len(t1) <= 40 else 36 if len(t1) <= 60 else 30
    text(s, Inches(0.8), Inches(2.3), W - Inches(1.6), Inches(1.7), t1, t1_size, c1, True,
         PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, Inches(1.2), Inches(4.3), W - Inches(2.4), Inches(1.2), t2, 22, c2,
         align=PP_ALIGN.CENTER)

    prs.save(path)


# =================================== YORDAMCHILAR ===================================
def is_admin(uid: int) -> bool:
    return bool(ADMIN_ID) and uid == ADMIN_ID


def safe_name(topic: str, ext: str = "pptx") -> str:
    return (re.sub(r"[^\w\-]+", "_", topic)[:40] or "taqdimot") + "." + ext


def rm(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


async def safe_edit(msg: Message, body: str, **kw):
    try:
        return await msg.edit_text(body, **kw)
    except Exception:
        return None


def err_text(e: Exception) -> str:
    msg = "😔 Xatolik bo'ldi. Qayta urinib ko'ring: /new"
    if DEBUG:
        msg += f"\n\n🐞 {type(e).__name__}: {str(e)[:300]}"
    return msg


async def render_file(content, theme, style, lang, uid, images) -> str:
    """Taqdimotni .pptx fayl qilib yig'adi (muallif va logo bilan). Fayl yo'lini qaytaradi."""
    author, logo = get_profile(uid)
    fd, path = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    try:
        await asyncio.to_thread(build_pptx, content, theme, style, lang, path, images, author, logo)
    except Exception:
        rm(path)
        raise
    return path


async def deck_images(deck, content):
    """Saqlangan rasmlarni qaytaradi; yo'q bo'lsa (eski taqdimot) qaytadan qidiradi."""
    if not deck["want_images"]:
        return None
    if deck["images"]:
        try:
            imgs = pickle.loads(deck["images"])
            if imgs and len(imgs) == len(content["slides"]):
                return imgs
        except Exception:
            pass
    imgs = await fetch_all_images(content["slides"])
    set_deck_images(deck["id"], imgs)
    return imgs


async def pptx_to_pdf(src: str):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None
    outdir = tempfile.mkdtemp()
    try:
        proc = await asyncio.create_subprocess_exec(
            exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, src,
            env=dict(os.environ, HOME=outdir),
            stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
        try:
            await asyncio.wait_for(proc.communicate(), timeout=150)
        except asyncio.TimeoutError:
            proc.kill()
            return None
        pdf = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
        if not os.path.exists(pdf):
            return None
        fd, dst = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        shutil.move(pdf, dst)
        return dst
    finally:
        shutil.rmtree(outdir, ignore_errors=True)


async def send_style_picker(msg: Message, theme: str, prefix: str, caption: str):
    """Barcha maketlarning namunasi (rasm) bilan tanlash tugmalarini yuboradi."""
    kb = style_kb(prefix)
    try:
        png = await asyncio.to_thread(style_preview, theme)
        await msg.answer_photo(BufferedInputFile(png, filename="maketlar.jpg"),
                               caption=caption, parse_mode="HTML", reply_markup=kb)
    except Exception as e:
        log.warning("Maket namunasi chizilmadi: %s", e)
        await msg.answer(caption, parse_mode="HTML", reply_markup=kb)
    try:
        await msg.delete()
    except Exception:
        pass


async def send_deck(msg: Message, uid: int, did: int, note: str = "", theme=None, style=None):
    """Saqlangan taqdimotni (kerak bo'lsa yangi rang/maket bilan) qayta yig'ib yuboradi."""
    deck = get_deck(did, uid)
    if not deck:
        await msg.answer("😔 Taqdimot topilmadi.")
        return
    if uid in busy_users:
        await msg.answer("⏳ Oldingi so'rov bajarilmoqda, kuting...")
        return
    busy_users.add(uid)
    status = await msg.answer("🛠 Fayl tayyorlanmoqda...")
    path = None
    try:
        content = json.loads(deck["content"])
        theme, style = theme or deck["theme"], style or deck["style"]
        async with gen_slot(status):
            await safe_edit(status, "🛠 Fayl tayyorlanmoqda...")
            images = await deck_images(deck, content)
            path = await render_file(content, theme, style, deck["lang"], uid, images)
        if (theme, style) != (deck["theme"], deck["style"]):
            set_deck_design(did, theme, style)
        await msg.answer_document(
            FSInputFile(path, filename=safe_name(deck["topic"])),
            caption=f"{note}\n🎯 Mavzu: {deck['topic']}\n🎨 {THEMES[theme][0]} · {STYLES[style]}".strip(),
            reply_markup=deck_kb(did))
        await status.delete()
    except Exception as e:
        log.exception("Taqdimotni yuborishda xato")
        await safe_edit(status, err_text(e))
    finally:
        busy_users.discard(uid)
        rm(path)


# =================================== HANDLERLAR ===================================
async def ask_topic(m: Message, state: FSMContext):
    cancel_prefetch(m.from_user.id)
    await state.clear()
    await state.set_state(Form.topic)
    await m.answer(
        "👋 Salom! Men sizga professional taqdimot tayyorlab beraman.\n\n"
        "✍️ Taqdimot mavzusini yozing, masalan:\n"
        "• <i>Global isish</i>\n• <i>Sun'iy intellekt kelajagi</i>\n• <i>Amir Temur</i>\n\n"
        "📄 Yoki <b>PDF / DOCX / TXT</b> fayl yuboring, men shu asosida taqdimot qilaman.",
        parse_mode="HTML", reply_markup=main_kb(),
    )


# ---------- Buyruqlar (holat filtrisiz; ular state handlerlaridan oldin turishi shart) ----------
@router.message(CommandStart())
async def cmd_start(m: Message, state: FSMContext, command: CommandObject):
    uid = m.from_user.id
    arg = (command.args or "").strip()
    if arg.startswith("ref_") and not user_exists(uid):
        try:
            ref = int(arg[4:])
        except ValueError:
            ref = 0
        if ref and ref != uid and user_exists(ref):
            ensure_user(uid, m.from_user.full_name)
            set_referrer(uid, ref)
            await m.answer("🎁 Do'stingiz taklifi bilan keldingiz! Birinchi taqdimotingiz tayyor bo'lgach, "
                           "taklif qilgan do'stingiz bonus oladi.")
    ensure_user(uid, m.from_user.full_name)
    await ask_topic(m, state)


@router.message(Command("new"))
async def cmd_new(m: Message, state: FSMContext):
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
        "1️⃣ Mavzuni yozasiz (yoki PDF/DOCX/TXT fayl yuborasiz)\n"
        f"2️⃣ Slayd soni ({MIN_SLIDES}-{MAX_SLIDES}), til, rang va maketni tanlaysiz (maket namunalari ko'rsatiladi)\n"
        "3️⃣ Rasm kerakmi — tanlaysiz\n"
        "4️⃣ Tayyor <b>rejani</b> ko'rib tasdiqlaysiz\n"
        "5️⃣ .pptx faylni olasiz, so'ng uni tahrirlash, dizaynini almashtirish yoki PDF qilish mumkin\n\n"
        "/new — yangi taqdimot\n/history — oldingi taqdimotlar\n/balance — balans\n"
        f"/bonus — kunlik bonus (+{som(BONUS_SOM)} so'm)\n"
        "/invite — do'st taklif qilib bonus olish\n/settings — muallif ismi va logo\n/cancel — bekor qilish",
        parse_mode="HTML",
    )


@router.message(Command("balance"))
async def cmd_balance(m: Message):
    n = get_credits(m.from_user.id, m.from_user.full_name)
    if is_admin(m.from_user.id):
        await m.answer("👑 Siz adminsiz: cheksiz.")
    else:
        money = get_money(m.from_user.id)
        await m.answer(f"💎 Sizda <b>{n}</b> ta taqdimot qoldi.\n"
                       f"💰 Bonus balans: <b>{som(money)} so'm</b> (faqat botda ishlatiladi)\n"
                       f"🎉 Har kuni /bonus orqali +{som(BONUS_SOM)} so'm oling!\n\n"
                       "➕ Paketni tanlang (balansdan ham to'lash mumkin):",
                       parse_mode="HTML", reply_markup=packs_kb())


@router.message(Command("bonus"))
@router.message(F.text == BTN_BONUS)
async def cmd_bonus(m: Message):
    """Kunlik bonus: har kuni bir marta."""
    uid = m.from_user.id
    ensure_user(uid, m.from_user.full_name)
    if claim_daily(uid):
        await m.answer(
            f"🎉 <b>Kunlik bonus: +{som(BONUS_SOM)} so'm!</b>\n"
            f"💰 Balansingiz: <b>{som(get_money(uid))} so'm</b>\n\n"
            "Ertaga yana kelib, bonusni oling. Bu pul faqat bot ichida taqdimot uchun to'lashda ishlatiladi.",
            parse_mode="HTML", reply_markup=main_kb())
    else:
        await m.answer(
            "⏳ Bugungi bonusni allaqachon olgansiz.\n"
            f"🕐 Keyingi bonus: <b>{bonus_wait_text()}</b> dan keyin.\n"
            f"💰 Balansingiz: <b>{som(get_money(uid))} so'm</b>",
            parse_mode="HTML", reply_markup=main_kb())


@router.message(Command("history"))
async def cmd_history(m: Message):
    rows = list_decks(m.from_user.id)
    if not rows:
        await m.answer("🗂 Hozircha taqdimotlaringiz yo'q. Boshlash: /new")
        return
    kb = InlineKeyboardBuilder()
    for r in rows:
        when = time.strftime("%d.%m", time.localtime(r["ts"]))
        kb.button(text=f"📊 {(r['topic'] or '')[:28]} · {when}", callback_data=f"hist:{r['id']}")
    kb.adjust(1)
    await m.answer(f"🗂 <b>Oxirgi {len(rows)} ta taqdimotingiz</b>\nTanlang, men faylni qayta yuboraman "
                   "(tahrirlash ham mumkin):", parse_mode="HTML", reply_markup=kb.as_markup())


@router.message(Command("invite"))
async def cmd_invite(m: Message):
    uid = m.from_user.id
    a, b = ref_stats(uid)
    link = f"https://t.me/{BOT_USERNAME}?start=ref_{uid}"
    await m.answer(
        "🎁 <b>Do'st taklif qiling — bepul taqdimot oling!</b>\n\n"
        f"Sizning havolangiz:\n{link}\n\n"
        f"Do'stingiz birinchi taqdimotini yaratganda sizning balansingizga <b>{som(REF_BONUS_SOM)} so'm</b> "
        f"tushadi (maksimum {REF_MAX} ta do'st uchun).\n"
        "💰 Bu pulni yechib bo'lmaydi, faqat bot ichida taqdimot uchun to'lashda ishlatiladi.\n\n"
        f"💰 Hozirgi balansingiz: {som(get_money(uid))} so'm\n"
        f"👥 Taklif qilinganlar: {a}\n🎉 Bonus olingan: {b}",
        parse_mode="HTML", disable_web_page_preview=True)


async def show_settings(m: Message, uid: int):
    author, logo = get_profile(uid)
    await m.answer(
        "⚙️ <b>Sozlamalar</b>\n\n"
        f"👤 Muallif: <b>{html.escape(author) if author else 'kiritilmagan'}</b>\n"
        f"🖼 Logo: <b>{'yuklangan' if logo else 'yo`q'}</b>\n\n"
        "Muallif ismi va logo taqdimotning birinchi (titul) slaydida chiqadi.",
        parse_mode="HTML", reply_markup=settings_kb(bool(author), bool(logo)))


@router.message(Command("settings"))
async def cmd_settings(m: Message):
    await show_settings(m, m.from_user.id)


@router.message(Command("author"))
async def cmd_author(m: Message, state: FSMContext, command: CommandObject):
    name = (command.args or "").strip()
    if name:
        set_author(m.from_user.id, name[:60])
        await m.answer(f"✅ Muallif: {name[:60]}")
    else:
        await state.set_state(Form.author)
        await m.answer("👤 Muallif ismini yozing (masalan: Aziz Karimov).\nBekor qilish: /cancel")


@router.message(Command("logo"))
async def cmd_logo(m: Message, state: FSMContext):
    await state.set_state(Form.logo)
    await m.answer("🖼 Logo rasmini yuboring (PNG tavsiya etiladi, fon shaffof bo'lsa yaxshi).\n"
                   "Bekor qilish: /cancel")


@router.message(Command("broadcast"))
async def cmd_broadcast(m: Message, state: FSMContext):
    if not is_admin(m.from_user.id):
        return
    await state.set_state(Form.broadcast)
    await m.answer("📣 Hammaga yuboriladigan xabarni yuboring (matn, rasm, video — farqi yo'q).\n"
                   "Bekor qilish: /cancel")


@router.message(Command("backup"))
async def cmd_backup(m: Message):
    """Admin: bazaning zaxira nusxasini hozir yuborish."""
    if not is_admin(m.from_user.id):
        return
    st = await m.answer("💾 Zaxira tayyorlanmoqda...")
    ok = await run_backup(m.bot)
    await safe_edit(st, "✅ Zaxira nusxa yuborildi." if ok else "⚠️ Zaxira olib bo'lmadi (loglarni ko'ring).")


@router.message(Command("stats"))
async def cmd_stats(m: Message):
    if not is_admin(m.from_user.id):
        return
    now = int(time.time())
    with closing(db()) as c:
        users = c.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        new24 = c.execute("SELECT COUNT(*) FROM users WHERE created>=?", (now - 86400,)).fetchone()[0]
        made = c.execute("SELECT COALESCE(SUM(made),0) FROM users").fetchone()[0]
        edits = c.execute("SELECT COALESCE(SUM(edits),0) FROM decks").fetchone()[0]
        refs = c.execute("SELECT COUNT(*), COALESCE(SUM(ref_paid),0) FROM users "
                         "WHERE referrer IS NOT NULL").fetchone()
        stars = c.execute("SELECT COUNT(*), COALESCE(SUM(amount),0) FROM payments "
                          "WHERE kind='stars'").fetchone()
        card = c.execute("SELECT COUNT(*), COALESCE(SUM(amount),0) FROM payments "
                         "WHERE kind='card' AND status='approved'").fetchone()
        pend = c.execute("SELECT COUNT(*) FROM payments WHERE status='pending'").fetchone()[0]
        fb = c.execute("SELECT COUNT(score), COALESCE(AVG(score),0) FROM feedback "
                       "WHERE score IS NOT NULL").fetchone()
        by_style = c.execute("SELECT style, COUNT(score) n, AVG(score) a FROM feedback "
                             "WHERE score IS NOT NULL GROUP BY style ORDER BY a DESC").fetchall()
        bonus_today = c.execute("SELECT COUNT(*) FROM users WHERE bonus_day=?", (today_key(),)).fetchone()[0]
        never = c.execute("SELECT COUNT(*) FROM users WHERE made=0").fetchone()[0]
        act1 = c.execute("SELECT COUNT(*) FROM users WHERE last_seen>=?", (now - 86400,)).fetchone()[0]
        act7 = c.execute("SELECT COUNT(*) FROM users WHERE last_seen>=?", (now - 7 * 86400,)).fetchone()[0]
        act30 = c.execute("SELECT COUNT(*) FROM users WHERE last_seen>=?", (now - 30 * 86400,)).fetchone()[0]
    lines = [f"  {STYLES.get(r['style'], r['style'])}: {r['a']:.1f} ⭐️ ({r['n']} ta)" for r in by_style]
    try:
        last_bk = int(meta_get("last_backup", "0") or 0)
    except ValueError:
        last_bk = 0
    bk_txt = (time.strftime("%d.%m %H:%M", time.gmtime(last_bk + TZ_OFFSET)) if last_bk else "hali yo'q")
    await m.answer(
        f"📈 <b>Statistika</b>\n👥 Foydalanuvchilar: {users} (24 soatda +{new24})\n"
        f"📊 Yaratilgan taqdimotlar: {made}\n✏️ Tahrirlar: {edits}\n"
        f"🎁 Referallar: {refs[0]} ta (bonus berilgan: {refs[1]})\n"
        f"🎉 Bugun kunlik bonus olganlar: {bonus_today}\n"
        f"😴 Hali taqdimot yaratmaganlar: {never} ta\n"
        f"🟢 Faol: 24 soatda {act1}, 7 kunda {act7}, 30 kunda {act30}\n"
        f"💤 7 kundan beri kirmaganlar: {max(users - act7, 0)} ta\n"
        f"⭐️ Stars: {stars[0]} ta, jami {stars[1]} ⭐️\n"
        f"💳 Karta: {card[0]} ta, jami {som(card[1])} so'm\n⏳ Kutilayotgan: {pend}\n"
        f"💾 Oxirgi zaxira: {bk_txt}\n\n"
        f"💬 Baholar: {fb[0]} ta, o'rtacha {fb[1]:.1f}/5"
        + ("\n🎨 Maketlar bo'yicha:\n" + "\n".join(lines) if lines else ""),
        parse_mode="HTML")


@router.message(Command("give"))
async def cmd_give(m: Message):
    """Admin: /give <user_id> <son>  - qo'lda limit berish"""
    if not is_admin(m.from_user.id):
        return
    try:
        _, uid, n = m.text.split()
        add_credits(int(uid), int(n))
        await m.answer(f"✅ {uid} ga +{n} ta berildi. Jami: {get_credits(int(uid))}")
    except Exception:
        await m.answer("Format: /give 123456789 5")


@router.message(Command("user"))
async def cmd_user(m: Message):
    """Admin: /user <user_id> - foydalanuvchi haqida ma'lumot"""
    if not is_admin(m.from_user.id):
        return
    try:
        uid = int(m.text.split()[1])
    except (IndexError, ValueError):
        await m.answer("Format: /user 123456789")
        return
    with closing(db()) as c:
        r = c.execute("SELECT * FROM users WHERE user_id=?", (uid,)).fetchone()
        nd = c.execute("SELECT COUNT(*) FROM decks WHERE user_id=?", (uid,)).fetchone()[0]
    if not r:
        await m.answer("Topilmadi.")
        return
    a, b = ref_stats(uid)
    await m.answer(
        f"👤 {html.escape(r['name'] or '')} (id {uid})\n💎 Limit: {r['credits']}\n"
        f"💰 Bonus balans: {som(r['balance'] or 0)} so'm\n📊 Yaratgan: {r['made']}\n🗂 Tarixda: {nd}\n🎁 Taklif qilgan: {a} (bonus: {b})\n"
        f"👥 Taklif qilgan odam: {r['referrer'] or '-'}\n✍️ Muallif: {html.escape(r['author'] or '-')}\n"
        f"🎉 Oxirgi kunlik bonus: {r['bonus_day'] or '-'}",
        parse_mode="HTML")


# ---------- Pastki menyu tugmalari ----------
@router.message(F.text == BTN_NEW)
async def btn_new(m: Message, state: FSMContext):
    await ask_topic(m, state)


@router.message(F.text == BTN_HIST)
async def btn_hist(m: Message):
    await cmd_history(m)


@router.message(F.text == BTN_BAL)
async def btn_balance(m: Message):
    await cmd_balance(m)


@router.message(F.text == BTN_BUY)
async def btn_buy(m: Message):
    await m.answer("➕ <b>Balansni to'ldirish</b>\n\nPaketni tanlang:",
                   parse_mode="HTML", reply_markup=packs_kb())


@router.message(F.text == BTN_REF)
async def btn_ref(m: Message):
    await cmd_invite(m)


@router.message(F.text == BTN_SET)
async def btn_set(m: Message):
    await show_settings(m, m.from_user.id)


@router.message(F.text == BTN_HELP)
async def btn_help(m: Message):
    await cmd_help(m)


async def show_paywall(m: Message):
    await m.answer(
        "😔 <b>Bepul limitingiz tugadi.</b>\n\n"
        "➕ Davom etish uchun balansni to'ldiring, do'stlaringizni taklif qiling (/invite) "
        f"yoki har kuni /bonus orqali +{som(BONUS_SOM)} so'm yig'ing.\n"
        "💳 Karta orqali to'lasangiz, chekni <b>skrinshot</b> yoki <b>PDF</b> ko'rinishida yuborishingiz kifoya.\n"
        "Paketni tanlang:",
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
    await state.set_data({"topic": topic})  # eski fayl matni va tanlovlar tozalanadi
    await m.answer(f"📝 Mavzu: <b>{html.escape(topic)}</b>\n\n📊 Nechta slayd kerak?",
                   parse_mode="HTML", reply_markup=count_kb())


# ---------- Tanlov bosqichlari ----------
@router.callback_query(F.data == "count:other")
async def count_other(c: CallbackQuery, state: FSMContext):
    """'Boshqa' tugmasi: foydalanuvchi slayd sonini o'zi yozadi."""
    if "topic" not in await state.get_data():
        await c.answer("Avval mavzuni yozing: /new", show_alert=True)
        return
    await state.set_state(Form.count_input)
    await c.message.edit_text(
        f"✏️ Nechta slayd kerak? <b>{MIN_SLIDES}</b> dan <b>{MAX_SLIDES}</b> gacha son yozing "
        "(masalan: <i>14</i>).\nBekor qilish: /cancel",
        parse_mode="HTML")
    await c.answer()


@router.message(Form.count_input, F.text)
async def got_count_text(m: Message, state: FSMContext):
    if "topic" not in await state.get_data():
        await state.clear()
        await m.answer("Avval mavzuni yozing: /new")
        return
    t = (m.text or "").strip()
    if not t.isdigit() or not MIN_SLIDES <= int(t) <= MAX_SLIDES:
        await m.answer(f"⚠️ {MIN_SLIDES} dan {MAX_SLIDES} gacha bo'lgan son yozing (masalan: 14).")
        return
    n = int(t)
    await state.set_state(None)
    await state.update_data(count=n)
    await m.answer(f"📊 Slaydlar soni: <b>{n}</b>\n\n🌐 Taqdimot tilini tanlang:",
                   parse_mode="HTML", reply_markup=lang_kb())


@router.callback_query(F.data.startswith("count:"))
async def got_count(c: CallbackQuery, state: FSMContext):
    if "topic" not in await state.get_data():
        await c.answer("Avval mavzuni yozing: /new", show_alert=True)
        return
    try:
        n = int(c.data.split(":")[1])
    except ValueError:
        await c.answer("Noma'lum tanlov", show_alert=True)
        return
    if not MIN_SLIDES <= n <= MAX_SLIDES:
        await c.answer("Noto'g'ri son", show_alert=True)
        return
    await state.set_state(None)  # "Boshqa" holatida qolgan bo'lsa tozalaymiz
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
    if lang not in LANGS:
        await c.answer("Noma'lum til", show_alert=True)
        return
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
    if theme not in THEMES:
        await c.answer("Noma'lum rang", show_alert=True)
        return
    await state.update_data(theme=theme)
    await c.answer()
    await send_style_picker(
        c.message, theme, "style:",
        f"🎨 Rang: <b>{THEMES[theme][0]}</b>\n\n📐 Slayd maketini tanlang (namunalar rasmda):")


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
    await c.message.answer(
        f"📐 Maket: <b>{STYLES[style]}</b>\n\n🖼 Slaydlarga rasm qo'shilsinmi?",
        parse_mode="HTML", reply_markup=images_kb())
    try:
        await c.message.delete()
    except Exception:
        pass
    await c.answer()


# ---------- Reja ----------
ICONS = {"chart": "📊 ", "table": "🧮 "}


async def show_plan(msg: Message, content: dict, regen: int):
    lines = [f"{i}. {ICONS.get(sl['type'], '')}{html.escape(sl['title'][:60])}"
             for i, sl in enumerate(content["slides"], start=1)]
    body = (f"📋 <b>{html.escape(content.get('title', ''))}</b>\n"
            f"<i>{html.escape(content.get('subtitle', ''))}</i>\n\n" + "\n".join(lines) +
            "\n\nReja ma'qulmi? Ma'qul bo'lsa, faylni yig'amiz.")
    await msg.edit_text(body, parse_mode="HTML", reply_markup=plan_kb(regen))


async def make_plan(c: CallbackQuery, state: FSMContext, d: dict, regen: int):
    """Reja yaratadi (yoki qayta yaratadi) va tasdiqlash uchun ko'rsatadi."""
    uid = c.from_user.id
    busy_users.add(uid)
    cancel_prefetch(uid)
    status = await c.message.edit_text(
        f"⏳ <b>{html.escape(d['topic'])}</b> bo'yicha reja tayyorlanmoqda...", parse_mode="HTML")
    try:
        async with gen_slot(status):
            await safe_edit(status, f"⏳ <b>{html.escape(d['topic'])}</b> bo'yicha reja tayyorlanmoqda...",
                            parse_mode="HTML")
            t0 = time.monotonic()
            content = await generate_content(d["topic"], d["count"], d["lang"], d.get("source", ""))
            log.info("Reja AI vaqti: %.1f soniya (%s slayd)", time.monotonic() - t0, d["count"])
        await state.update_data(content=content, regen=regen)
        await show_plan(status, content, regen)
        if (await state.get_data()).get("want_images"):  # foydalanuvchi rejani o'qiguncha rasmlar yuklanadi
            img_prefetch[uid] = (content["slides"], asyncio.create_task(fetch_all_images(content["slides"])))
    except Exception as e:
        log.exception("Reja xatosi")
        await safe_edit(status, err_text(e))
        await state.clear()
    finally:
        busy_users.discard(uid)


@router.callback_query(F.data.startswith("img:"))
async def got_images(c: CallbackQuery, state: FSMContext):
    uid = c.from_user.id
    d = await state.get_data()
    if not all(k in d for k in ("topic", "count", "lang", "theme", "style")):
        await c.answer("Avval /new ni bosing", show_alert=True)
        return
    if uid in busy_users:
        await c.answer("⏳ Oldingi so'rov bajarilmoqda, kuting...", show_alert=True)
        return
    if get_credits(uid, c.from_user.full_name) <= 0:
        await c.answer()
        await state.clear()
        await show_paywall(c.message)
        return
    await state.update_data(want_images=c.data.endswith("yes"))
    await c.answer()
    await make_plan(c, state, d, 0)


@router.callback_query(F.data == "plan:redo")
async def plan_redo(c: CallbackQuery, state: FSMContext):
    d = await state.get_data()
    if not all(k in d for k in ("topic", "count", "lang", "content")):
        await c.answer("Reja topilmadi. /new ni bosing", show_alert=True)
        return
    regen = d.get("regen", 0) + 1
    if regen > 2:
        await c.answer("Boshqa reja limiti tugadi", show_alert=True)
        return
    if c.from_user.id in busy_users:
        await c.answer("⏳ Kuting...", show_alert=True)
        return
    await c.answer()
    await make_plan(c, state, d, regen)


@router.callback_query(F.data == "plan:no")
async def plan_no(c: CallbackQuery, state: FSMContext):
    cancel_prefetch(c.from_user.id)
    await state.clear()
    await c.message.edit_text("❌ Bekor qilindi. Yangi taqdimot uchun /new ni bosing.")
    await c.answer()


@router.callback_query(F.data == "plan:ok")
async def plan_ok(c: CallbackQuery, state: FSMContext):
    uid = c.from_user.id
    d = await state.get_data()
    if not all(k in d for k in ("topic", "count", "lang", "theme", "style", "content")):
        await c.answer("Reja topilmadi. /new ni bosing", show_alert=True)
        return
    if uid in busy_users:
        await c.answer("⏳ Oldingi so'rov bajarilmoqda, kuting...", show_alert=True)
        return
    if get_credits(uid, c.from_user.full_name) <= 0:
        await c.answer()
        await state.clear()
        await show_paywall(c.message)
        return

    topic, count, lang, theme, style = d["topic"], d["count"], d["lang"], d["theme"], d["style"]
    content, want_images = d["content"], bool(d.get("want_images"))
    if style == "random":
        style = random.choice(STYLE_KEYS)
    if not reserve_credit(uid):  # limit shu yerda yechiladi (poyga holatlarini oldini oladi)
        await c.answer()
        await state.clear()
        await show_paywall(c.message)
        return
    busy_users.add(uid)
    sent = False
    await c.answer()
    status = await c.message.edit_text(
        f"⏳ <b>{html.escape(topic)}</b> taqdimoti yig'ilmoqda...", parse_mode="HTML")

    path = None
    try:
        async with gen_slot(status):
            images = None
            if want_images:
                await safe_edit(status, "🖼 Rasmlar qidirilmoqda...")
                images = await take_prefetched(uid, content["slides"])
            found = sum(1 for x in (images or []) if x)
            await safe_edit(status, "🛠 Fayl yig'ilmoqda...")
            path = await render_file(content, theme, style, lang, uid, images)

        did = save_deck(uid, topic, content, theme, style, lang, want_images, images)
        extra = f"\n🎨 {THEMES[theme][0]} · {STYLES[style]}"
        if want_images:
            extra += f"\n🖼 Rasmlar: {found} ta"
        if not is_admin(uid):
            extra += f"\n💎 Qolgan limit: {max(get_credits(uid), 0)} ta"
        real_count = len(content["slides"]) + 2  # AI kam slayd qaytargan bo'lishi mumkin
        await c.message.answer_document(
            FSInputFile(path, filename=safe_name(topic)),
            caption=(f"✅ Tayyor! 📊 {real_count} slayd\n🎯 Mavzu: {topic}{extra}\n"
                     "🗣 Har slaydda ma'ruzachi uchun izohlar (Notes) bor.\n\n"
                     "Pastdagi tugmalar bilan tahrirlash, dizaynni almashtirish yoki PDF olish mumkin."),
            reply_markup=deck_kb(did),
        )
        sent = True
        ref = consume(uid)  # statistika va referal bonus
        if ref:
            try:
                await c.bot.send_message(
                    ref, f"🎁 Do'stingiz taklifingiz bilan birinchi taqdimotini yaratdi! "
                         f"Balansingizga +{som(REF_BONUS_SOM)} so'm tushdi.\n"
                         f"💰 Balans: {som(get_money(ref))} so'm (bot ichida to'lashda ishlatiladi)")
            except Exception as e:
                log.warning("Referalga xabar ketmadi: %s", e)

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
        await safe_edit(status, err_text(e))
    finally:
        if not sent:
            refund_credit(uid)  # fayl yuborilmadi: limit qaytariladi
        busy_users.discard(uid)
        await state.clear()
        rm(path)


# ---------- Fayldan taqdimot ----------
@router.message(F.document)
async def got_document(m: Message, state: FSMContext):
    # Boshqa holatda (chek yoki logo) yuborilgan fayl bo'lsa, tegishli handlerga uzatamiz
    cur = await state.get_state()
    if cur == Form.receipt.state:      # chek PDF/hujjat ko'rinishida yuborilgan
        return await got_receipt(m, state)
    if cur == Form.logo.state:         # logo fayl sifatida yuborilgan
        return await got_logo(m, state)

    doc = m.document
    name = doc.file_name or "fayl"
    ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""
    if ext not in ("pdf", "docx", "txt", "md"):
        await m.answer("📄 Faqat PDF, DOCX yoki TXT fayl yuboring. Yoki mavzuni yozing.")
        return
    if (doc.file_size or 0) > 10 * 1024 * 1024:
        await m.answer("⚠️ Fayl katta (10 MB gacha bo'lsin).")
        return
    if get_credits(m.from_user.id, m.from_user.full_name) <= 0:
        await state.clear()
        await show_paywall(m)
        return
    wait = await m.answer("📥 Fayl o'qilmoqda...")
    fd, path = tempfile.mkstemp(suffix="." + ext)
    os.close(fd)
    try:
        await m.bot.download(doc, destination=path)
        src = await asyncio.to_thread(extract_text, path, ext)
    except Exception as e:
        log.warning("Faylni o'qib bo'lmadi: %s", e)
        await safe_edit(wait, "😔 Faylni o'qib bo'lmadi. Boshqa fayl yuboring yoki mavzuni yozing.")
        return
    finally:
        rm(path)
    if len(src) < 200:
        await safe_edit(wait, "⚠️ Fayldan yetarli matn topilmadi (skaner qilingan PDF bo'lishi mumkin). "
                              "Boshqa fayl yuboring yoki mavzuni yozing.")
        return
    topic = ((m.caption or "").strip() or name.rsplit(".", 1)[0].replace("_", " "))[:200]
    if len(topic) < 3:
        topic = "Hujjat asosida taqdimot"
    await state.set_state(None)
    await state.set_data({"topic": topic, "source": src})
    await safe_edit(wait, f"📄 Fayl o'qildi: <b>{html.escape(name)}</b> ({len(src)} belgi)\n"
                          f"📝 Mavzu: <b>{html.escape(topic)}</b>\n\n📊 Nechta slayd kerak?",
                    parse_mode="HTML", reply_markup=count_kb())


# ---------- Tarix ----------
@router.callback_query(F.data.startswith("hist:"))
async def hist_open(c: CallbackQuery):
    await c.answer()
    await send_deck(c.message, c.from_user.id, int(c.data.split(":")[1]))


# ---------- Tahrirlash ----------
def edits_left(deck, uid: int) -> str:
    return "∞" if is_admin(uid) else str(max(EDIT_LIMIT - deck["edits"], 0))


@router.callback_query(F.data.startswith("ed:menu:"))
async def ed_menu(c: CallbackQuery):
    did = int(c.data.split(":")[2])
    deck = get_deck(did, c.from_user.id)
    if not deck:
        await c.answer("Taqdimot topilmadi", show_alert=True)
        return
    content = json.loads(deck["content"])
    kb = InlineKeyboardBuilder()
    for i, sl in enumerate(content["slides"]):
        kb.button(text=f"{i + 1}. {sl['title'][:20]}", callback_data=f"ed:s:{did}:{i}")
    kb.adjust(2)
    body = f"✏️ Qaysi slaydni tahrirlaymiz?\nQolgan tahrirlar: {edits_left(deck, c.from_user.id)}"
    if c.message.document:  # fayl xabaridan chaqirilsa - yangi xabar
        await c.message.answer(body, reply_markup=kb.as_markup())
    else:
        await c.message.edit_text(body, reply_markup=kb.as_markup())
    await c.answer()


@router.callback_query(F.data.startswith("ed:s:"))
async def ed_slide_menu(c: CallbackQuery):
    _, _, did, idx = c.data.split(":")
    deck = get_deck(int(did), c.from_user.id)
    if not deck:
        await c.answer("Taqdimot topilmadi", show_alert=True)
        return
    content = json.loads(deck["content"])
    idx = int(idx)
    if not 0 <= idx < len(content["slides"]):
        await c.answer("Slayd topilmadi", show_alert=True)
        return
    kb = InlineKeyboardBuilder()
    for key, (label, _) in EDIT_ACTIONS.items():
        kb.button(text=label, callback_data=f"ed:a:{did}:{idx}:{key}")
    kb.button(text="🖊 O'z ko'rsatmam bor", callback_data=f"ed:c:{did}:{idx}")
    kb.button(text="⬅️ Orqaga", callback_data=f"ed:menu:{did}")
    kb.adjust(1)
    await c.message.edit_text(
        f"✏️ <b>{idx + 1}-slayd:</b> {html.escape(content['slides'][idx]['title'])}\n"
        "Nima qilamiz?", parse_mode="HTML", reply_markup=kb.as_markup())
    await c.answer()


async def run_edit(msg: Message, uid: int, did: int, idx: int, instruction: str):
    deck = get_deck(did, uid)
    if not deck:
        await msg.answer("😔 Taqdimot topilmadi.")
        return
    content = json.loads(deck["content"])
    if not 0 <= idx < len(content["slides"]):
        await msg.answer("😔 Slayd topilmadi.")
        return
    if not is_admin(uid) and deck["edits"] >= EDIT_LIMIT:
        await msg.answer(f"⚠️ Bu taqdimot uchun tahrir limiti ({EDIT_LIMIT} ta) tugadi. "
                         "Yangi taqdimot yaratishingiz mumkin: /new")
        return
    if uid in busy_users:
        await msg.answer("⏳ Oldingi so'rov bajarilmoqda, kuting...")
        return
    busy_users.add(uid)
    status = await msg.answer("✏️ Slayd yangilanmoqda...")
    path = None
    try:
        async with gen_slot(status):
            await safe_edit(status, "✏️ Slayd yangilanmoqda...")
            new = await edit_slide(deck["topic"], deck["lang"], content["slides"][idx], instruction)
            content["slides"][idx] = new
            images = await deck_images(deck, content)
            if (deck["want_images"] and images and new["type"] == "bullets"
                    and images[idx] is None and new.get("image_query")):
                images[idx] = (await fetch_all_images([new]))[0]
            path = await render_file(content, deck["theme"], deck["style"], deck["lang"], uid, images)
        update_deck(did, content, images, deck["edits"] + 1)
        left = "∞" if is_admin(uid) else str(max(EDIT_LIMIT - deck["edits"] - 1, 0))
        await msg.answer_document(
            FSInputFile(path, filename=safe_name(deck["topic"])),
            caption=f"✅ {idx + 1}-slayd yangilandi.\n✏️ Qolgan tahrirlar: {left}",
            reply_markup=deck_kb(did))
        await status.delete()
    except Exception as e:
        log.exception("Tahrirlash xatosi")
        await safe_edit(status, err_text(e))
    finally:
        busy_users.discard(uid)
        rm(path)


@router.callback_query(F.data.startswith("ed:a:"))
async def ed_action(c: CallbackQuery):
    _, _, did, idx, key = c.data.split(":")
    if key not in EDIT_ACTIONS:
        await c.answer("Noma'lum amal", show_alert=True)
        return
    await c.answer()
    await run_edit(c.message, c.from_user.id, int(did), int(idx), EDIT_ACTIONS[key][1])


@router.callback_query(F.data.startswith("ed:c:"))
async def ed_custom(c: CallbackQuery, state: FSMContext):
    _, _, did, idx = c.data.split(":")
    if not get_deck(int(did), c.from_user.id):
        await c.answer("Taqdimot topilmadi", show_alert=True)
        return
    await state.set_state(Form.edit_text)
    await state.set_data({"ed_did": int(did), "ed_idx": int(idx)})
    await c.message.answer(f"🖊 {int(idx) + 1}-slayd uchun ko'rsatmangizni yozing "
                           "(masalan: <i>Iqtisodiy ta'sirga e'tibor ber</i>).\nBekor qilish: /cancel",
                           parse_mode="HTML")
    await c.answer()


@router.message(Form.edit_text, F.text)
async def got_edit_text(m: Message, state: FSMContext):
    d = await state.get_data()
    await state.clear()
    if "ed_did" not in d:
        return
    instr = (m.text or "").strip()[:300]
    await run_edit(m, m.from_user.id, d["ed_did"], d["ed_idx"],
                   f"Follow this user instruction for the slide: {instr}")


# ---------- Dizaynni almashtirish ----------
@router.callback_query(F.data.startswith("rd:menu:"))
async def rd_menu(c: CallbackQuery):
    did = int(c.data.split(":")[2])
    if not get_deck(did, c.from_user.id):
        await c.answer("Taqdimot topilmadi", show_alert=True)
        return
    await c.message.answer("🎨 Yangi rangni tanlang (matn o'zgarmaydi, bepul):",
                           reply_markup=theme_kb(f"rd:t:{did}:"))
    await c.answer()


@router.callback_query(F.data.startswith("rd:t:"))
async def rd_theme(c: CallbackQuery):
    _, _, did, theme = c.data.split(":")
    if theme not in THEMES:
        await c.answer("Noma'lum rang", show_alert=True)
        return
    await c.answer()
    await send_style_picker(c.message, theme, f"rd:s:{did}:{theme}:",
                            f"🎨 Rang: <b>{THEMES[theme][0]}</b>\n\n📐 Yangi maketni tanlang:")


@router.callback_query(F.data.startswith("rd:s:"))
async def rd_style(c: CallbackQuery):
    _, _, did, theme, style = c.data.split(":")
    if theme not in THEMES or style not in STYLES:
        await c.answer("Noma'lum tanlov", show_alert=True)
        return
    if style == "random":
        style = random.choice(STYLE_KEYS)
    await c.answer()
    try:
        await c.message.delete()
    except Exception:
        pass
    await send_deck(c.message, c.from_user.id, int(did), note="🎨 Dizayn almashtirildi.",
                    theme=theme, style=style)


# ---------- PDF ----------
@router.callback_query(F.data.startswith("pdf:"))
async def make_pdf(c: CallbackQuery):
    uid = c.from_user.id
    if not PDF_OK:
        await c.answer("PDF hozircha mavjud emas", show_alert=True)
        return
    deck = get_deck(int(c.data.split(":")[1]), uid)
    if not deck:
        await c.answer("Taqdimot topilmadi", show_alert=True)
        return
    if uid in busy_users:
        await c.answer("⏳ Oldingi so'rov bajarilmoqda, kuting...", show_alert=True)
        return
    busy_users.add(uid)
    await c.answer()
    status = await c.message.answer("📄 PDF tayyorlanmoqda (30-60 soniya)...")
    pptx_path = pdf_path = None
    try:
        content = json.loads(deck["content"])
        async with gen_slot(status):
            await safe_edit(status, "📄 PDF tayyorlanmoqda (30-60 soniya)...")
            images = await deck_images(deck, content)
            pptx_path = await render_file(content, deck["theme"], deck["style"], deck["lang"], uid, images)
            async with pdf_sem:
                pdf_path = await pptx_to_pdf(pptx_path)
        if not pdf_path:
            raise RuntimeError("PDF yaratib bo'lmadi")
        await c.message.answer_document(
            FSInputFile(pdf_path, filename=safe_name(deck["topic"], "pdf")), caption="📄 PDF tayyor!")
        await status.delete()
    except Exception as e:
        log.exception("PDF xatosi")
        await safe_edit(status, err_text(e))
    finally:
        busy_users.discard(uid)
        rm(pptx_path)
        rm(pdf_path)


# ---------- Sozlamalar: muallif va logo ----------
@router.callback_query(F.data == "set:author")
async def set_author_cb(c: CallbackQuery, state: FSMContext):
    await state.set_state(Form.author)
    await c.message.answer("👤 Muallif ismini yozing (masalan: Aziz Karimov).\nBekor qilish: /cancel")
    await c.answer()


@router.callback_query(F.data == "set:logo")
async def set_logo_cb(c: CallbackQuery, state: FSMContext):
    await state.set_state(Form.logo)
    await c.message.answer("🖼 Logo rasmini yuboring (PNG tavsiya etiladi).\nBekor qilish: /cancel")
    await c.answer()


@router.callback_query(F.data == "set:author_off")
async def set_author_off(c: CallbackQuery):
    set_author(c.from_user.id, "")
    await c.answer("Muallif o'chirildi")
    await c.message.delete()


@router.callback_query(F.data == "set:logo_off")
async def set_logo_off(c: CallbackQuery):
    set_logo(c.from_user.id, None)
    await c.answer("Logo o'chirildi")
    await c.message.delete()


@router.message(Form.author, F.text)
async def got_author(m: Message, state: FSMContext):
    name = (m.text or "").strip()[:60]
    if len(name) < 2:
        await m.answer("⚠️ Ism juda qisqa. Qayta yozing yoki /cancel.")
        return
    set_author(m.from_user.id, name)
    await state.clear()
    await m.answer(f"✅ Muallif saqlandi: <b>{html.escape(name)}</b>\nEndi titul slaydda chiqadi.",
                   parse_mode="HTML", reply_markup=main_kb())


@router.message(Form.logo, F.photo | F.document)
async def got_logo(m: Message, state: FSMContext):
    obj = m.photo[-1] if m.photo else m.document
    if not m.photo and not (m.document.mime_type or "").startswith("image/"):
        await m.answer("📸 Iltimos, rasm (PNG/JPG) yuboring yoki /cancel bosing.")
        return
    if (getattr(obj, "file_size", 0) or 0) > 8 * 1024 * 1024:
        await m.answer("⚠️ Rasm katta (8 MB gacha bo'lsin).")
        return
    try:
        buf = await m.bot.download(obj)
        data = await asyncio.to_thread(prep_logo, buf.read())
    except Exception as e:
        log.warning("Logo o'qilmadi: %s", e)
        await m.answer("😔 Rasmni o'qib bo'lmadi. Boshqa rasm yuboring.")
        return
    set_logo(m.from_user.id, data)
    await state.clear()
    await m.answer("✅ Logo saqlandi. Endi titul slaydda chiqadi.", reply_markup=main_kb())


@router.message(Form.logo)
async def logo_not_image(m: Message):
    await m.answer("📸 Iltimos, logo <b>rasmini</b> yuboring yoki /cancel bosing.", parse_mode="HTML")


# ---------- Broadcast (admin) ----------
@router.message(Form.broadcast)
async def bc_preview(m: Message, state: FSMContext):
    if not is_admin(m.from_user.id):
        await state.clear()
        return
    await state.set_data({"bc_chat": m.chat.id, "bc_msg": m.message_id})
    with closing(db()) as c:
        n = c.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    await m.bot.copy_message(m.chat.id, m.chat.id, m.message_id)  # ko'rinishi
    kb = InlineKeyboardBuilder()
    kb.button(text=f"✅ Yuborish ({n} ta)", callback_data="bc:yes")
    kb.button(text="❌ Bekor", callback_data="bc:no")
    kb.adjust(2)
    await m.answer("👆 Shu xabar barcha foydalanuvchilarga yuboriladi. Tasdiqlaysizmi?",
                   reply_markup=kb.as_markup())


async def do_broadcast(bot: Bot, src_chat: int, src_msg: int, admin_chat: int):
    with closing(db()) as c:
        ids = [r[0] for r in c.execute("SELECT user_id FROM users")]
    ok = fail = 0
    for uid in ids:
        try:
            await bot.copy_message(uid, src_chat, src_msg)
            ok += 1
        except TelegramRetryAfter as e:
            await asyncio.sleep(e.retry_after + 1)
            try:
                await bot.copy_message(uid, src_chat, src_msg)
                ok += 1
            except Exception:
                fail += 1
        except Exception:
            fail += 1  # botni bloklagan bo'lishi mumkin
        await asyncio.sleep(0.05)
    await bot.send_message(admin_chat, f"📣 Yuborildi: {ok} ta, xato: {fail} ta")


@router.callback_query(F.data.startswith("bc:"))
async def bc_confirm(c: CallbackQuery, state: FSMContext):
    if not is_admin(c.from_user.id):
        await c.answer("Ruxsat yo'q", show_alert=True)
        return
    d = await state.get_data()
    await state.clear()
    if c.data == "bc:no" or "bc_msg" not in d:
        await c.message.edit_text("❌ Bekor qilindi.")
        await c.answer()
        return
    await c.message.edit_text("📣 Yuborish boshlandi, tugagach xabar beraman...")
    await c.answer()
    task = asyncio.create_task(do_broadcast(c.bot, d["bc_chat"], d["bc_msg"], c.message.chat.id))
    bg_tasks.add(task)
    task.add_done_callback(bg_tasks.discard)


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
    await state.set_data({"fb": int(fid)})
    await c.message.edit_text(
        f"🙏 Rahmat! Bahoyingiz: {'⭐️' * int(score)}\n\n"
        "✍️ Xohlasangiz, fikringizni yozing (nima yoqdi, nima yoqmadi).\n"
        "Yozmasangiz — /skip ni bosing.")
    await c.answer()
    await tell_admin(c.bot, feedback_text(row))


@router.message(Form.comment, F.text)
async def got_comment(m: Message, state: FSMContext):
    fid = (await state.get_data()).get("fb")
    await state.clear()
    if fid:
        row = set_comment(fid, m.from_user.id, m.text[:1000])
        if row:
            await tell_admin(m.bot, feedback_text(row, comment=m.text[:1000]))
    await m.answer("✅ Fikringiz uchun rahmat! Yangi taqdimot: /new", reply_markup=main_kb())


# ================================= TO'LOVLAR =================================
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
    money = get_money(c.from_user.id)
    if money > 0:
        kb.button(text=f"💰 Balansdan to'lash ({som(money)} / {som(s)} so'm)", callback_data=f"pay:bal:{i}")
    if card_on():
        kb.button(text=f"💳 {som(s)} so'm (karta)", callback_data=f"pay:card:{i}")
    kb.button(text="⬅️ Orqaga", callback_data="buy:menu")
    kb.adjust(1)
    await c.message.edit_text(f"📦 Paket: <b>{n} ta taqdimot</b>\n\nTo'lov usulini tanlang:",
                              parse_mode="HTML", reply_markup=kb.as_markup())
    await c.answer()


@router.callback_query(F.data.startswith("pay:bal:"))
async def pay_balance(c: CallbackQuery):
    uid = c.from_user.id
    i = int(c.data.split(":")[2])
    n, _, s = PACKS[i]
    if not spend_money(uid, s):
        have = get_money(uid)
        await c.answer(f"Balans yetarli emas: {som(have)} so'm bor, kerak {som(s)} so'm.", show_alert=True)
        return
    new_payment(uid, "balance", s, "paid", "", n)
    add_credits(uid, n)
    await c.answer()
    await c.message.answer(
        f"✅ Balansdan {som(s)} so'm yechildi. +{n} ta taqdimot.\n"
        f"💎 Jami: {get_credits(uid)} ta\n💰 Qolgan balans: {som(get_money(uid))} so'm\n\n📊 Boshlash: /new",
        reply_markup=main_kb())


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
    await state.set_state(Form.receipt)
    await state.set_data({"pack": i})
    await c.message.answer(
        f"💳 <b>{som(s)} so'm</b> ({n} ta taqdimot) ni shu kartaga o'tkazing:\n\n"
        f"<code>{CARD_NUMBER}</code>\n👤 {html.escape(CARD_OWNER)}\n\n"
        "O'tkazgach <b>chek skrinshoti yoki PDF faylini</b> shu yerga yuboring.\n"
        "🙏 Kechirasiz, karta to'lovini admin qo'lda tekshiradi, shuning uchun "
        "tasdiqlanishi biroz vaqt olishi mumkin.\n\nBekor qilish: /cancel",
        parse_mode="HTML")


@router.message(Form.receipt, F.photo | F.document)
async def got_receipt(m: Message, state: FSMContext):
    uid = m.from_user.id
    # Spam himoyasi: bitta foydalanuvchining kutilayotgan cheklari 3 tadan oshmasin
    with closing(db()) as c:
        pend = c.execute("SELECT COUNT(*) FROM payments WHERE user_id=? AND kind='card' "
                         "AND status='pending'", (uid,)).fetchone()[0]
    if pend >= 3:
        await state.clear()
        await m.answer("⏳ Sizning 3 ta chekingiz hali tekshirilmoqda. Iltimos, admin javobini kuting.",
                       reply_markup=main_kb())
        return

    # Chek: skrinshot (photo) yoki PDF / rasm fayli (document)
    if m.photo:
        file_id, is_photo = m.photo[-1].file_id, True
    else:
        doc = m.document
        mime = doc.mime_type or ""
        if mime != "application/pdf" and not mime.startswith("image/"):
            await m.answer("📸 Chekni <b>skrinshot</b> yoki <b>PDF</b> ko'rinishida yuboring "
                           "yoki /cancel bosing.", parse_mode="HTML")
            return
        if (doc.file_size or 0) > 10 * 1024 * 1024:
            await m.answer("⚠️ Fayl katta (10 MB gacha bo'lsin).")
            return
        file_id, is_photo = doc.file_id, False

    i = (await state.get_data()).get("pack", 0)
    n, _, s = PACKS[i]
    pid = new_payment(uid, "card", s, "pending", file_id, n)
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Tasdiqlash", callback_data=f"adm:ok:{pid}")
    kb.button(text="❌ Rad etish", callback_data=f"adm:no:{pid}")
    kb.adjust(2)
    cap = (f"💳 To'lov so'rovi #{pid}\n👤 {m.from_user.full_name} (id {uid})\n"
           f"💰 {som(s)} so'm → {n} ta taqdimot" + ("" if is_photo else "\n📄 Chek PDF/fayl ko'rinishida"))
    if is_photo:
        await m.bot.send_photo(ADMIN_ID, file_id, reply_markup=kb.as_markup(), caption=cap)
    else:
        await m.bot.send_document(ADMIN_ID, file_id, reply_markup=kb.as_markup(), caption=cap)
    await state.clear()
    await m.answer("🙏 Kechirasiz, chekingizni admin tekshirishi kerak.\n"
                   "Iltimos, biroz kuting. Tasdiqlangach balansingiz to'ldiriladi "
                   "va sizga xabar keladi.",
                   reply_markup=main_kb())


@router.message(Form.receipt)
async def receipt_not_photo(m: Message):
    await m.answer("📸 Iltimos, chekni <b>skrinshot</b> yoki <b>PDF</b> ko'rinishida yuboring "
                   "yoki /cancel bosing.", parse_mode="HTML")


@router.callback_query(F.data.startswith("adm:"))
async def admin_decide(c: CallbackQuery):
    if not is_admin(c.from_user.id):
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
    log.info("Chek #%s %s: uid=%s +%s", pid, "TASDIQLANDI" if ok else "RAD", uid, n)
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


# ================================== ASOSIY ==================================
@router.message(Form.topic, F.text)
async def got_topic(m: Message, state: FSMContext):
    await take_topic(m, state)


@router.message(F.text, ~F.text.startswith("/"))
async def fallback(m: Message, state: FSMContext):
    """Holat bo'lmasa, matnni yangi mavzu deb olamiz."""
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
    global BOT_USERNAME, gen_sem, pdf_sem
    if BOT_TOKEN.startswith("BOT_TOKEN_NI") or GEMINI_KEYS[0].startswith("GEMINI_KALIT"):
        raise SystemExit("❗ BOT_TOKEN va GEMINI_API_KEY ni kiriting (Railway Variables yoki fayl yuqorisida)")
    init_db()
    gen_sem = asyncio.Semaphore(GEN_CONCURRENCY)
    pdf_sem = asyncio.Semaphore(1)
    bot = Bot(BOT_TOKEN)
    BOT_USERNAME = (await bot.get_me()).username
    dp = Dispatcher(storage=SQLiteStorage())
    spam = AntiSpam()                       # spam himoyasi (xabar va tugma bosishlar uchun)
    dp.message.outer_middleware(spam)
    dp.callback_query.outer_middleware(spam)
    dp.include_router(router)
    await bot.set_my_commands([
        BotCommand(command="new", description="📊 Yangi taqdimot"),
        BotCommand(command="bonus", description="🎉 Kunlik bonus"),
        BotCommand(command="history", description="🗂 Oldingi taqdimotlar"),
        BotCommand(command="balance", description="💎 Balans"),
        BotCommand(command="invite", description="🎁 Do'st taklif qilish"),
        BotCommand(command="settings", description="⚙️ Muallif va logo"),
        BotCommand(command="help", description="ℹ️ Yordam"),
        BotCommand(command="cancel", description="❌ Bekor qilish"),
    ])
    if ADMIN_ID and not PERSISTENT_DB:
        try:
            await bot.send_message(
                ADMIN_ID,
                "⚠️ <b>Diqqat:</b> baza doimiy joyda emas (Volume ulanmagan).\n"
                "Har deploydan keyin foydalanuvchilar, limitlar va to'lovlar o'chib ketadi!\n\n"
                "Tuzatish: Railway'da Volume qo'shing (mount path: <code>/data</code>) va "
                "Variables ga <code>DB_PATH=/data/bot.db</code> yozing.",
                parse_mode="HTML")
        except Exception as e:
            log.warning("Volume ogohlantirishi ketmadi: %s", e)
    await bot.delete_webhook(drop_pending_updates=False)  # eski xabarlar saqlansin
    await notify_missed(bot)
    task = asyncio.create_task(backup_loop(bot))          # avtomatik zaxira
    bg_tasks.add(task)
    task.add_done_callback(bg_tasks.discard)
    log.info("Bot ishga tushdi (PDF: %s, Gemini kalitlari: %d, zaxira: har %s soat, baza: %s, doimiy: %s)",
             "bor" if PDF_OK else "yo'q", len(GEMINI_KEYS), BACKUP_HOURS, DB_PATH, PERSISTENT_DB)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
