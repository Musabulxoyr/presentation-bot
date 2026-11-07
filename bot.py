"""
📊 AI Slayd Bot (TEKIN versiya)
Mavzu -> slayd soni -> til -> dizayn -> rasm (ha/yo'q) -> tayyor .pptx

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
import re
import tempfile

import aiohttp
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import CallbackQuery, FSInputFile, Message
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

router = Router()
busy_users: set[int] = set()


class Form(StatesGroup):
    topic = State()


# ================================ KLAVIATURALAR ================================
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


def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


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


def build_pptx(data, theme_key, lang, path, images=None):
    _, dark, accent, light, txt = THEMES[theme_key]
    on_dark = "FFFFFF"
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    W, H = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]
    slides = data["slides"]
    total = len(slides) + 2

    # --- Titul ---
    s = prs.slides.add_slide(blank)
    set_bg(s, dark)
    rect(s, 0, 0, Inches(0.35), H, accent)
    rect(s, W - Inches(3.2), Inches(-1.2), Inches(4.2), Inches(4.2), accent, MSO_SHAPE.OVAL)
    rect(s, W - Inches(2.2), Inches(5.3), Inches(3.0), Inches(3.0), accent, MSO_SHAPE.OVAL)
    rect(s, Inches(1.2), Inches(4.4), Inches(2.2), Inches(0.08), accent)
    text(s, Inches(1.2), Inches(1.7), Inches(9.0), Inches(2.6), data.get("title", ""), 46,
         on_dark, True, anchor=MSO_ANCHOR.BOTTOM)
    text(s, Inches(1.2), Inches(4.65), Inches(9.0), Inches(1.2), data.get("subtitle", ""), 22,
         "D1D5DB")

    # --- Asosiy slaydlar ---
    for i, sl in enumerate(slides, start=2):
        s = prs.slides.add_slide(blank)
        set_bg(s, light)
        rect(s, 0, 0, W, Inches(1.35), dark)
        rect(s, 0, Inches(1.35), W, Inches(0.07), accent)
        badge = rect(s, Inches(0.55), Inches(0.35), Inches(0.65), Inches(0.65), accent,
                     MSO_SHAPE.OVAL)
        tf = badge.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(i - 1)
        r.font.size, r.font.bold = Pt(18), True
        r.font.color.rgb = rgb("FFFFFF")
        text(s, Inches(1.45), Inches(0.1), Inches(11.4), Inches(1.15), sl.get("title", ""), 30,
             on_dark, True, anchor=MSO_ANCHOR.MIDDLE)

        img = images[i - 2] if images and images[i - 2] else None
        text_w = Inches(6.9) if img else Inches(11.5)
        if img:
            rect(s, Inches(8.4), Inches(1.8), Inches(IMG_W + 0.2), Inches(IMG_H + 0.2), accent)
            s.shapes.add_picture(io.BytesIO(img), Inches(8.5), Inches(1.9),
                                 Inches(IMG_W), Inches(IMG_H))

        bullets = [b for b in sl.get("bullets", []) if b][:6] or [""]
        n = len(bullets)
        chars = sum(len(b) for b in bullets)
        size = 22 if chars < 260 else 20 if chars < 340 else 18
        if img:
            size -= 2
        top, avail = Inches(1.85), Inches(4.9)
        step = int(avail / n)
        for j, b in enumerate(bullets):
            y = top + step * j
            rect(s, Inches(0.8), y + Inches(0.13), Inches(0.2), Inches(0.2), accent,
                 MSO_SHAPE.OVAL)
            text(s, Inches(1.2), y, text_w, step, b, size, txt)

        text(s, Inches(11.6), Inches(6.95), Inches(1.4), Inches(0.4), f"{i} / {total}", 12,
             "9CA3AF", align=PP_ALIGN.RIGHT)
        if sl.get("notes"):
            s.notes_slide.notes_text_frame.text = sl["notes"]

    # --- Yakuniy ---
    s = prs.slides.add_slide(blank)
    set_bg(s, dark)
    rect(s, W - Inches(0.35), 0, Inches(0.35), H, accent)
    rect(s, Inches(-1.5), Inches(4.8), Inches(4.0), Inches(4.0), accent, MSO_SHAPE.OVAL)
    t1, t2 = THANKS[lang]
    text(s, 0, Inches(2.5), W, Inches(1.6), t1, 44, on_dark, True, PP_ALIGN.CENTER,
         MSO_ANCHOR.MIDDLE)
    text(s, 0, Inches(4.2), W, Inches(0.8), t2, 22, "D1D5DB", align=PP_ALIGN.CENTER)

    prs.save(path)


# =================================== HANDLERS ===================================
async def ask_topic(m: Message, state: FSMContext):
    await state.clear()
    await state.set_state(Form.topic)
    await m.answer(
        "👋 Salom! Men sizga professional taqdimot tayyorlab beraman.\n\n"
        "✍️ Taqdimot mavzusini yozing, masalan:\n"
        "• <i>Global isish</i>\n• <i>Sun'iy intellekt kelajagi</i>\n• <i>Amir Temur</i>",
        parse_mode="HTML",
    )


@router.message(CommandStart())
@router.message(Command("new"))
async def cmd_start(m: Message, state: FSMContext):
    await ask_topic(m, state)


@router.message(Command("cancel"))
async def cmd_cancel(m: Message, state: FSMContext):
    await state.clear()
    busy_users.discard(m.from_user.id)
    await m.answer("❌ Bekor qilindi. Yangi taqdimot uchun /new ni bosing.")


@router.message(Command("help"))
async def cmd_help(m: Message):
    await m.answer(
        "ℹ️ <b>Qanday ishlaydi?</b>\n"
        "1️⃣ Mavzuni yozasiz\n2️⃣ Slayd sonini tanlaysiz\n3️⃣ Tilni tanlaysiz\n"
        "4️⃣ Dizaynni tanlaysiz\n5️⃣ Rasm kerakmi — tanlaysiz\n"
        "6️⃣ Tayyor .pptx faylni olasiz\n\n"
        "/new — yangi taqdimot\n/cancel — bekor qilish",
        parse_mode="HTML",
    )


async def take_topic(m: Message, state: FSMContext):
    topic = (m.text or "").strip()
    if len(topic) < 3 or len(topic) > 200:
        await m.answer("⚠️ Mavzu 3 dan 200 belgigacha bo'lishi kerak. Qayta yozing.")
        return
    await state.update_data(topic=topic)
    await m.answer(f"📝 Mavzu: <b>{html.escape(topic)}</b>\n\n📊 Nechta slayd kerak?",
                   parse_mode="HTML", reply_markup=count_kb())


@router.message(Form.topic, F.text)
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
    await c.message.edit_text(f"🌐 Til: <b>{LANGS[lang][0]}</b>\n\n🎨 Dizaynni tanlang:",
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
        f"🎨 Dizayn: <b>{THEMES[theme][0]}</b>\n\n🖼 Slaydlarga rasm qo'shilsinmi?",
        parse_mode="HTML", reply_markup=images_kb())
    await c.answer()


@router.callback_query(F.data.startswith("img:"))
async def got_images(c: CallbackQuery, state: FSMContext):
    uid = c.from_user.id
    d = await state.get_data()
    if not all(k in d for k in ("topic", "count", "lang", "theme")):
        await c.answer("Avval /new ni bosing", show_alert=True)
        return
    if uid in busy_users:
        await c.answer("⏳ Oldingi taqdimot tayyorlanmoqda, kuting...", show_alert=True)
        return

    want_images = c.data.endswith("yes")
    topic, count, lang, theme = d["topic"], d["count"], d["lang"], d["theme"]
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
        await asyncio.to_thread(build_pptx, content, theme, lang, path, images)

        safe = re.sub(r"[^\w\-]+", "_", topic)[:40] or "taqdimot"
        extra = f"\n🖼 Rasmlar: {found} ta" if want_images else ""
        await c.message.answer_document(
            FSInputFile(path, filename=f"{safe}.pptx"),
            caption=(f"✅ Tayyor! 📊 {count} slayd\n🎯 Mavzu: {topic}{extra}\n"
                     "🗣 Har slaydda ma'ruzachi uchun izohlar (Notes) bor.\n\n"
                     "Yangi taqdimot: /new"),
        )
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


@router.message(F.text)
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
    bot = Bot(BOT_TOKEN)
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    await bot.delete_webhook(drop_pending_updates=False)  # eski xabarlar saqlansin
    await notify_missed(bot)
    log.info("Bot ishga tushdi")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
