"""
📊 AI Slayd Bot
Mavzu -> slayd soni -> dizayn -> tayyor .pptx fayl

O'rnatish:
    pip install aiogram python-pptx anthropic aiohttp pillow

Ishga tushirish:
    export BOT_TOKEN="telegram_bot_token"
    export ANTHROPIC_API_KEY="sk-ant-..."
    export PEXELS_API_KEY="..."   # ixtiyoriy, rasm qo'shish uchun
    python bot.py
"""
import asyncio
import io
import json
import logging
import os
import re
import tempfile

import aiohttp
from PIL import Image
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import CallbackQuery, FSInputFile, Message
from aiogram.utils.keyboard import InlineKeyboardBuilder
from anthropic import AsyncAnthropic
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("slide-bot")

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
MODEL = os.getenv("AI_MODEL", "claude-sonnet-4-6")
DEBUG = os.getenv("DEBUG", "0") == "1"  # 1 bo'lsa xato sababini chatda ko'rsatadi
PEXELS_KEY = os.getenv("PEXELS_API_KEY", "")  # rasmlar uchun (bepul: pexels.com/api)
ai = AsyncAnthropic()  # ANTHROPIC_API_KEY ni o'zi oladi

SLIDE_COUNTS = [5, 8, 10, 12, 15, 18, 20]

# nomi: (to'q fon, aksent, och fon, matn rangi)
THEMES = {
    "blue":   ("🔵 Ko'k",      "0B3D91", "2F80ED", "F4F8FF", "1F2937"),
    "green":  ("🟢 Yashil",    "064E3B", "10B981", "F0FDF4", "1F2937"),
    "purple": ("🟣 Binafsha",  "3B0764", "A855F7", "FAF5FF", "1F2937"),
    "orange": ("🟠 To'q sariq", "7C2D12", "F97316", "FFF7ED", "1F2937"),
    "dark":   ("⚫️ Qora",      "111827", "38BDF8", "1F2937", "F9FAFB"),
}

router = Router()
busy_users: set[int] = set()


class Form(StatesGroup):
    topic = State()


# ---------------------------------------------------------------- keyboards
def count_kb():
    kb = InlineKeyboardBuilder()
    for n in SLIDE_COUNTS:
        kb.button(text=f"📊 {n} slayd", callback_data=f"count:{n}")
    kb.adjust(2, 2, 2, 1)
    return kb.as_markup()


def theme_kb():
    kb = InlineKeyboardBuilder()
    for key, (name, *_rest) in THEMES.items():
        kb.button(text=name, callback_data=f"theme:{key}")
    kb.adjust(2)
    return kb.as_markup()


def images_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="🖼 Ha, rasm qo'shilsin", callback_data="img:yes")
    kb.button(text="📄 Yo'q, rasmsiz", callback_data="img:no")
    kb.adjust(1)
    return kb.as_markup()


# ---------------------------------------------------------------- AI
def extract_json(text: str) -> dict:
    text = re.sub(r"```(?:json)?", "", text)
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("AI JSON qaytarmadi")
    return json.loads(text[start : end + 1])


async def generate_content(topic: str, total: int) -> dict:
    content_slides = total - 2  # birinchi (titul) va oxirgi (rahmat) slaydlar alohida
    prompt = f"""Mavzu: "{topic}"

Shu mavzuda professional taqdimot uchun matn tayyorla. O'zbek tilida (lotin yozuvi) yoz.
Aynan {content_slides} ta asosiy slayd bo'lsin.

FAQAT quyidagi JSON formatda javob ber, boshqa hech narsa yozma:
{{
  "title": "Taqdimot sarlavhasi",
  "subtitle": "Qisqa tagsarlavha",
  "slides": [
    {{"title": "Slayd sarlavhasi", "bullets": ["fikr 1", "fikr 2", "fikr 3", "fikr 4"], "image_query": "2-3 english words for stock photo search"}}
  ]
}}

Qoidalar:
- Har slaydda 3-5 ta band, har biri 8-18 so'zdan iborat, mazmunli va aniq.
- Slaydlar mantiqiy ketma-ketlikda: kirish, asosiy qismlar, misollar/statistika, xulosa.
- Sarlavhalar qisqa (6 so'zgacha)."""
    resp = await ai.messages.create(
        model=MODEL,
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )
    data = extract_json(resp.content[0].text)
    slides = data.get("slides", [])
    if not slides:
        raise ValueError("Slaydlar bo'sh")
    return data


# ---------------------------------------------------------------- IMAGES
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


async def fetch_image(session: aiohttp.ClientSession, query: str):
    if not query:
        return None
    try:
        async with session.get(
            "https://api.pexels.com/v1/search",
            params={"query": query, "per_page": 1, "orientation": "portrait"},
            headers={"Authorization": PEXELS_KEY},
        ) as r:
            photos = (await r.json()).get("photos") or []
        if not photos:
            return None
        async with session.get(photos[0]["src"]["large"]) as r:
            raw = await r.read()
        return await asyncio.to_thread(crop_image, raw)
    except Exception as e:
        log.warning("Rasm olinmadi (%s): %s", query, e)
        return None


async def fetch_all_images(slides: list) -> list:
    timeout = aiohttp.ClientTimeout(total=20)
    async with aiohttp.ClientSession(timeout=timeout) as session:
        return await asyncio.gather(
            *[fetch_image(session, sl.get("image_query", "")) for sl in slides]
        )


# ---------------------------------------------------------------- PPTX
def rgb(hex_: str) -> RGBColor:
    return RGBColor.from_string(hex_)


def set_bg(slide, color: str):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def add_rect(slide, x, y, w, h, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(color)
    shp.line.fill.background()
    return shp


def add_text(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = "Calibri"
    r.font.color.rgb = rgb(color)
    return tb


def build_pptx(data: dict, theme_key: str, path: str, images: list | None = None):
    _, dark, accent, light, text = THEMES[theme_key]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    W, H = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]
    slides = data["slides"]
    total = len(slides) + 2

    # 1) Titul slayd
    s = prs.slides.add_slide(blank)
    set_bg(s, dark)
    add_rect(s, 0, 0, Inches(0.35), H, accent)
    add_rect(s, Inches(1.2), Inches(4.35), Inches(2.2), Inches(0.08), accent)
    add_text(s, Inches(1.2), Inches(1.9), Inches(10.8), Inches(2.4),
             data.get("title", "Taqdimot"), 48, "FFFFFF", True, anchor=MSO_ANCHOR.BOTTOM)
    add_text(s, Inches(1.2), Inches(4.6), Inches(10.8), Inches(1.2),
             data.get("subtitle", ""), 22, "D1D5DB")

    # 2) Asosiy slaydlar
    for i, sl in enumerate(slides, start=2):
        s = prs.slides.add_slide(blank)
        set_bg(s, light)
        add_rect(s, 0, 0, W, Inches(1.35), dark)
        add_rect(s, 0, Inches(1.35), W, Inches(0.07), accent)
        add_text(s, Inches(0.7), Inches(0.1), Inches(11.9), Inches(1.15),
                 sl.get("title", ""), 32, "FFFFFF", True, anchor=MSO_ANCHOR.MIDDLE)

        img = images[i - 2] if images and images[i - 2] else None
        text_w = Inches(6.9) if img else Inches(11.3)
        if img:
            s.shapes.add_picture(io.BytesIO(img), Inches(8.5), Inches(1.9),
                                 Inches(IMG_W), Inches(IMG_H))

        bullets = sl.get("bullets", [])[:6]
        n = max(len(bullets), 1)
        top, avail = Inches(1.9), Inches(4.8)
        step = int(avail / n)
        size = (20 if img else 22) if n <= 4 else (17 if img else 19)
        for j, b in enumerate(bullets):
            y = top + step * j
            dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), y + Inches(0.12),
                                     Inches(0.22), Inches(0.22))
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(accent)
            dot.line.fill.background()
            add_text(s, Inches(1.25), y, text_w, step, b, size, text)

        add_text(s, Inches(11.6), Inches(6.95), Inches(1.4), Inches(0.4),
                 f"{i} / {total}", 12, "9CA3AF", align=PP_ALIGN.RIGHT)

    # 3) Rahmat slayd
    s = prs.slides.add_slide(blank)
    set_bg(s, dark)
    add_rect(s, W - Inches(0.35), 0, Inches(0.35), H, accent)
    add_text(s, 0, Inches(2.6), W, Inches(1.6), "E'tiboringiz uchun rahmat!", 46,
             "FFFFFF", True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    add_text(s, 0, Inches(4.2), W, Inches(0.8), "Savollaringiz bormi?", 22,
             "D1D5DB", align=PP_ALIGN.CENTER)

    prs.save(path)


# ---------------------------------------------------------------- handlers
@router.message(CommandStart())
@router.message(Command("new"))
async def cmd_start(m: Message, state: FSMContext):
    await state.clear()
    await state.set_state(Form.topic)
    await m.answer(
        "👋 Salom! Men sizga professional taqdimot tayyorlab beraman.\n\n"
        "✍️ Taqdimot mavzusini yozing, masalan:\n"
        "• <i>Global isish</i>\n• <i>Sun'iy intellekt kelajagi</i>",
        parse_mode="HTML",
    )


@router.message(Command("cancel"))
async def cmd_cancel(m: Message, state: FSMContext):
    await state.clear()
    busy_users.discard(m.from_user.id)
    await m.answer("❌ Bekor qilindi. Yangi taqdimot uchun /new ni bosing.")


@router.message(Command("help"))
async def cmd_help(m: Message):
    await m.answer(
        "ℹ️ <b>Qanday ishlaydi?</b>\n"
        "1️⃣ Mavzuni yozasiz\n2️⃣ Slayd sonini tanlaysiz\n3️⃣ Dizaynni tanlaysiz\n"
        "4️⃣ Tayyor .pptx faylni olasiz\n\n"
        "/new — yangi taqdimot\n/cancel — bekor qilish",
        parse_mode="HTML",
    )


@router.message(Form.topic, F.text)
async def got_topic(m: Message, state: FSMContext):
    topic = m.text.strip()
    if len(topic) < 3 or len(topic) > 200:
        await m.answer("⚠️ Mavzu 3 dan 200 belgigacha bo'lishi kerak. Qayta yozing.")
        return
    await state.update_data(topic=topic)
    await m.answer(f"📝 Mavzu: <b>{topic}</b>\n\n📊 Nechta slayd kerak?",
                   parse_mode="HTML", reply_markup=count_kb())


@router.callback_query(F.data.startswith("count:"))
async def got_count(c: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    if "topic" not in data:
        await c.answer("Avval mavzuni yozing: /new", show_alert=True)
        return
    n = int(c.data.split(":")[1])
    await state.update_data(count=n)
    await c.message.edit_text(f"📊 Slaydlar soni: <b>{n}</b>\n\n🎨 Dizaynni tanlang:",
                              parse_mode="HTML", reply_markup=theme_kb())
    await c.answer()


@router.callback_query(F.data.startswith("theme:"))
async def got_theme(c: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    if "topic" not in data or "count" not in data:
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
    data = await state.get_data()
    if not all(k in data for k in ("topic", "count", "theme")):
        await c.answer("Avval /new ni bosing", show_alert=True)
        return
    if uid in busy_users:
        await c.answer("⏳ Oldingi taqdimot tayyorlanmoqda, kuting...", show_alert=True)
        return

    want_images = c.data.endswith("yes")
    topic, count, theme = data["topic"], data["count"], data["theme"]
    no_key = want_images and not PEXELS_KEY
    if no_key:
        want_images = False

    busy_users.add(uid)
    await c.answer()
    status = await c.message.edit_text(
        f"⏳ <b>{topic}</b> bo'yicha {count} slaydli taqdimot tayyorlanmoqda...\n"
        "Bu 20-90 soniya olishi mumkin.", parse_mode="HTML")

    path = None
    try:
        content = None
        for attempt in range(3):
            try:
                content = await generate_content(topic, count)
                break
            except Exception as e:
                log.warning("AI urinish %s xato: %s", attempt + 1, e)
                if attempt == 2:
                    raise
                await asyncio.sleep(2)

        images = None
        if want_images:
            await status.edit_text("🖼 Rasmlar qidirilmoqda...")
            images = await fetch_all_images(content["slides"])

        fd, path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)
        await asyncio.to_thread(build_pptx, content, theme, path, images)

        safe = re.sub(r"[^\w\-]+", "_", topic)[:40] or "taqdimot"
        note = "\n⚠️ Rasm xizmati sozlanmagan, rasmsiz tayyorlandi." if no_key else ""
        await c.message.answer_document(
            FSInputFile(path, filename=f"{safe}.pptx"),
            caption=f"✅ Tayyor! 📊 {count} slayd\n🎯 Mavzu: {topic}{note}\n\nYangi taqdimot: /new",
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


@router.message()
async def fallback(m: Message, state: FSMContext):
    await state.set_state(Form.topic)
    await got_topic(m, state) if m.text else await m.answer("✍️ Mavzuni matn ko'rinishida yozing.")


async def main():
    if not BOT_TOKEN:
        raise SystemExit("BOT_TOKEN o'rnatilmagan")
    bot = Bot(BOT_TOKEN)
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
