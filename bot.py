import os
import io
import re
import json
import asyncio
import logging
from datetime import date

from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.types import (
    Message,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    BufferedInputFile,
)
from anthropic import AsyncAnthropic
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.environ["BOT_TOKEN"]
MODEL = os.getenv("MODEL", "claude-sonnet-5-5")
FREE_LIMIT = int(os.getenv("FREE_LIMIT", "2"))  # kuniga bepul taqdimotlar soni

client = AsyncAnthropic()  # ANTHROPIC_API_KEY ni o'zi oladi
bot = Bot(BOT_TOKEN)
dp = Dispatcher()

topics = {}  # user_id -> mavzu
usage = {}  # (user_id, sana) -> son


def used_today(uid):
    return usage.get((uid, date.today().isoformat()), 0)


def add_usage(uid):
    key = (uid, date.today().isoformat())
    usage[key] = usage.get(key, 0) + 1


async def generate_slides(topic, n):
    prompt = (
        f"Mavzu: {topic}\n"
        f"Shu mavzuda {n} ta slayddan iborat taqdimot tayyorla. "
        "Taqdimot tilini mavzu yozilgan tilda qil. "
        "Har bir slaydda qisqa sarlavha va 3-5 ta qisqa band bo'lsin. "
        "FAQAT JSON qaytar, boshqa hech narsa yozma:\n"
        '{"title": "...", "slides": [{"title": "...", "points": ["...", "..."]}]}'
    )
    resp = await client.messages.create(
        model=MODEL,
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    start, end = text.find("{"), text.rfind("}")
    return json.loads(text[start : end + 1])


def build_pptx(data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    dark = RGBColor(0x1F, 0x2A, 0x44)
    accent = RGBColor(0xE8, 0x6A, 0x33)
    light = RGBColor(0xF7, 0xF7, 0xF2)
    white = RGBColor(0xFF, 0xFF, 0xFF)

    def set_bg(slide, color):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    # Birinchi slayd
    s = prs.slides.add_slide(blank)
    set_bg(s, dark)
    box = s.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.3), Inches(2.5))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data["title"]
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = white

    # Qolgan slaydlar
    for sl in data["slides"]:
        s = prs.slides.add_slide(blank)
        set_bg(s, light)

        bar = s.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.35), prs.slide_height
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()

        tbox = s.shapes.add_textbox(Inches(1), Inches(0.6), Inches(11.3), Inches(1.2))
        ttf = tbox.text_frame
        ttf.word_wrap = True
        tp = ttf.paragraphs[0]
        tp.text = sl["title"]
        tp.font.size = Pt(34)
        tp.font.bold = True
        tp.font.color.rgb = dark

        bbox = s.shapes.add_textbox(Inches(1), Inches(2), Inches(11.3), Inches(4.8))
        btf = bbox.text_frame
        btf.word_wrap = True
        for i, point in enumerate(sl["points"]):
            para = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
            para.text = "• " + point
            para.font.size = Pt(24)
            para.font.color.rgb = dark
            para.space_after = Pt(14)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


@dp.message(CommandStart())
async def start(m: Message):
    await m.answer(
        "Salom! Men taqdimot tayyorlab beraman.\n"
        "Menga taqdimot mavzusini yozing, masalan: Global isish"
    )


@dp.message(F.text)
async def got_topic(m: Message):
    topics[m.from_user.id] = m.text.strip()
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="5 slayd", callback_data="n:5"),
                InlineKeyboardButton(text="8 slayd", callback_data="n:8"),
                InlineKeyboardButton(text="10 slayd", callback_data="n:10"),
            ]
        ]
    )
    await m.answer("Nechta slayd kerak?", reply_markup=kb)


@dp.callback_query(F.data.startswith("n:"))
async def make(c: CallbackQuery):
    uid = c.from_user.id
    topic = topics.get(uid)
    if not topic:
        await c.answer("Avval mavzu yozing", show_alert=True)
        return
    if used_today(uid) >= FREE_LIMIT:
        await c.message.answer(
            f"Bugungi bepul limit ({FREE_LIMIT} ta) tugadi. Ertaga qaytib keling."
        )
        await c.answer()
        return

    n = int(c.data.split(":")[1])
    await c.answer()
    wait = await c.message.answer("Taqdimot tayyorlanmoqda, 20-40 soniya kuting...")
    try:
        data = await generate_slides(topic, n)
        file_bytes = await asyncio.to_thread(build_pptx, data)
    except Exception:
        logging.exception("Xatolik")
        await wait.edit_text("Xatolik bo'ldi, qayta urinib ko'ring.")
        return

    add_usage(uid)
    name = re.sub(r"[^\w\- ]", "", data["title"])[:40].strip() or "taqdimot"
    await c.message.answer_document(
        BufferedInputFile(file_bytes, filename=f"{name}.pptx"),
        caption="Tayyor! Fayl PowerPoint yoki Google Slides'da ochiladi.",
    )
    await wait.delete()


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
