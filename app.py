import io
import math
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor

st.set_page_config(page_title="Kaziony PPT Maker", layout="wide")
st.title("Kaziony PPT Maker")
st.caption("Generates the Kaziony explainer deck and lets you download it.")

# ---- Theme / constants ----
W, H = Inches(13.333), Inches(7.5)

COL = {
    "bg": RGBColor(10, 10, 14),
    "panel": RGBColor(18, 18, 24),
    "panel2": RGBColor(24, 24, 32),
    "gold": RGBColor(212, 175, 55),
    "white": RGBColor(245, 245, 247),
    "muted": RGBColor(170, 170, 178),
    "line": RGBColor(60, 60, 72),
    "ok": RGBColor(68, 214, 144),
    "warn": RGBColor(255, 204, 0),
}

FONT_EN = "Segoe UI"
FONT_AR = "Segoe UI Arabic"

# ---- Slide helpers ----
def add_bg(slide):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    r.fill.solid()
    r.fill.fore_color.rgb = COL["bg"]
    r.line.fill.background()

def add_header(slide, title_en, title_ar, app_name="Kaziony"):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Inches(0.85))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COL["panel"]
    bar.line.fill.background()

    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.82), W, Inches(0.03))
    accent.fill.solid()
    accent.fill.fore_color.rgb = COL["gold"]
    accent.line.fill.background()

    tb_en = slide.shapes.add_textbox(Inches(0.65), Inches(0.18), Inches(6.2), Inches(0.55))
    tf = tb_en.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title_en
    p.font.name = FONT_EN
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COL["white"]
    p.alignment = PP_ALIGN.LEFT

    tb_ar = slide.shapes.add_textbox(Inches(6.9), Inches(0.18), Inches(5.8), Inches(0.55))
    tf = tb_ar.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title_ar
    p.font.name = FONT_AR
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COL["white"]
    p.alignment = PP_ALIGN.RIGHT

    brand = slide.shapes.add_textbox(Inches(11.7), Inches(0.22), Inches(1.55), Inches(0.4))
    tf = brand.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = app_name
    p.font.name = FONT_EN
    p.font.size = Pt(14)
    p.font.color.rgb = COL["muted"]
    p.alignment = PP_ALIGN.RIGHT

def add_footer(slide, idx, total):
    tb = slide.shapes.add_textbox(Inches(0.65), Inches(7.18), Inches(12.0), Inches(0.3))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"{idx}/{total}"
    p.font.name = FONT_EN
    p.font.size = Pt(12)
    p.font.color.rgb = COL["muted"]
    p.alignment = PP_ALIGN.RIGHT

def panel(slide, x, y, w, h, accent=None, lw=1):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = COL["panel"]
    s.line.color.rgb = accent if accent else COL["line"]
    s.line.width = Pt(lw)
    return s

def chip(slide, x, y, text, font=FONT_EN, fg=COL["bg"], bgc=COL["gold"], w=Inches(2.2), h=Inches(0.42), size=14):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = bgc
    s.line.fill.background()
    tf = s.text_frame; tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font
    p.font.size = Pt(size)
    p.font.bold = True
    p.font.color.rgb = fg
    p.alignment = PP_ALIGN.CENTER
    return s

def add_bilingual_cards(slide, items, x, y, w, h, cols=2):
    # items: (en_title, ar_title, en_desc, ar_desc, accent_rgb)
    rows = math.ceil(len(items) / cols)
    gapx, gapy = Inches(0.35), Inches(0.28)
    card_w = (w - gapx*(cols-1)) / cols
    card_h = (h - gapy*(rows-1)) / rows
    for i, it in enumerate(items):
        r = i // cols
        c = i % cols
        cx = x + c * (card_w + gapx)
        cy = y + r * (card_h + gapy)
        accent = it[4] if len(it) > 4 and it[4] is not None else COL["gold"]
        panel(slide, cx, cy, card_w, card_h, accent=accent, lw=2)
        tb = slide.shapes.add_textbox(cx+Inches(0.28), cy+Inches(0.18), card_w-Inches(0.56), card_h-Inches(0.32))
        tf = tb.text_frame; tf.clear(); tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = it[0]
        p.font.name = FONT_EN; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = accent

        p2 = tf.add_paragraph()
        p2.text = it[1]
        p2.font.name = FONT_AR; p2.font.size = Pt(16); p2.font.bold = True; p2.font.color.rgb = COL["white"]
        p2.alignment = PP_ALIGN.RIGHT

        p3 = tf.add_paragraph()
        p3.text = it[2]
        p3.font.name = FONT_EN; p3.font.size = Pt(13); p3.font.color.rgb = COL["muted"]

        p4 = tf.add_paragraph()
        p4.text = it[3]
        p4.font.name = FONT_AR; p4.font.size = Pt(13); p4.font.color.rgb = COL["muted"]
        p4.alignment = PP_ALIGN.RIGHT

def add_three_step(slide):
    items = [
        ("Order", "الطلب", "Customer selects merchant → items → checkout", "العميل يختار المتجر → يضيف المنتجات → تأكيد", COL["gold"]),
        ("Credit", "كريديت", "Driver accepts using 1 credit (deducted instantly)", "السائق يقبل بكريديت واحد (خصم فوري)", COL["ok"]),
        ("Cash", "كاش", "Deliver → collect cash (items + delivery + tip)", "تسليم → تحصيل الكاش (قيمة الطلب + التوصيل + التيب)", COL["warn"]),
    ]
    xs = [Inches(0.9), Inches(4.8), Inches(8.7)]
    for i, it in enumerate(items):
        panel(slide, xs[i], Inches(2.0), Inches(3.8), Inches(2.6), accent=it[4], lw=2)
        tb = slide.shapes.add_textbox(xs[i]+Inches(0.3), Inches(2.2), Inches(3.2), Inches(2.2))
        tf = tb.text_frame; tf.clear(); tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"{it[0]} | {it[1]}"
        p.font.name = FONT_EN; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = it[4]
        p2 = tf.add_paragraph()
        p2.text = it[2]
        p2.font.name = FONT_EN; p2.font.size = Pt(14); p2.font.color.rgb = COL["white"]
        p3 = tf.add_paragraph()
        p3.text = it[3]
        p3.font.name = FONT_AR; p3.font.size = Pt(14); p3.font.color.rgb = COL["white"]; p3.alignment = PP_ALIGN.RIGHT

        if i < 2:
            arr = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, xs[i]+Inches(3.9), Inches(3.0), Inches(0.8), Inches(0.75))
            arr.fill.solid(); arr.fill.fore_color.rgb = COL["gold"]; arr.line.fill.background()

def build_deck(app_name="Kaziony", city_en="Tripoli", city_ar="طرابلس"):
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    total = 6
    idx = 1

    # 1 Cover
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s)
    title = s.shapes.add_textbox(Inches(0.9), Inches(2.1), Inches(11.6), Inches(2.2))
    tf = title.text_frame; tf.clear(); tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = f"{app_name} | كازيوني"; p.font.name = FONT_EN; p.font.size = Pt(56); p.font.bold = True; p.font.color.rgb = COL["gold"]
    p2 = tf.add_paragraph(); p2.text = "Cash-first delivery + driver-credit acceptance"; p2.font.name = FONT_EN; p2.font.size = Pt(20); p2.font.color.rgb = COL["white"]
    p3 = tf.add_paragraph(); p3.text = "توصيل كاش + قبول الطلب بكريديت للسائق"; p3.font.name = FONT_AR; p3.font.size = Pt(20); p3.font.color.rgb = COL["white"]; p3.alignment = PP_ALIGN.RIGHT
    chip(s, Inches(0.9), Inches(6.2), f"Launch: {city_en} / {city_ar}", FONT_EN, COL["bg"], COL["gold"], w=Inches(3.8), h=Inches(0.45))
    add_footer(s, idx, total); idx += 1

    # 2 What it is
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s); add_header(s, "What is Kaziony", "ما هو كازيوني", app_name)
    items = [
        ("Restaurants", "مطاعم", "Order food from nearby merchants", "طلب الطعام من المطاعم القريبة", COL["gold"]),
        ("Groceries", "بقالة", "Daily items and essentials", "احتياجات يومية وأساسية", COL["gold"]),
        ("Pharmacy", "صيدلية", "Fast pharmacy deliveries", "توصيل الصيدلية بسرعة", COL["gold"]),
        ("Scheduled", "مجدول", "Choose a delivery time window", "حدد نافذة زمنية للتوصيل", COL["gold"]),
    ]
    add_bilingual_cards(s, items, Inches(0.9), Inches(1.55), Inches(12.0), Inches(5.7), cols=2)
    add_footer(s, idx, total); idx += 1

    # 3 How it works
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s); add_header(s, "How it works (3 steps)", "كيف يعمل (٣ خطوات)", app_name)
    add_three_step(s)
    add_footer(s, idx, total); idx += 1

    # 4 Credits
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s); add_header(s, "Credits (core idea)", "الكريديت (الفكرة الأساسية)", app_name)
    rules = [
        ("Accept requires credit", "القبول يحتاج كريديت", "Driver must have ≥ 1 credit", "السائق لازم عنده كريديت 1 أو أكثر", COL["gold"]),
        ("1 credit per accepted order", "كريديت لكل قبول", "Deducted instantly on Accept", "خصم فوري عند القبول", COL["ok"]),
        ("Refund on cancellation", "استرجاع عند الإلغاء", "Always refund credit if order cancels", "يرجع الكريديت عند إلغاء الطلب", COL["warn"]),
        ("Why", "السبب", "Reduces spam accepts + improves reliability", "يقلل العشوائية ويزيد الموثوقية", COL["gold"]),
    ]
    add_bilingual_cards(s, rules, Inches(0.9), Inches(1.55), Inches(12.0), Inches(5.7), cols=2)
    add_footer(s, idx, total); idx += 1

    # 5 Money flow
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s); add_header(s, "Money flow (cash)", "مسار الأموال (كاش)", app_name)
    cards = [
        ("Pickup", "استلام", "Driver pays merchant for items", "السائق يدفع قيمة الطلب للمتجر", COL["gold"]),
        ("Delivery", "تسليم", "Customer pays: items + delivery fee + tip", "العميل يدفع: الطلب + التوصيل + التيب", COL["gold"]),
        ("Driver keeps", "دخل السائق", "Delivery fee + tips (items money is reimbursement)", "السائق يأخذ رسوم التوصيل + التيب (قيمة الطلب استرجاع)", COL["ok"]),
        ("Fees upfront", "الرسوم واضحة", "Distance-based fee shown before order", "الرسوم حسب المسافة وتظهر قبل التأكيد", COL["warn"]),
    ]
    add_bilingual_cards(s, cards, Inches(0.9), Inches(1.55), Inches(12.0), Inches(5.7), cols=2)
    add_footer(s, idx, total); idx += 1

    # 6 Ops
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s); add_header(s, "Operations", "التشغيل", app_name)
    ops = [
        ("Hybrid dispatch", "إسناد هجين", "Auto-offer + dispatcher override", "عرض تلقائي + تدخل موزّع", COL["gold"]),
        ("Live tracking", "تتبع مباشر", "Customer sees driver on map", "العميل يشوف السائق على الخريطة", COL["gold"]),
        ("In-app support", "دعم داخل التطبيق", "Chat-based support", "دعم عبر محادثة", COL["gold"]),
        ("Stats", "إحصائيات", "Orders, completion, active drivers", "طلبات، إكمال، سائقين متصلين", COL["gold"]),
    ]
    add_bilingual_cards(s, ops, Inches(0.9), Inches(1.55), Inches(12.0), Inches(5.7), cols=2)
    add_footer(s, idx, total); idx += 1

    return prs

# ---- UI ----
col1, col2, col3 = st.columns(3)
with col1:
    app_name = st.text_input("App name", value="Kaziony")
with col2:
    city_en = st.text_input("Launch city (EN)", value="Tripoli")
with col3:
    city_ar = st.text_input("Launch city (AR)", value="طرابلس")

st.write("")
if st.button("Generate PPT"):
    prs = build_deck(app_name=app_name, city_en=city_en, city_ar=city_ar)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    st.success("PPT generated.")
    st.download_button(
        "Download PPTX",
        data=buf.getvalue(),
        file_name=f"{app_name}_Explainer.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
else:
    st.info("Click Generate PPT to create the deck.")
