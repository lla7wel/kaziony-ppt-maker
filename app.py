import io
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor

PPT_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# Premium dark theme (black + gold)
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

# Fonts (Modern choice). If Inter/Cairo not installed on viewer PC, PowerPoint will fallback.
FONT_EN = "Inter"
FONT_AR = "Cairo"

def _bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COL["bg"]
    bg.line.fill.background()

def _header(slide, title_en, title_ar, app_name="Kaziony"):
    # Header bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.85))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COL["panel"]
    bar.line.fill.background()

    # Gold accent line
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0.82), Inches(13.333), Inches(0.03))
    accent.fill.solid()
    accent.fill.fore_color.rgb = COL["gold"]
    accent.line.fill.background()

    # EN title
    tb_en = slide.shapes.add_textbox(Inches(0.6), Inches(0.18), Inches(6.2), Inches(0.55))
    tf = tb_en.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title_en
    p.font.name = FONT_EN
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COL["white"]
    p.alignment = PP_ALIGN.LEFT

    # AR title
    tb_ar = slide.shapes.add_textbox(Inches(6.6), Inches(0.18), Inches(6.1), Inches(0.55))
    tf = tb_ar.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title_ar
    p.font.name = FONT_AR
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COL["white"]
    p.alignment = PP_ALIGN.RIGHT

    # Brand (small)
    brand = slide.shapes.add_textbox(Inches(11.4), Inches(0.22), Inches(1.8), Inches(0.4))
    tf = brand.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = app_name
    p.font.name = FONT_EN
    p.font.size = Pt(14)
    p.font.color.rgb = COL["muted"]
    p.alignment = PP_ALIGN.RIGHT

def _footer(slide, idx, total):
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(7.2), Inches(12.2), Inches(0.3))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"{idx}/{total}"
    p.font.name = FONT_EN
    p.font.size = Pt(12)
    p.font.color.rgb = COL["muted"]
    p.alignment = PP_ALIGN.RIGHT

def _two_col_bullets(slide, bullets_en, bullets_ar):
    # Panels
    left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.2), Inches(6.2), Inches(5.75))
    right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), Inches(1.2), Inches(6.33), Inches(5.75))
    for panel in (left, right):
        panel.fill.solid()
        panel.fill.fore_color.rgb = COL["panel"]
        panel.line.color.rgb = COL["line"]
        panel.line.width = Pt(1)

    # EN label
    tag_en = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.33), Inches(0.9), Inches(0.35))
    tag_en.fill.solid()
    tag_en.fill.fore_color.rgb = COL["panel2"]
    tag_en.line.color.rgb = COL["line"]
    tf = tag_en.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = "EN"
    p.font.name = FONT_EN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COL["gold"]
    p.alignment = PP_ALIGN.CENTER

    # AR label
    tag_ar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.4), Inches(1.33), Inches(0.9), Inches(0.35))
    tag_ar.fill.solid()
    tag_ar.fill.fore_color.rgb = COL["panel2"]
    tag_ar.line.color.rgb = COL["line"]
    tf = tag_ar.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = "AR"
    p.font.name = FONT_EN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COL["gold"]
    p.alignment = PP_ALIGN.CENTER

    # Textboxes inside panels
    tb_en = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(5.6), Inches(5.0))
    tb_ar = slide.shapes.add_textbox(Inches(7.4), Inches(1.8), Inches(5.8), Inches(5.0))

    # Font sizing based on bullet count
    fs = 20 if max(len(bullets_en), len(bullets_ar)) <= 5 else 18

    tf = tb_en.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, b in enumerate(bullets_en):
        pp = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        pp.text = f"• {b}"
        pp.font.name = FONT_EN
        pp.font.size = Pt(fs)
        pp.font.color.rgb = COL["white"]
        pp.space_after = Pt(8)

    tf = tb_ar.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, b in enumerate(bullets_ar):
        pp = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        pp.text = f"• {b}"
        pp.font.name = FONT_AR
        pp.font.size = Pt(fs)
        pp.font.color.rgb = COL["white"]
        pp.alignment = PP_ALIGN.RIGHT
        pp.space_after = Pt(8)

def _three_step_summary(slide):
    # 3 cards
    cards = [
        ("Order", "طلب", "Customer places order", "العميل يأكد الطلب"),
        ("Credit", "كريديت", "Driver accepts using 1 credit", "السائق يقبل بكريديت واحد"),
        ("Cash", "كاش", "Driver delivers + collects cash", "السائق يسلّم ويحصّل الكاش"),
    ]
    x = [0.8, 4.75, 8.7]
    for i, (t1, t2, s1, s2) in enumerate(cards):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x[i]), Inches(2.1), Inches(3.8), Inches(2.3))
        box.fill.solid()
        box.fill.fore_color.rgb = COL["panel"]
        box.line.color.rgb = COL["gold"]
        box.line.width = Pt(2)

        tf = box.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = f"{t1} | {t2}"
        p.font.name = FONT_EN
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = COL["gold"]
        p.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = s1
        p2.font.name = FONT_EN
        p2.font.size = Pt(16)
        p2.font.color.rgb = COL["white"]
        p2.alignment = PP_ALIGN.CENTER

        p3 = tf.add_paragraph()
        p3.text = s2
        p3.font.name = FONT_AR
        p3.font.size = Pt(16)
        p3.font.color.rgb = COL["white"]
        p3.alignment = PP_ALIGN.CENTER

        if i < 2:
            arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x[i] + 3.95), Inches(2.85), Inches(0.75), Inches(0.8))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = COL["gold"]
            arrow.line.fill.background()

def _flow_diagram(slide, steps_en, steps_ar):
    x0, y, w, h, gap = 0.8, 2.2, 1.95, 1.25, 0.25
    for i in range(len(steps_en)):
        x = x0 + i * (w + gap)
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        box.fill.solid()
        box.fill.fore_color.rgb = COL["panel"]
        box.line.color.rgb = COL["line"]
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = steps_en[i]
        p.font.name = FONT_EN
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COL["gold"]
        p.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = steps_ar[i]
        p2.font.name = FONT_AR
        p2.font.size = Pt(13)
        p2.font.color.rgb = COL["white"]
        p2.alignment = PP_ALIGN.CENTER

        if i < len(steps_en) - 1:
            x1 = x + w
            x2 = x0 + (i + 1) * (w + gap)
            line = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                Inches(x1),
                Inches(y + h / 2),
                Inches(x2),
                Inches(y + h / 2),
            )
            line.line.color.rgb = COL["gold"]
            line.line.width = Pt(2)

def _money_flow(slide):
    # Merchant -> Driver -> Customer diagram
    nodes = [
        ("Merchant", "المتجر", 1.2),
        ("Driver", "السائق", 5.6),
        ("Customer", "العميل", 10.0),
    ]
    y = 2.4
    for label_en, label_ar, x in nodes:
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(2.2), Inches(1.2))
        box.fill.solid()
        box.fill.fore_color.rgb = COL["panel"]
        box.line.color.rgb = COL["line"]
        tf = box.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = label_en
        p.font.name = FONT_EN
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = COL["gold"]
        p.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = label_ar
        p2.font.name = FONT_AR
        p2.font.size = Pt(16)
        p2.font.color.rgb = COL["white"]
        p2.alignment = PP_ALIGN.CENTER

    # Arrows + labels
    def arrow(x1, x2, text_en, text_ar):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y + 0.6), Inches(x2), Inches(y + 0.6))
        line.line.color.rgb = COL["gold"]
        line.line.width = Pt(3)

        tb = slide.shapes.add_textbox(Inches((x1+x2)/2 - 1.2), Inches(y + 1.35), Inches(2.4), Inches(0.7))
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = text_en
        p.font.name = FONT_EN
        p.font.size = Pt(12)
        p.font.color.rgb = COL["white"]
        p.alignment = PP_ALIGN.CENTER
        p2 = tf.add_paragraph()
        p2.text = text_ar
        p2.font.name = FONT_AR
        p2.font.size = Pt(12)
        p2.font.color.rgb = COL["white"]
        p2.alignment = PP_ALIGN.CENTER

    # Driver pays merchant at pickup
    arrow(3.4, 5.6, "Cash for items (pickup)", "كاش قيمة الطلب (عند الاستلام)")
    # Customer pays driver at delivery
    arrow(7.8, 10.0, "Cash: items + delivery + tip", "كاش: قيمة الطلب + التوصيل + التيب")

def _phone_mockups(slide, titles_en, titles_ar, is_driver=False):
    # 4 phones (2x2)
    positions = [(0.9, 1.9), (4.1, 1.9), (7.3, 1.9), (10.5, 1.9)]
    for i in range(4):
        x, y = positions[i]
        phone = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(2.6), Inches(4.9))
        phone.fill.solid()
        phone.fill.fore_color.rgb = COL["panel"]
        phone.line.color.rgb = COL["gold"]
        phone.line.width = Pt(2)

        screen = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x+0.18), Inches(y+0.45), Inches(2.24), Inches(4.2))
        screen.fill.solid()
        screen.fill.fore_color.rgb = COL["panel2"]
        screen.line.fill.background()

        # Small top notch
        notch = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x+1.05), Inches(y+0.25), Inches(0.5), Inches(0.12))
        notch.fill.solid()
        notch.fill.fore_color.rgb = COL["line"]
        notch.line.fill.background()

        # Screen title
        tb = slide.shapes.add_textbox(Inches(x+0.25), Inches(y+0.6), Inches(2.1), Inches(0.6))
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = titles_en[i]
        p.font.name = FONT_EN
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COL["gold"]
        p.alignment = PP_ALIGN.LEFT

        p2 = tf.add_paragraph()
        p2.text = titles_ar[i]
        p2.font.name = FONT_AR
        p2.font.size = Pt(14)
        p2.font.color.rgb = COL["white"]
        p2.alignment = PP_ALIGN.RIGHT

        # Fake UI blocks
        for k in range(4):
            b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x+0.28), Inches(y+1.35 + k*0.78), Inches(2.05), Inches(0.55))
            b.fill.solid()
            b.fill.fore_color.rgb = COL["panel"]
            b.line.color.rgb = COL["line"]

        # CTA button
        btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x+0.6), Inches(y+4.3), Inches(1.4), Inches(0.45))
        btn.fill.solid()
        btn.fill.fore_color.rgb = COL["gold"]
        btn.line.fill.background()
        tf = btn.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = "Accept" if is_driver else "Checkout"
        p.font.name = FONT_EN
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COL["bg"]
        p.alignment = PP_ALIGN.CENTER

def _kpi_dashboard(slide):
    kpis = [
        ("Orders / day", "طلبات / يوم"),
        ("Completion rate", "نسبة الإكمال"),
        ("Avg delivery time", "متوسط وقت التوصيل"),
        ("Drivers online", "السائقون المتصلون"),
        ("Credits sold", "الكريديت المباعة"),
        ("Credits used", "الكريديت المستخدمة"),
    ]
    x0, y0 = 0.8, 1.8
    w, h = 3.95, 1.1
    for i, (en, ar) in enumerate(kpis):
        x = x0 + (i % 3) * (w + 0.3)
        y = y0 + (i // 3) * (h + 0.35)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        card.fill.solid()
        card.fill.fore_color.rgb = COL["panel"]
        card.line.color.rgb = COL["line"]

        tf = card.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = en
        p.font.name = FONT_EN
        p.font.size = Pt(14)
        p.font.color.rgb = COL["gold"]
        p.font.bold = True

        p2 = tf.add_paragraph()
        p2.text = ar
        p2.font.name = FONT_AR
        p2.font.size = Pt(14)
        p2.font.color.rgb = COL["white"]
        p2.alignment = PP_ALIGN.RIGHT

        # Placeholder metric line
        p3 = tf.add_paragraph()
        p3.text = "—"
        p3.font.name = FONT_EN
        p3.font.size = Pt(26)
        p3.font.bold = True
        p3.font.color.rgb = COL["white"]
        p3.alignment = PP_ALIGN.LEFT

    # Simple mini-chart bars
    base_x, base_y = 0.9, 5.2
    chart_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(base_x), Inches(base_y), Inches(12.0), Inches(2.0))
    chart_bg.fill.solid()
    chart_bg.fill.fore_color.rgb = COL["panel"]
    chart_bg.line.color.rgb = COL["line"]

    for i, height in enumerate([0.6, 1.2, 0.9, 1.5, 1.0, 1.7, 1.1]):
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(base_x+0.6 + i*1.5), Inches(base_y+1.7-height), Inches(0.6), Inches(height))
        bar.fill.solid()
        bar.fill.fore_color.rgb = COL["gold"]
        bar.line.fill.background()

def build_deck(app_name, city, next_cities, base_fee, per_km, min_fee, credit_price):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_total = 18
    s_idx = 1

    # 1 Title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    title = s.shapes.add_textbox(Inches(0.9), Inches(2.6), Inches(11.6), Inches(1.2))
    tf = title.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"{app_name}"
    p.font.name = FONT_EN
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = COL["gold"]
    p.alignment = PP_ALIGN.LEFT

    p2 = tf.add_paragraph()
    p2.text = f"How it works ({city}) | كيف يعمل ({'طرابلس' if city.lower()=='tripoli' else city})"
    p2.font.name = FONT_EN
    p2.font.size = Pt(22)
    p2.font.color.rgb = COL["white"]

    _footer(s, s_idx, slides_total); s_idx += 1

    # 2 Summary
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s); _header(s, "3-step summary", "ملخص بثلاث خطوات", app_name)
    _three_step_summary(s)
    _footer(s, s_idx, slides_total); s_idx += 1

    # Helper to create bilingual bullet slides
    def add_bullets(title_en, title_ar, ben, bar):
        nonlocal s_idx
        ss = prs.slides.add_slide(prs.slide_layouts[6])
        _bg(ss); _header(ss, title_en, title_ar, app_name)
        _two_col_bullets(ss, ben, bar)
        _footer(ss, s_idx, slides_total); s_idx += 1

    # 3 What it is
    add_bullets(
        "What it is",
        "ما هو",
        [
            f"Delivery app in {city}: food, groceries, pharmacy, scheduled deliveries",
            "Cash on delivery",
            "Hybrid dispatch: auto + dispatcher override",
            "Driver acceptance requires 1 credit per order",
        ],
        [
            f"تطبيق توصيل في {('طرابلس' if city.lower()=='tripoli' else city)}: مطاعم، بقالة، صيدلية، وتوصيل مجدول",
            "الدفع كاش عند الاستلام",
            "إسناد هجين: تلقائي + تدخل الموزّع",
            "قبول السائق للطلب يحتاج كريديت واحد لكل طلب",
        ],
    )

    # 4 End-to-end flow diagram
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s); _header(s, "Order flow (end-to-end)", "مسار الطلب (من البداية للنهاية)", app_name)
    steps_en = ["Placed", "Merchant accepts", "Driver accepts (1 credit)", "Pickup + pay merchant", "On the way", "Delivered + cash"]
    steps_ar = ["تأكيد الطلب", "قبول المتجر", "قبول السائق (كريديت 1)", "استلام + دفع للمتجر", "في الطريق", "تسليم + كاش"]
    _flow_diagram(s, steps_en, steps_ar)
    _footer(s, s_idx, slides_total); s_idx += 1

    # 5 Customer journey (phone mockups)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s); _header(s, "Customer journey (screens)", "رحلة العميل (شاشات)", app_name)
    _phone_mockups(
        s,
        ["Browse", "Cart", "Checkout", "Tracking"],
        ["تصفح", "السلة", "الدفع", "التتبع"],
        is_driver=False
    )
    _footer(s, s_idx, slides_total); s_idx += 1

    # 6 Driver journey (phone mockups)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s); _header(s, "Driver journey (screens)", "رحلة السائق (شاشات)", app_name)
    _phone_mockups(
        s,
        ["Offers", "Accept (credit)", "Pickup", "Deliver + cash"],
        ["عروض", "قبول (كريديت)", "استلام", "تسليم + كاش"],
        is_driver=True
    )
    _footer(s, s_idx, slides_total); s_idx += 1

    # 7 Driver core idea (fixed)
    add_bullets(
        "Driver flow (core idea)",
        "خطوات السائق (الفكرة الأساسية)",
        [
            "Driver receives an offer in the app",
            "To accept: driver must have at least 1 credit",
            "On accept: 1 credit is deducted immediately",
            "Driver picks up, pays merchant for items, then delivers",
            "At delivery: driver collects cash (items + delivery fee + tip)",
        ],
        [
            "السائق توصله عروض الطلبات داخل التطبيق",
            "باش يقبل: لازم عنده كريديت واحد أو أكثر",
            "عند القبول: ينخصم كريديت واحد مباشرة",
            "السائق يستلم الطلب ويدفع للمتجر قيمة الطلب",
            "عند التسليم: يجمع كاش (قيمة الطلب + التوصيل + التيب)",
        ],
    )

    # 8 Credits (top-up via store/agent)
    add_bullets(
        "Credits (how drivers top up)",
        "الكريديت (كيف يشحن السائق)",
        [
            "Drivers top up credits through partner stores/agents",
            "Store enters driver number (or scans) and adds credits",
            "Each accepted order costs 1 credit",
            "If order cancels: credit is returned (your rule)",
        ],
        [
            "السائق يشحن الكريديت عبر محلات/وكلاء معتمدين",
            "المحل يدخل رقم السائق (أو يمسح) ويضيف كريديت",
            "كل طلب يقبله السائق = كريديت واحد",
            "إذا ألغي الطلب: يرجع الكريديت (حسب القاعدة)",
        ],
    )

    # 9 Money flow diagram
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s); _header(s, "Money flow (cash)", "مسار الأموال (كاش)", app_name)
    _money_flow(s)
    _footer(s, s_idx, slides_total); s_idx += 1

    # 10 Pricing (base + per km + minimum)
    add_bullets(
        "Pricing (distance + minimum)",
        "التسعير (مسافة + حد أدنى)",
        [
            "Delivery fee = max( Minimum fee , Base + (Per-km × Distance) )",
            f"Base: {base_fee or '[set later]'} LYD",
            f"Per-km: {per_km or '[set later]'} LYD/km",
            f"Minimum fee: {min_fee or '[set later]'} LYD",
            "Fees are shown before ordering",
        ],
        [
            "سعر التوصيل = أكبر( الحد الأدنى , الأساسي + (سعر/كم × المسافة) )",
            f"الأساسي: {base_fee or '[يتحدد لاحقاً]'} دينار",
            f"سعر/كم: {per_km or '[يتحدد لاحقاً]'} دينار/كم",
            f"الحد الأدنى: {min_fee or '[يتحدد لاحقاً]'} دينار",
            "السعر يظهر قبل تأكيد الطلب",
        ],
    )

    # 11 Scheduled deliveries (no extra fee)
    add_bullets(
        "Scheduled deliveries",
        "التوصيل المجدول",
        [
            "Customer selects a delivery time window",
            "Merchant prepares closer to the time",
            "Driver is assigned near the scheduled slot",
            "No extra scheduled fee (your choice)",
        ],
        [
            "العميل يحدد نافذة زمنية للتوصيل",
            "المتجر يجهّز قرب وقت التوصيل",
            "تعيين السائق يتم قرب وقت الموعد",
            "بدون رسوم إضافية للجدولة (حسب الاختيار)",
        ],
    )

    # 12 Dispatch hybrid
    add_bullets(
        "Dispatch (hybrid)",
        "الإسناد (هجين)",
        [
            "Auto: offer to nearby drivers first",
            "Dispatcher override: manual assignment when needed",
            "No-driver-found: notify customer automatically (your rule)",
        ],
        [
            "تلقائي: عرض الطلب على السائقين القريبين أولاً",
            "تدخل الموزّع: تعيين يدوي عند الحاجة",
            "عدم توفر سائق: إشعار العميل تلقائياً (حسب القاعدة)",
        ],
    )

    # 13 Cancellations (credit returns)
    add_bullets(
        "Cancellations & credit",
        "الإلغاء والكريديت",
        [
            "If order is canceled: the driver credit is returned",
            "Reason tracking: customer / merchant / driver / system",
            "Prevent abuse: repeated cancels flagged in admin stats",
        ],
        [
            "إذا ألغي الطلب: يرجع كريديت السائق",
            "تسجيل سبب الإلغاء: عميل / متجر / سائق / النظام",
            "منع الاستغلال: الإلغاءات المتكررة تُرصد في لوحة الإدارة",
        ],
    )

    # 14 Support (in-app chat only)
    add_bullets(
        "Support (in-app)",
        "الدعم (داخل التطبيق)",
        [
            "In-app chat for customers, drivers, and merchants",
            "Common cases: wrong/missing items, delays, unreachable customer",
            "Resolution workflow tracked in admin panel",
        ],
        [
            "دعم عبر المحادثة داخل التطبيق للعميل والسائق والمتجر",
            "حالات شائعة: نقص/خطأ بالطلب، تأخير، عدم الرد",
            "كل الحالات تُسجّل وتُتابع عبر لوحة الإدارة",
        ],
    )

    # 15 Safety & verification (filled, editable later)
    add_bullets(
        "Safety & verification (planned)",
        "السلامة والتحقق (مخطط)",
        [
            "Driver phone verification + ID + selfie",
            "Vehicle info (plate) stored",
            "Ratings + flags for fraud prevention",
        ],
        [
            "تحقق رقم الهاتف + هوية + صورة سيلفي للسائق",
            "تسجيل بيانات المركبة (اللوحة)",
            "تقييمات وتنبيهات لمنع الاحتيال",
        ],
    )

    # 16 Merchant flow (driver pays cash at pickup)
    add_bullets(
        "Merchant flow",
        "خطوات المتجر",
        [
            "Merchant receives order in the app",
            "Accept → prepare → mark ready",
            "Driver arrives, verifies code, pays cash for items",
            "Hand-off completed",
        ],
        [
            "المتجر يستقبل الطلب داخل التطبيق",
            "قبول → تجهيز → جاهز للاستلام",
            "السائق يتحقق من الكود ويدفع كاش قيمة الطلب",
            "تسليم الطلب للسائق",
        ],
    )

    # 17 Admin dashboard (visual)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s); _header(s, "Admin stats (dashboard)", "إحصائيات الإدارة (لوحة)", app_name)
    _kpi_dashboard(s)
    _footer(s, s_idx, slides_total); s_idx += 1

    # 18 Why + Expansion
    add_bullets(
        "Why Kaziony + expansion",
        "لماذا كازيوني + التوسع",
        [
            "Cash-first model fits the market",
            "Credits reduce spam accepts and improve reliability",
            "Hybrid dispatch improves completion rate",
            f"Tripoli now → Next: {next_cities or '[set later]'}",
        ],
        [
            "نظام كاش مناسب للسوق",
            "الكريديت يقلل القبول العشوائي ويزيد الموثوقية",
            "الإسناد الهجين يرفع نسبة الإكمال",
            f"الآن: طرابلس → القادم: {next_cities or '[يتحدد لاحقاً]'}",
        ],
    )

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()

# Streamlit UI
st.set_page_config(page_title="Kaziony PPT Maker", layout="centered")
st.title("Kaziony PPT Maker | مولّد عرض كازيوني (Premium)")

c1, c2 = st.columns(2)
with c1:
    app_name = st.text_input("App name", value="Kaziony")
with c2:
    city = st.text_input("City", value="Tripoli")

next_cities = st.text_input("Next cities (comma-separated)", value="Benghazi, Misrata")

c3, c4, c5, c6 = st.columns(4)
with c3:
    base_fee = st.text_input("Base fee (LYD)", value="")
with c4:
    per_km = st.text_input("Per-km fee (LYD/km)", value="")
with c5:
    min_fee = st.text_input("Minimum fee (LYD)", value="")
with c6:
    credit_price = st.text_input("Credit price (LYD/order)", value="")

st.caption("Tip: For best Arabic look, install 'Cairo' font on your PC (optional).")

if st.button("Generate PowerPoint"):
    ppt = build_deck(app_name, city, next_cities, base_fee, per_km, min_fee, credit_price)
    st.download_button(
        "Download PPTX",
        data=ppt,
        file_name=f"{app_name}_Premium_Explainer.pptx",
        mime=PPT_MIME,
    )
