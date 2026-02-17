import io
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor

PPT_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# ---------- Helpers ----------
def _rtl(paragraph):
    # Forces Right-To-Left on Arabic paragraphs
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("rtl", "1")

def _header(slide, title_en, title_ar):
    # Top dark bar
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.85)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(20, 28, 41)
    bar.line.fill.background()

    # English title (left)
    en = slide.shapes.add_textbox(Inches(0.6), Inches(0.12), Inches(6.2), Inches(0.6))
    tf = en.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title_en
    p.font.name = "Calibri"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.LEFT

    # Arabic title (right)
    ar = slide.shapes.add_textbox(Inches(6.6), Inches(0.12), Inches(6.1), Inches(0.6))
    tf = ar.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title_ar
    p.font.name = "Tahoma"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.RIGHT
    _rtl(p)

def _bilingual_bullets(slide, bullets_en, bullets_ar):
    # Small EN / AR tags
    for x, label in [(0.6, "EN"), (7.0, "AR")]:
        tag = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.1), Inches(0.9), Inches(0.35)
        )
        tag.fill.solid()
        tag.fill.fore_color.rgb = RGBColor(240, 243, 247)
        tag.line.color.rgb = RGBColor(210, 215, 224)
        tf = tag.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = label
        p.font.name = "Calibri"
        p.font.size = Pt(14)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER

    # Two columns
    left = slide.shapes.add_textbox(Inches(0.6), Inches(1.55), Inches(6.1), Inches(5.6))
    right = slide.shapes.add_textbox(Inches(7.0), Inches(1.55), Inches(6.33), Inches(5.6))

    # English box
    tfl = left.text_frame
    tfl.clear()
    tfl.word_wrap = True
    tfl.margin_left = Inches(0.15)
    tfl.margin_right = Inches(0.10)
    tfl.margin_top = Inches(0.08)
    tfl.margin_bottom = Inches(0.08)

    for i, b in enumerate(bullets_en):
        p = tfl.paragraphs[0] if i == 0 else tfl.add_paragraph()
        p.text = f"• {b}"
        p.font.name = "Calibri"
        p.font.size = Pt(20)
        p.space_after = Pt(8)

    # Arabic box
    tfr = right.text_frame
    tfr.clear()
    tfr.word_wrap = True
    tfr.margin_left = Inches(0.10)
    tfr.margin_right = Inches(0.15)
    tfr.margin_top = Inches(0.08)
    tfr.margin_bottom = Inches(0.08)

    for i, b in enumerate(bullets_ar):
        p = tfr.paragraphs[0] if i == 0 else tfr.add_paragraph()
        p.text = f"• {b}"
        p.font.name = "Tahoma"
        p.font.size = Pt(20)
        p.alignment = PP_ALIGN.RIGHT
        p.space_after = Pt(8)
        _rtl(p)

def _order_flow_diagram(slide):
    steps_en = [
        "Order placed",
        "Merchant accepts",
        "Driver accepts (1 credit)",
        "Pickup + pay merchant",
        "On the way",
        "Delivered + cash collected",
    ]
    steps_ar = [
        "تأكيد الطلب",
        "قبول المتجر",
        "قبول السائق (كريديت 1)",
        "استلام + دفع للمتجر",
        "في الطريق",
        "تسليم + تحصيل الكاش",
    ]

    x0, y, w, h, gap = 0.6, 2.1, 1.95, 1.1, 0.25

    for i in range(6):
        x = x0 + i * (w + gap)
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(240, 243, 247)
        box.line.color.rgb = RGBColor(210, 215, 224)

        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = steps_en[i]
        p1.font.name = "Calibri"
        p1.font.size = Pt(14)
        p1.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = steps_ar[i]
        p2.font.name = "Tahoma"
        p2.font.size = Pt(14)
        p2.alignment = PP_ALIGN.CENTER
        _rtl(p2)

        if i < 5:
            x1 = x + w
            x2 = x0 + (i + 1) * (w + gap)
            line = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                Inches(x1),
                Inches(y + h / 2),
                Inches(x2),
                Inches(y + h / 2),
            )
            line.line.color.rgb = RGBColor(130, 140, 155)
            line.line.width = Pt(2)

# ---------- Build deck ----------
def build_pptx_bytes(app_name, city, base_fee, per_km, credit_cost):
    prs = Presentation()
    # Force 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Title slide
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    s0.shapes.title.text = f"{app_name} | كازيوني"
    s0.placeholders[1].text = f"How it works ({city}) | كيف يخدم ({city})"

    slides = [
        (
            "What it is",
            "شنو هو",
            [
                f"Delivery app in {city}: food, groceries, pharmacy, scheduled deliveries",
                "Customer pays cash on delivery",
                "Driver accepts orders using 1 credit per order",
            ],
            [
                f"تطبيق توصيل في {city}: أكل، بقالة، صيدلية، وتوصيل مجدول",
                "العميل يدفع كاش عند الاستلام",
                "السائق يقبل الطلبات بكريديت واحد لكل طلب",
            ],
        ),
        (
            "Customer flow",
            "خطوات العميل",
            [
                "Choose merchant → add items → checkout",
                "See delivery fee before ordering",
                "Track driver live on map",
                "Pay cash + tip at delivery",
            ],
            [
                "يختار المتجر/المطعم → يضيف الطلب → تأكيد",
                "يشوف سعر التوصيل قبل ما يأكد",
                "يتابع السائق على الخريطة",
                "يدفع كاش + تيب عند الاستلام",
            ],
        ),
        (
            "Driver flow (core idea)",
            "خطوات السائق (الفكرة الأساسية)",
            [
                "Driver receives an order offer in the app",
                "To accept: driver must have at least 1 credit",
                "On accept: 1 credit is deducted immediately",
                "Driver picks up, pays for items, then collects cash + delivery fee + tip",
            ],
            [
                "السائق توصله عروض طلبات في التطبيق",
                "باش يقبل: لازم عنده كريديت واحد أو أكثر",
                "أول ما يقبل: ينخصم كريديت واحد مباشرة",
                "يلتقط الطلب ويدفع قيمته، وبعدها ياخذ الكاش + التوصيل + التيب",
            ],
        ),
        (
            "Credits",
            "الكريديت",
            [
                "Drivers buy prepaid credit cards from grocery stores",
                "Each accepted order costs 1 credit",
                f"Credit cost (example): {credit_cost if credit_cost else '[set later]'} LYD per order",
            ],
            [
                "السائق يشتري كروت كريديت مسبقة الدفع من البقالات",
                "كل طلب يقبله = كريديت واحد",
                f"سعر الكريديت (مثال): {credit_cost if credit_cost else '[يتحدد لاحقاً]'} دينار لكل طلب",
            ],
        ),
        (
            "Pricing (distance-based)",
            "التسعير (حسب المسافة)",
            [
                "Delivery fee = Base + (Per-km × Distance) + Scheduled fee (if any)",
                "Fees are shown before placing the order",
                f"Example inputs: Base={base_fee if base_fee else '[set later]'} LYD, Per-km={per_km if per_km else '[set later]'} LYD/km",
            ],
            [
                "سعر التوصيل = أساسي + (سعر/كم × المسافة) + رسوم الجدولة (لو موجودة)",
                "السعر يظهر قبل تأكيد الطلب",
                f"قيم مثال: أساسي={base_fee if base_fee else '[يتحدد لاحقاً]'} دينار، سعر/كم={per_km if per_km else '[يتحدد لاحقاً]'} دينار/كم",
            ],
        ),
        (
            "Dispatch (hybrid)",
            "الإسناد (هجين)",
            [
                "System offers order to nearby drivers first (auto)",
                "Dispatcher can assign manually if needed (override)",
                "Only drivers with credit can accept",
            ],
            [
                "السيستم يعرض الطلب على السائقين القريبين أولاً (تلقائي)",
                "الموزّع يقدر يعيّن سائق يدوي لو احتاج",
                "فقط السائق اللي عنده كريديت يقدر يقبل",
            ],
        ),
        (
            "Live tracking",
            "التتبع المباشر",
            [
                "Customer sees driver on the map",
                "Status updates: accepted → picked up → on the way → delivered",
            ],
            [
                "العميل يشوف السائق على الخريطة",
                "تحديثات الحالة: قبول → استلام → في الطريق → تسليم",
            ],
        ),
        (
            "Admin stats",
            "إحصائيات الإدارة",
            [
                "Orders/day, completion rate, average delivery time",
                "Driver online count, acceptance rate",
                "Cancellations + reasons, zones heatmap",
                "Credits sold vs credits used",
            ],
            [
                "طلبات/يوم، نسبة الإكمال، متوسط وقت التوصيل",
                "عدد السائقين المتصلين، نسبة القبول",
                "الإلغاءات وأسبابها، خريطة المناطق",
                "الكريديت المباعة مقابل المستخدمة",
            ],
        ),
    ]

    for title_en, title_ar, ben, bar in slides:
        s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
        _header(s, title_en, title_ar)
        _bilingual_bullets(s, ben, bar)

    # Diagram slide (looks way less “basic”)
    sflow = prs.slides.add_slide(prs.slide_layouts[6])
    _header(sflow, "Order flow (end-to-end)", "رحلة الطلب (من البداية للنهاية)")
    _order_flow_diagram(sflow)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ---------- Streamlit UI ----------
st.set_page_config(page_title="Kaziony PPT Maker", layout="centered")
st.title("Kaziony PowerPoint Maker | مولّد عرض كازيوني")

app_name = st.text_input("App name", value="Kaziony")
city = st.text_input("City", value="Tripoli")

col1, col2, col3 = st.columns(3)
with col1:
    base_fee = st.text_input("Base fee (LYD)", value="")
with col2:
    per_km = st.text_input("Per-km fee (LYD/km)", value="")
with col3:
    credit_cost = st.text_input("Credit cost (LYD/order)", value="")

st.write("Click generate, then download the PPTX.")
if st.button("Generate PowerPoint"):
    ppt_bytes = build_pptx_bytes(app_name, city, base_fee, per_km, credit_cost)
    st.download_button(
        "Download PowerPoint (.pptx)",
        data=ppt_bytes,
        file_name=f"{app_name}_Explainer.pptx",
        mime=PPT_MIME,
    )
