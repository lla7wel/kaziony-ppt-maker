import io
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

PPT_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

def add_bilingual_slide(prs, title_en, title_ar, bullets_en, bullets_ar):
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(12.2), Inches(0.8))
    tf = title_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"{title_en} | {title_ar}"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.name = "Arial"

    # Left (English)
    left = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(6.0), Inches(5.6))
    tfl = left.text_frame
    tfl.clear()
    for i, b in enumerate(bullets_en):
        pp = tfl.paragraphs[0] if i == 0 else tfl.add_paragraph()
        pp.text = b
        pp.font.size = Pt(18)
        pp.font.name = "Arial"

    # Right (Arabic)
    right = slide.shapes.add_textbox(Inches(7.0), Inches(1.3), Inches(6.0), Inches(5.6))
    tfr = right.text_frame
    tfr.clear()
    for i, b in enumerate(bullets_ar):
        pp = tfr.paragraphs[0] if i == 0 else tfr.add_paragraph()
        pp.text = b
        pp.font.size = Pt(18)
        pp.font.name = "Arial"
        pp.alignment = PP_ALIGN.RIGHT

def build_pptx_bytes():
    prs = Presentation()

    # Title slide
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    s0.shapes.title.text = "Kaziony | كازيوني"
    s0.placeholders[1].text = "How it works (Tripoli) | كيف يخدم (طرابلس)"

    # Slides (edit text anytime)
    add_bilingual_slide(
        prs,
        "What it is",
        "شنو هو",
        ["Delivery app in Tripoli: food, groceries, pharmacy, scheduled deliveries",
         "Customer pays cash on delivery",
         "Driver must spend 1 credit to accept an order"],
        ["تطبيق توصيل في طرابلس: أكل، بقالة، صيدلية، وتوصيل مجدول",
         "العميل يدفع كاش عند الاستلام",
         "السائق لازم يخصم كريديت واحد باش يقبل الطلب"]
    )

    add_bilingual_slide(
        prs,
        "Customer flow",
        "خطوات العميل",
        ["Choose merchant → add items → checkout",
         "See delivery fee before ordering",
         "Track driver live on map",
         "Pay cash + tip at delivery"],
        ["يختار المتجر/المطعم → يضيف الطلب → يدفع",
         "يشوف سعر التوصيل قبل ما يأكد",
         "يتابع السائق على الخريطة",
         "يدفع كاش + تيب عند الاستلام"]
    )

    add_bilingual_slide(
        prs,
        "Driver flow (the core idea)",
        "خطوات السائق (الفكرة الأساسية)",
        ["Driver receives an order offer in the app",
         "To accept: driver must have ≥ 1 credit",
         "When driver accepts: 1 credit is deducted",
         "Driver picks up, pays for items, then collects cash + delivery fee + tip"],
        ["السائق توصله عروض طلبات في التطبيق",
         "باش يقبل: لازم عنده كريديت واحد أو أكثر",
         "لما يقبل: ينخصم كريديت واحد",
         "يلتقط الطلب ويدفع قيمته، وبعدها ياخذ الكاش + التوصيل + التيب"]
    )

    add_bilingual_slide(
        prs,
        "Credits",
        "الكريديت",
        ["Drivers buy prepaid credit cards from grocery stores",
         "Each accepted order costs 1 credit",
         "If no credit: driver can’t accept orders"],
        ["السائق يشتري كروت كريديت مسبقة الدفع من البقالات",
         "كل طلب يقبله = كريديت واحد",
         "لو ما فيش كريديت: ما يقدرش يقبل طلبات"]
    )

    add_bilingual_slide(
        prs,
        "Pricing (distance-based)",
        "التسعير (حسب المسافة)",
        ["Delivery fee = Base + (Per-km × Distance) + Scheduled fee (if any)",
         "Fees are shown before placing the order"],
        ["سعر التوصيل = أساسي + (سعر/كم × المسافة) + رسوم الجدولة (لو موجودة)",
         "السعر يظهر قبل تأكيد الطلب"]
    )

    add_bilingual_slide(
        prs,
        "Dispatch (hybrid)",
        "الإسناد (هجين)",
        ["System offers order to nearby drivers first (auto)",
         "Dispatcher can assign manually if needed"],
        ["السيستم يعرض الطلب على السائقين القريبين أولاً (تلقائي)",
         "الموزّع يقدر يعيّن سائق يدوي لو احتاج"]
    )

    add_bilingual_slide(
        prs,
        "Admin stats",
        "إحصائيات الإدارة",
        ["Orders/day, completion rate, average delivery time",
         "Driver online count, acceptance rate",
         "Cancellations + reasons, zones heatmap",
         "Credits sold vs credits used"],
        ["طلبات/يوم، نسبة الإكمال، متوسط وقت التوصيل",
         "عدد السائقين المتصلين، نسبة القبول",
         "الإلغاءات وأسبابها، خريطة المناطق",
         "الكريديت المباعة مقابل المستخدمة"]
    )

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

st.set_page_config(page_title="Kaziony PPT Maker", layout="centered")
st.title("Kaziony PowerPoint Maker | مولّد عرض كازيوني")

st.write("Click the button, then download the PPTX file.")
if st.button("Generate PowerPoint"):
    ppt_bytes = build_pptx_bytes()
    st.download_button(
        label="Download PowerPoint (.pptx)",
        data=ppt_bytes,
        file_name="Kaziony_Explainer.pptx",
        mime=PPT_MIME
    )
