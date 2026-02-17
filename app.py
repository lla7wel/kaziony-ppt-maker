import io
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# === Configuration & Colors ===
PPT_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# Theme Colors
PRIMARY_COLOR = RGBColor(0, 128, 128)    # Teal/Cyan - for headers
SECONDARY_BG_COLOR = RGBColor(240, 245, 247) # Very light gray/blue - for content boxes
TEXT_COLOR_DARK = RGBColor(50, 50, 50)   # Dark gray - for standard text
TEXT_COLOR_WHITE = RGBColor(255, 255, 255) # White - for text on primary color

# Dimensions
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)
MARGIN = Inches(0.5)
HEADER_HEIGHT = Inches(1.2)
CONTENT_TOP = HEADER_HEIGHT + Inches(0.3)
CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - MARGIN
BOX_WIDTH = (SLIDE_WIDTH - (MARGIN * 3)) / 2


# === Helper Functions for Styling ===

def style_title_shape(shape, text_en, text_ar):
    """Styles the header bar shape."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = PRIMARY_COLOR
    shape.line.fill.background() # No outline

    tf = shape.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    p = tf.paragraphs[0]
    # English part
    run_en = p.add_run()
    run_en.text = title_en
    run_en.font.name = "Arial"
    run_en.font.size = Pt(28)
    run_en.font.bold = True
    run_en.font.color.rgb = TEXT_COLOR_WHITE

    # Separator
    run_sep = p.add_run()
    run_sep.text = "  |  "
    run_sep.font.size = Pt(28)
    run_sep.font.color.rgb = TEXT_COLOR_WHITE
    
    # Arabic part
    run_ar = p.add_run()
    run_ar.text = title_ar
    run_ar.font.name = "Arial"
    run_ar.font.size = Pt(28)
    run_ar.font.bold = True
    run_ar.font.color.rgb = TEXT_COLOR_WHITE

def style_content_box(shape, bullets, is_arabic=False):
    """Styles the rounded rectangles and adds bullet points."""
    # Shape style
    shape.fill.solid()
    shape.fill.fore_color.rgb = SECONDARY_BG_COLOR
    shape.line.color.rgb = PRIMARY_COLOR
    shape.line.width = Pt(1.5)

    # Text styling
    tf = shape.text_frame
    tf.clear()
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.2)
    
    for i, b in enumerate(bullets):
        pp = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        pp.text = b
        pp.font.size = Pt(18)
        pp.font.name = "Arial"
        pp.font.color.rgb = TEXT_COLOR_DARK
        # Add spacing between bullets
        pp.space_before = Pt(10) 
        
        if is_arabic:
            pp.alignment = PP_ALIGN.RIGHT

# === Main Slide Building Function ===

def add_bilingual_slide_styled(prs, title_en, title_ar, bullets_en, bullets_ar):
    # Use blank layout (usually index 6 in standard themes, but let's ensure it's blank)
    blank_layout = None
    for layout in prs.slide_layouts:
        if len(layout.placeholders) == 0:
            blank_layout = layout
            break
    if blank_layout is None:
        blank_layout = prs.slide_layouts[len(prs.slide_layouts)-1] # Fallback

    slide = prs.slides.add_slide(blank_layout)

    # 1. Header Bar (Rectangle)
    header_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 
        0, 0, SLIDE_WIDTH, HEADER_HEIGHT
    )
    style_title_shape(header_shape, title_en, title_ar)

    # 2. Left Box (English Content - Rounded Rectangle)
    left_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        MARGIN, CONTENT_TOP, BOX_WIDTH, CONTENT_HEIGHT
    )
    style_content_box(left_box, bullets_en, is_arabic=False)

    # 3. Right Box (Arabic Content - Rounded Rectangle)
    right_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        MARGIN + BOX_WIDTH + MARGIN, CONTENT_TOP, BOX_WIDTH, CONTENT_HEIGHT
    )
    style_content_box(right_box, bullets_ar, is_arabic=True)


def build_pptx_bytes():
    prs = Presentation()
    # Set slide dimensions to Widescreen (16:9) for better modern look
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # --- Title Slide ---
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    
    # Style Title
    title = s0.shapes.title
    title.text = "Kaziony | كازيوني"
    title.text_frame.paragraphs[0].font.color.rgb = PRIMARY_COLOR
    title.text_frame.paragraphs[0].font.bold = True
    
    # Style Subtitle
    subtitle = s0.placeholders[1]
    subtitle.text = "How it works (Tripoli) | كيف يخدم (طرابلس)"
    subtitle.text_frame.paragraphs[0].font.color.rgb = TEXT_COLOR_DARK


    # --- Content Slides ---
    add_bilingual_slide_styled(
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

    add_bilingual_slide_styled(
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

    add_bilingual_slide_styled(
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

    add_bilingual_slide_styled(
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

    add_bilingual_slide_styled(
        prs,
        "Pricing (distance-based)",
        "التسعير (حسب المسافة)",
        ["Delivery fee = Base + (Per-km × Distance) + Scheduled fee (if any)",
         "Fees are shown before placing the order"],
        ["سعر التوصيل = أساسي + (سعر/كم × المسافة) + رسوم الجدولة (لو موجودة)",
         "السعر يظهر قبل تأكيد الطلب"]
    )

    add_bilingual_slide_styled(
        prs,
        "Dispatch (hybrid)",
        "الإسناد (هجين)",
        ["System offers order to nearby drivers first (auto)",
         "Dispatcher can assign manually if needed"],
        ["السيستم يعرض الطلب على السائقين القريبين أولاً (تلقائي)",
         "الموزّع يقدر يعيّن سائق يدوي لو احتاج"]
    )

    add_bilingual_slide_styled(
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

# === Streamlit UI ===
st.set_page_config(page_title="Kaziony PPT Maker", layout="centered")
st.title("Kaziony PowerPoint Maker | مولّد عرض كازيوني")
st.markdown("---")
st.write("Click the button below to generate a professionally styled PowerPoint presentation describing how Kaziony works.")
st.write("اضغط على الزر أدناه لإنشاء عرض تقديمي بتصميم احترافي يشرح كيفية عمل كازيوني.")

if st.button("Generate Styled PowerPoint | إنشاء العرض المطور", type="primary"):
    with st.spinner("Generating presentation... | جاري إنشاء العرض..."):
        ppt_bytes = build_pptx_bytes()
    st.success("Done! Download below. | تم! قم بالتنزيل أدناه.")
    st.download_button(
        label="Download PowerPoint (.pptx) | تنزيل العرض",
        data=ppt_bytes,
        file_name="Kaziony_Explainer_Styled.pptx",
        mime=PPT_MIME
    )
