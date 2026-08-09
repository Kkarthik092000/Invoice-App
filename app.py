import streamlit as st
import pandas as pd
import os
from datetime import datetime
from num2words import num2words

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image,
)

# --------------------------
# CONFIG
# --------------------------
EXCEL_FILE = "invoices.xlsx"
PDF_FOLDER = "invoices"
ASSETS_FOLDER = "assets"
LOGO_PATH = os.path.join(ASSETS_FOLDER, "logo.png")
SIGNATURE_PATH = os.path.join(ASSETS_FOLDER, "signature.jpg")

COMPANY_NAME = "Imarticus Learning Private Limited"
COMPANY_ADDRESS_LINES = [
    "5th Floor, B-Wing Kaledonia,",
    "HDIL Building",
    "Mumbai Maharashtra 400058",
    "India",
]
COMPANY_GSTIN = "GSTIN 27AACCI9233Q1ZY"
PLACE_OF_SUPPLY = "Maharashtra (27)"

# Course List
COURSES = [
    "Architecture Launchpad",
    "Corporate Lawyer Launchpad",
    "All Access Pack",
    "Finance Launchpad",
    "Marketing Launchpad"
]

# Item codes shown under the course name on the invoice, same as the sample doc
COURSE_ITEM_CODES = {
    "Architecture Launchpad": "IL3073391A",
    "Corporate Lawyer Launchpad": "IL3073393C",
    "All Access Pack": "IL3073395P",
    "Finance Launchpad": "IL3073397F",
    "Marketing Launchpad": "IL3073399M",
}

# --------------------------
# FONTS (Arial has full Rupee-sign glyph coverage; Helvetica does not)
# --------------------------
WINDOWS_FONTS = r"C:\Windows\Fonts"
FONT_REGULAR, FONT_BOLD, FONT_ITALIC, FONT_BOLD_ITALIC = (
    "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique"
)
RUPEE = "\u20b9"

try:
    pdfmetrics.registerFont(TTFont("Arial", os.path.join(WINDOWS_FONTS, "arial.ttf")))
    pdfmetrics.registerFont(TTFont("Arial-Bold", os.path.join(WINDOWS_FONTS, "arialbd.ttf")))
    pdfmetrics.registerFont(TTFont("Arial-Italic", os.path.join(WINDOWS_FONTS, "ariali.ttf")))
    pdfmetrics.registerFont(TTFont("Arial-BoldItalic", os.path.join(WINDOWS_FONTS, "arialbi.ttf")))
    FONT_REGULAR, FONT_BOLD, FONT_ITALIC, FONT_BOLD_ITALIC = (
        "Arial", "Arial-Bold", "Arial-Italic", "Arial-BoldItalic"
    )
except Exception:
    # Fall back to core fonts (no Rupee glyph) if system fonts aren't available
    RUPEE = "Rs. "

# --------------------------
# SETUP
# --------------------------
if not os.path.exists(PDF_FOLDER):
    os.makedirs(PDF_FOLDER)

if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=[
        "Invoice No", "UID", "Date", "Name", "Course", "Amount"
    ])
    df.to_excel(EXCEL_FILE, index=False)

df = pd.read_excel(EXCEL_FILE)

# --------------------------
# FUNCTIONS
# --------------------------

def _next_numeric_id(value, default):
    """Increment a numeric-looking ID while preserving its zero-padded width."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    s = str(value).strip()
    if s.endswith(".0"):  # pandas round-trips numeric-looking strings through Excel as floats
        s = s[:-2]
    return str(int(s) + 1).zfill(len(s))

def generate_invoice_no():
    if df.empty:
        return "INV-0194705760"
    last = str(df["Invoice No"].iloc[-1])
    prefix, _, num_part = last.rpartition("-")
    return f"{prefix}-{_next_numeric_id(num_part, num_part)}"

def generate_uid():
    if df.empty:
        return "32095838547235"
    return _next_numeric_id(df["UID"].iloc[-1], "32095838547235")

def amount_words(amount):
    words = num2words(amount, to='cardinal').replace(",", "").replace(" and ", " ")
    return words[0].upper() + words[1:]

def _styles():
    return {
        "company": ParagraphStyle("company", fontName=FONT_REGULAR, fontSize=9, leading=12.5),
        "company_bold": ParagraphStyle("company_bold", fontName=FONT_BOLD, fontSize=12, leading=15),
        "title": ParagraphStyle("title", fontName=FONT_BOLD, fontSize=24, alignment=TA_RIGHT),
        "label": ParagraphStyle("label", fontName=FONT_REGULAR, fontSize=9, textColor=colors.HexColor("#333333")),
        "value": ParagraphStyle("value", fontName=FONT_BOLD, fontSize=9),
        "bar_label": ParagraphStyle("bar_label", fontName=FONT_BOLD, fontSize=9, textColor=colors.HexColor("#333333")),
        "customer": ParagraphStyle("customer", fontName=FONT_BOLD, fontSize=12),
        "item_header": ParagraphStyle("item_header", fontName=FONT_BOLD, fontSize=9, textColor=colors.HexColor("#333333")),
        "item_header_c": ParagraphStyle("item_header_c", fontName=FONT_BOLD, fontSize=9, textColor=colors.HexColor("#333333"), alignment=TA_CENTER),
        "item_header_r": ParagraphStyle("item_header_r", fontName=FONT_BOLD, fontSize=9, textColor=colors.HexColor("#333333"), alignment=TA_RIGHT),
        "item_c": ParagraphStyle("item_c", fontName=FONT_REGULAR, fontSize=9, alignment=TA_CENTER),
        "item_r": ParagraphStyle("item_r", fontName=FONT_REGULAR, fontSize=9, alignment=TA_RIGHT),
        "item_desc": ParagraphStyle("item_desc", fontName=FONT_BOLD, fontSize=10, leading=13),
        "totals_label": ParagraphStyle("totals_label", fontName=FONT_REGULAR, fontSize=9),
        "totals_value": ParagraphStyle("totals_value", fontName=FONT_REGULAR, fontSize=9, alignment=TA_RIGHT),
        "totals_note": ParagraphStyle("totals_note", fontName=FONT_REGULAR, fontSize=7.5, textColor=colors.HexColor("#555555")),
        "total_label": ParagraphStyle("total_label", fontName=FONT_BOLD, fontSize=10.5),
        "total_value": ParagraphStyle("total_value", fontName=FONT_BOLD, fontSize=10.5, alignment=TA_RIGHT),
        "words_value": ParagraphStyle("words_value", fontName=FONT_BOLD_ITALIC, fontSize=9, leading=12),
        "note_label": ParagraphStyle("note_label", fontName=FONT_REGULAR, fontSize=9, textColor=colors.HexColor("#333333")),
        "terms": ParagraphStyle("terms", fontName=FONT_REGULAR, fontSize=8.5, leading=12),
        "terms_bold": ParagraphStyle("terms_bold", fontName=FONT_BOLD, fontSize=9, textColor=colors.HexColor("#333333")),
        "signatory": ParagraphStyle("signatory", fontName=FONT_REGULAR, fontSize=9, alignment=TA_CENTER),
    }

def generate_pdf(data):
    path = f"{PDF_FOLDER}/{data['Invoice No']}.pdf"
    s = _styles()
    amount = data["Amount"]
    amount_str = f"{amount:.2f}"
    grey = colors.HexColor("#9e9e9e")
    bar_bg = colors.HexColor("#f0f0f3")

    doc = SimpleDocTemplate(
        path, pagesize=A4,
        leftMargin=35, rightMargin=35, topMargin=30, bottomMargin=30,
    )
    story = []

    # ---- Header: logo + company info | FEE RECEIPT ----
    if os.path.exists(LOGO_PATH):
        logo = Image(LOGO_PATH, width=42, height=47)
    else:
        logo = Spacer(42, 47)

    company_info = Paragraph(
        f"<b>{COMPANY_NAME}</b><br/>"
        + "<br/>".join(COMPANY_ADDRESS_LINES) + "<br/>"
        + COMPANY_GSTIN,
        s["company"],
    )
    header_inner = Table([[logo, company_info]], colWidths=[50, 300])
    header_inner.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    header_table = Table(
        [[header_inner, Paragraph("FEE RECEIPT", s["title"])]],
        colWidths=[350, 175],
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 10))

    # ---- Meta info block: invoice #, dates, UID | place of supply ----
    try:
        date_display = datetime.strptime(str(data["Date"]), "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        date_display = str(data["Date"])

    meta_rows = [
        [Paragraph("#", s["label"]), Paragraph(f": {data['Invoice No']}", s["value"]),
         "", Paragraph("Place Of Supply", s["label"]), Paragraph(f": {PLACE_OF_SUPPLY}", s["value"])],
        [Paragraph("Receipt Date", s["label"]), Paragraph(f": {date_display}", s["value"]), "", "", ""],
        [Paragraph("Due Date", s["label"]), Paragraph(": NA", s["value"]), "", "", ""],
        [Paragraph("UID", s["label"]), Paragraph(f": {data['UID']}", s["value"]), "", "", ""],
    ]
    meta_table = Table(meta_rows, colWidths=[75, 150, 15, 105, 180])
    meta_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.75, grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(meta_table)

    # ---- Body block: Bill To / items / totals / notes / terms ----
    col_widths = [35, 235, 55, 85, 115]
    body_rows = []
    style_cmds = [
        ("BOX", (0, 0), (-1, -1), 0.75, grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    r = 0

    # Bill To bar
    body_rows.append([Paragraph("Bill To", s["bar_label"]), "", "", "", ""])
    style_cmds += [
        ("SPAN", (0, r), (-1, r)),
        ("BACKGROUND", (0, r), (-1, r), bar_bg),
    ]
    r += 1

    # Customer name
    body_rows.append([Paragraph(data["Name"], s["customer"]), "", "", "", ""])
    style_cmds.append(("SPAN", (0, r), (-1, r)))
    r += 1

    # Item table header
    body_rows.append([
        Paragraph("#", s["item_header_c"]),
        Paragraph("Item &amp; Description", s["item_header"]),
        Paragraph("Qty", s["item_header_r"]),
        Paragraph("Rate", s["item_header_r"]),
        Paragraph("Amount", s["item_header_r"]),
    ])
    style_cmds.append(("BACKGROUND", (0, r), (-1, r), bar_bg))
    r += 1

    # Item row
    item_code = COURSE_ITEM_CODES.get(data["Course"], "")
    item_desc_html = f"<b>MyCaptain {data['Course']}</b>"
    if item_code:
        item_desc_html += f"<br/><font size=7.5 color='#777777'>{item_code}</font>"

    body_rows.append([
        Paragraph("1", s["item_c"]),
        Paragraph(item_desc_html, s["item_desc"]),
        Paragraph("1.00", s["item_r"]),
        Paragraph(amount_str, s["item_r"]),
        Paragraph(amount_str, s["item_r"]),
    ])
    style_cmds.append(("LINEBELOW", (0, r), (-1, r), 0.5, grey))
    r += 1

    # Totals row: words on the left, sub total / total / payment / balance on the right.
    # Width must fit inside the (col3+col4) span minus its left/right padding (6+6),
    # otherwise the nested table overflows past the outer box border.
    right_block_width = col_widths[3] + col_widths[4] - 12
    totals_mini = Table([
        [Paragraph("Sub Total", s["totals_label"]), Paragraph(amount_str, s["totals_value"])],
        [Paragraph("(Tax Inclusive)", s["totals_note"]), ""],
        [Paragraph("Total", s["total_label"]), Paragraph(f"{RUPEE}{amount_str}", s["total_value"])],
        [Paragraph("Payment Made", s["totals_label"]), Paragraph(amount_str, s["totals_value"])],
        [Paragraph("Balance Due", s["total_label"]), Paragraph("0.00", s["total_value"])],
    ], colWidths=[right_block_width - 88, 88])
    totals_mini.setStyle(TableStyle([
        ("SPAN", (0, 1), (1, 1)),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("LINEABOVE", (0, 2), (1, 2), 0.5, grey),
    ]))

    words_block = Table([
        [Paragraph("Total in Words", s["totals_label"])],
        [Paragraph(amount_words(amount), s["words_value"])],
    ], colWidths=[230])
    words_block.setStyle(TableStyle([
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
    ]))

    body_rows.append([words_block, "", "", totals_mini, ""])
    style_cmds += [
        ("SPAN", (0, r), (2, r)),
        ("SPAN", (3, r), (4, r)),
        ("TOPPADDING", (0, r), (-1, r), 10),
    ]
    r += 1

    # Note
    body_rows.append([Paragraph("Note:", s["note_label"]), "", "", "", ""])
    style_cmds += [
        ("SPAN", (0, r), (-1, r)),
        ("TOPPADDING", (0, r), (-1, r), 14),
    ]
    r += 1

    # Terms & Conditions + Signature
    terms_block = Paragraph(
        "<b>Terms &amp; Conditions</b><br/>"
        "Application fees/Registration fees is non-refundable and non- transferable.",
        s["terms"],
    )
    if os.path.exists(SIGNATURE_PATH):
        sig_img = Image(SIGNATURE_PATH, width=55, height=36)
    else:
        sig_img = Spacer(1, 36)
    sig_block = Table(
        [[sig_img], [Paragraph("Authorized Signatory", s["signatory"])]],
        colWidths=[150],
    )
    sig_block.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))

    body_rows.append([terms_block, "", "", sig_block, ""])
    style_cmds += [
        ("SPAN", (0, r), (2, r)),
        ("SPAN", (3, r), (4, r)),
        ("TOPPADDING", (0, r), (-1, r), 16),
        ("BOTTOMPADDING", (0, r), (-1, r), 10),
    ]
    r += 1

    body_table = Table(body_rows, colWidths=col_widths)
    body_table.setStyle(TableStyle(style_cmds))
    story.append(body_table)

    doc.build(story)
    return path

# --------------------------
# UI
# --------------------------

st.title("📄 Invoice Generator")

name = st.text_input("Bill To (Customer Name)")
course = st.selectbox("Select Course", COURSES)
amount = st.number_input("Amount", min_value=0, step=100)
date = st.date_input("Receipt Date")

# --------------------------
# GENERATE
# --------------------------

if st.button("Generate Invoice"):
    if not name or amount == 0:
        st.error("Fill all fields")
    else:
        df = pd.read_excel(EXCEL_FILE)

        invoice_no = generate_invoice_no()
        uid = generate_uid()

        data = {
            "Invoice No": invoice_no,
            "UID": uid,
            "Date": str(date),
            "Name": name,
            "Course": course,
            "Amount": int(amount)
        }

        # Save to Excel
        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)

        # Generate PDF
        pdf = generate_pdf(data)

        st.success(f"✅ Invoice {invoice_no} created")

        st.write("### Preview")
        st.json(data)

        # Download
        with open(pdf, "rb") as f:
            st.download_button(
                "Download Invoice",
                f,
                file_name=f"{invoice_no}.pdf"
            )

# --------------------------
# SEARCH
# --------------------------

st.markdown("---")
st.subheader("🔍 Search Invoices")

search = st.text_input("Search by Name / Invoice No")

if search:
    df = pd.read_excel(EXCEL_FILE)
    result = df[
        df["Name"].astype(str).str.contains(search, case=False, na=False, regex=False) |
        df["Invoice No"].astype(str).str.contains(search, case=False, na=False, regex=False)
    ]
    st.dataframe(result)
