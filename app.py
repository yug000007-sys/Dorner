import io
import re
import zipfile
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

try:
    import extract_msg
except Exception:
    extract_msg = None

APP_TITLE = "Dorner Lead Automation"
ORANGE = "FF6600"
BLUE = "C6D9F1"
DARK_BLUE = "004C83"

EXCEL_COLUMNS = [
    "Brand", "ReceivedDateTime", "FirstName", "LastName", "ContactTitle", "Email",
    "Company", "Address", "City", "State", "ZipCode", "Country", "LeadSource1",
    "LeadComments", "PhoneSupplied", "PhoneResearched", "PDF", "Keyword", "device",
]

CAD_COMMENT = (
    "Please find the following new RFQ and URGENT lead from Dorner CAD and process accordingly. "
    "Kindly contact the customer to review the application to ensure the proper equipment is quoted "
    "based upon the application requirements. Please click on 'Click Here' below to view the lead details."
)
CONFIG_COMMENT = (
    "Please find the following new RFQ and URGENT lead from Dorner Config and process accordingly. "
    "Kindly contact the customer to review the application to ensure the proper equipment is quoted "
    "based upon the application requirements. Please click on 'Click Here' below to view the lead details."
)


def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\xa0", " ").replace("\u200b", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{4,}", "\n\n\n", text)
    return text.strip()


def remove_urls(text: str) -> str:
    text = re.sub(r"<mailto:([^>]+)>", "", text, flags=re.I)
    text = re.sub(r"<https?://[^>]+>", "", text, flags=re.I)
    text = re.sub(r"https?://\S+", "", text, flags=re.I)
    text = re.sub(r"www\.\S+", "", text, flags=re.I)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return normalize_text(text)


def read_uploaded_file(uploaded_file):
    raw = uploaded_file.read()
    name = uploaded_file.name
    subject = ""
    created = ""
    body = ""

    if name.lower().endswith(".msg") and extract_msg:
        temp_path = Path("/tmp") / name
        temp_path.write_bytes(raw)
        msg = extract_msg.Message(str(temp_path))
        subject = msg.subject or ""
        created = str(msg.date or "")
        body = msg.body or ""
    else:
        body = raw.decode("utf-8", errors="ignore")
        m = re.search(r"^Subject:\s*(.+)$", body, flags=re.I | re.M)
        if m:
            subject = m.group(1).strip()
        m = re.search(r"Created:\s*(.+)$", body, flags=re.I | re.M)
        if m:
            created = m.group(1).strip()

    body = normalize_text(body)
    if not created:
        m = re.search(r"Created:\s*(.+)$", body, flags=re.I | re.M)
        created = m.group(1).strip() if m else ""
    if not subject:
        q = re.search(r"Dorner Quote:\s*([\w-]+)", body, re.I)
        subject = f"URGENT DORNER LEAD - Quote EQUIPMENT {q.group(1)}" if q else Path(name).stem
    return raw, subject, created, body


def value_after(label: str, text: str, stop_labels=None):
    """Extract a field value from Dorner body.

    Handles both:
      Distributor: Shaltz
    and:
      Distributor:\nShaltz

    It only matches labels at the beginning of a line so it does not capture
    the greeting line "Dorner Distributor,".
    """
    stop_labels = stop_labels or []
    lines = text.splitlines()
    label_re = re.compile(rf"^\s*{re.escape(label)}\s*:\s*(.*)$", re.I)
    stop_re = None
    if stop_labels:
        stop_re = re.compile(rf"^\s*(?:{'|'.join(map(re.escape, stop_labels))})\s*:", re.I)

    for i, line in enumerate(lines):
        m = label_re.match(line)
        if not m:
            continue
        value = m.group(1).strip()
        collected = []
        if value and not re.match(r"^(Customer Contact Info|Name|Title|Industry|Company|Phone|Email|Address|Dorner Quote|CAD Models)\s*:?", value, re.I):
            collected.append(value)
        j = i + 1
        while j < len(lines):
            nxt = lines[j].strip()
            if not nxt:
                j += 1
                continue
            if stop_re and stop_re.match(nxt):
                if label.lower() == "distributor" and not collected and re.match(r"^Customer Contact Info\s*:?", nxt, re.I):
                    j += 1
                    continue
                break
            # stop if another common label starts
            if re.match(r"^(Customer Contact Info|Name|Title|Industry|Company|Phone|Email|Address|Dorner Quote|CAD Models)\s*:", nxt, re.I):
                if label.lower() == "distributor" and re.match(r"^Customer Contact Info\s*:", nxt, re.I):
                    j += 1
                    continue
                break
            collected.append(nxt)
            # Distributor is one-line only; do not accidentally absorb the next section
            if label.lower() == "distributor" and collected:
                break
            j += 1
        return normalize_text("\n".join(collected))
    return ""

def parse_address(text: str):
    m = re.search(r"Address:\s*(.*?)(?=\n\s*\n|\n\s*Dorner Quote:|\n\s*CAD Models|$)", text, flags=re.I | re.S)
    address_block = normalize_text(m.group(1)) if m else ""
    lines = [x.strip() for x in address_block.splitlines() if x.strip()]
    street = lines[0] if lines else ""
    city = state = zip_code = country = ""
    if len(lines) >= 2:
        last = lines[-1]
        mm = re.search(r"(.+?)\s+([A-Z]{2})\s+([A-Z0-9 -]{3,10})\s+([A-Z]{2})$", last.strip())
        if mm:
            city, state, zip_code, country = [g.strip() for g in mm.groups()]
    return address_block, street, city, state, zip_code, country


def split_name(full_name: str):
    parts = full_name.split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


EASTERN_TZ = ZoneInfo("America/New_York")


def parse_created_dt(created: str):
    """Parse the MSG/email timestamp and keep the original instant.

    Dorner emails usually carry UTC timestamps such as:
      Created: Wed 27 May 2026 05:54:16 PM UTC
    The Excel ReceivedDateTime must show that instant in Eastern time,
    e.g. 5/27/2026 1:54 PM.
    """
    cleaned = normalize_text(str(created or ""))
    cleaned = re.sub(r"^Created:\s*", "", cleaned, flags=re.I).strip()
    cleaned = re.sub(r"\s*\([^)]*\)\s*$", "", cleaned).strip()

    if not cleaned:
        return datetime.now(timezone.utc)

    # Try RFC/email Date formats first.
    try:
        dt = parsedate_to_datetime(cleaned)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        pass

    # Common Dorner body format: Thu 28 May 2026 04:15:23 PM UTC
    candidates = [
        "%a %d %b %Y %I:%M:%S %p %Z",
        "%a %d %B %Y %I:%M:%S %p %Z",
        "%a, %d %b %Y %H:%M:%S %z",
        "%d %b %Y %H:%M:%S %z",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M",
    ]
    for fmt in candidates:
        try:
            dt = datetime.strptime(cleaned, fmt)
            if dt.tzinfo is None:
                # If the text says UTC/GMT, treat as UTC; otherwise keep as Eastern.
                if re.search(r"\b(?:UTC|GMT)\b", cleaned, re.I):
                    dt = dt.replace(tzinfo=timezone.utc)
                else:
                    dt = dt.replace(tzinfo=EASTERN_TZ)
            return dt
        except Exception:
            pass

    # Last-resort parse for loose strings with AM/PM and optional UTC.
    m = re.search(
        r"(?:[A-Za-z]{3,9},?\s+)?(\d{1,2})\s+([A-Za-z]{3,9})\s+(20\d{2})\s+"
        r"(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?\s*(UTC|GMT)?",
        cleaned,
        re.I,
    )
    if m:
        day, mon, year, hh, mm, ss, ampm, zone = m.groups()
        ss = ss or "00"
        fmt = "%d %b %Y %I:%M:%S %p" if ampm else "%d %b %Y %H:%M:%S"
        dt = datetime.strptime(f"{day} {mon[:3]} {year} {hh}:{mm}:{ss}" + (f" {ampm.upper()}" if ampm else ""), fmt)
        dt = dt.replace(tzinfo=timezone.utc if zone else EASTERN_TZ)
        return dt

    return datetime.now(timezone.utc)


def format_received_datetime(created: str):
    dt = parse_created_dt(created).astimezone(EASTERN_TZ)
    # Windows-safe/non-platform-specific m/d/yyyy h:mm AM/PM
    return f"{dt.month}/{dt.day}/{dt.year} {dt.strftime('%I').lstrip('0')}:{dt.strftime('%M')} {dt.strftime('%p')}"


def build_file_base(created: str):
    dt = parse_created_dt(created)
    return f"Dorner_{dt.strftime('%Y%m%d_%H%M%S')}"


def extract_device_text(body: str):
    clean = remove_urls(body)
    start = clean.lower().find("distributor:")
    if start < 0:
        start = 0
    end = len(clean)
    # keep Created line, remove only tracking/url noise after it
    created_match = re.search(r"Created:\s*.+", clean, flags=re.I)
    if created_match and created_match.start() > start:
        end = created_match.end()
    return normalize_text(clean[start:end])


def detect_brand_product(text: str):
    low = text.lower()
    if "garvey" in low:
        return "Garvey", ""
    if "montratec" in low:
        return "Montratec", ""
    if "aquagard" in low or "aquapruf" in low:
        return "Dorner", "AquaX"
    return "Dorner", ""


def parse_lead(subject: str, created: str, body: str):
    text = remove_urls(body)
    device = extract_device_text(body)
    brand, product = detect_brand_product(device)
    full_name = value_after("Name", text, ["Title", "Industry", "Company", "Phone", "Email", "Address"])
    first, last = split_name(full_name)
    title = value_after("Title", text, ["Industry", "Company", "Phone", "Email", "Address"])
    industry = value_after("Industry", text, ["Company", "Phone", "Email", "Address"])
    company = value_after("Company", text, ["Phone", "Email", "Address"])
    phone = value_after("Phone", text, ["Email", "Address"])
    email = value_after("Email", text, ["Address"])
    distributor = value_after("Distributor", text, ["Customer Contact Info", "Name"])
    address_block, street, city, state, zip_code, country = parse_address(text)
    quote = re.search(r"Dorner Quote:\s*([\w-]+)", text, re.I)
    total = re.search(r"Grand Total:\s*\$?([\d,]+\.\d{2})", text, re.I)
    lead_time = re.search(r"Lead Time\s*\(Business Days\)\s*(\d+)", text, re.I)
    is_cad = "dorner cad" in text.lower()
    lead_comment = CAD_COMMENT if is_cad else CONFIG_COMMENT
    base = build_file_base(created)
    files_cell = f"{base}.pdf, {base}.msg, {base}.docx"

    return {
        "Brand": brand,
        "Product": product,
        "ReceivedDateTime": format_received_datetime(created),
        "FirstName": first,
        "LastName": last,
        "ContactTitle": title,
        "Industry": industry,
        "Email": email,
        "Company": company,
        "Address": street,
        "FullAddress": address_block,
        "City": city,
        "State": state,
        "ZipCode": zip_code,
        "Country": "USA" if country in ["US", "USA"] else country,
        "LeadSource1": "Request For Quote",
        "LeadComments": lead_comment,
        "PhoneSupplied": phone,
        "PhoneResearched": format_phone(phone),
        "PDF": files_cell,
        "Keyword": subject,
        "device": device,
        "Distributor": distributor,
        "Quote": quote.group(1) if quote else "",
        "GrandTotal": f"${total.group(1)}" if total else "",
        "LeadTime": lead_time.group(1) if lead_time else "",
        "FileBase": base,
        "CleanBody": text,
    }


def format_phone(phone: str):
    digits = re.sub(r"\D", "", phone or "")
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return phone


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_width(cell, width_inches):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(int(width_inches * 1440)))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def style_run(run, bold=False, size=10, color=None, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.name = "Arial"
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_text_paragraph(doc, text, bold=False, italic=False, size=10, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    style_run(r, bold=bold, italic=italic, size=size)
    return p


def add_color_bar(doc, width_cols=1):
    table = doc.add_table(rows=1, cols=width_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for cell in table.rows[0].cells:
        set_cell_shading(cell, ORANGE)
        cell.height = Inches(0.28)
        cell.text = ""
    return table


def add_info_table(doc, rows, widths=(1.2, 4.2)):
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"
    for label, value in rows:
        cells = table.add_row().cells
        for cell in cells:
            set_cell_shading(cell, BLUE)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_width(cells[0], widths[0])
        set_cell_width(cells[1], widths[1])
        p0 = cells[0].paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(label)
        style_run(r0, bold=True, size=9)
        p1 = cells[1].paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(value or "")
        style_run(r1, size=9)
    return table


def add_quote_table(doc, lead):
    add_color_bar(doc, 1)
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"
    widths = [1.4, 2.1, 1.4, 1.3]
    vals = ["Dorner Quote:", lead.get("Quote", ""), "Grand Total:", lead.get("GrandTotal", "")]
    for i, cell in enumerate(table.rows[0].cells):
        set_cell_shading(cell, BLUE)
        set_cell_width(cell, widths[i])
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(vals[i])
        style_run(r, bold=(i in [0,2,3]), size=9)
        if i == 3:
            r.underline = True
    return table


def parse_product_rows(clean_body: str):
    lines = [x.strip() for x in clean_body.splitlines() if x.strip()]
    rows = []
    for i, line in enumerate(lines):
        m = re.match(r"^(\d+)\s+([A-Z0-9-]+)\s+(\$[\d,]+\.\d{2}\s*ea\.)", line)
        if m:
            qty, part, price = m.groups()
            desc_parts = []
            j = i + 1
            while j < len(lines):
                nxt = lines[j]
                if re.match(r"^\d+\s+[A-Z0-9-]+\s+\$", nxt) or nxt.startswith("Total Equipment Price") or nxt.startswith("Lead Time"):
                    break
                if not re.match(r"^\$[\d,]+\.\d{2}", nxt):
                    desc_parts.append(nxt)
                j += 1
            rows.append([qty, part, " ".join(desc_parts), price])
    return rows


def add_product_section_docx(doc, lead):
    body = lead["CleanBody"]
    # title between grand total and Qty, if present
    m = re.search(r"Grand Total:\s*\$?[\d,]+\.\d{2}\s*(.*?)\s*Qty\s+Part Number", body, flags=re.I | re.S)
    title_block = normalize_text(m.group(1)) if m else ""
    for line in title_block.splitlines():
        if line.strip():
            add_text_paragraph(doc, line.strip(), bold=True, size=10, space_after=3)

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    headers = ["Qty", "Part Number / Description", "Description", "Unit Price"]
    widths = [0.45, 1.9, 4.0, 1.0]
    for idx, cell in enumerate(table.rows[0].cells):
        set_cell_width(cell, widths[idx])
        p = cell.paragraphs[0]
        r = p.add_run(headers[idx])
        style_run(r, bold=True, size=8)
    product_rows = parse_product_rows(body)
    for qty, part, desc, price in product_rows:
        cells = table.add_row().cells
        vals = [qty, part, desc, price]
        for idx, cell in enumerate(cells):
            set_cell_width(cell, widths[idx])
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(vals[idx])
            style_run(r, size=8)
    if not product_rows:
        add_text_paragraph(doc, lead["device"], size=8)


def add_notes_docx(doc, lead):
    body = lead["CleanBody"]
    m = re.search(r"Total Equipment Price:\s*(\$[\d,]+\.\d{2})", body, re.I)
    if m:
        add_text_paragraph(doc, f"Total Equipment Price: {m.group(1)}", bold=True, size=9)
    if lead.get("LeadTime"):
        add_text_paragraph(doc, f"Lead Time (Business Days) {lead['LeadTime']}", bold=True, size=9)

    notes_start = re.search(r"General Notes", body, re.I)
    if notes_start:
        notes = body[notes_start.start():]
        # keep from General Notes through Created, no URLs
        add_text_paragraph(doc, notes, size=8)


def generate_docx(lead):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.45)
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)

    for txt in [
        "Dorner Distributor,",
        "Please find the following new RFQ and URGENT lead from Dorner CAD and process accordingly." if "Dorner CAD" in lead["CleanBody"] else "Please find the following new RFQ and URGENT lead from Dorner Config and process accordingly.",
    ]:
        add_text_paragraph(doc, txt, size=10)
    add_text_paragraph(doc, "The pricing information has been submitted to the customer.", bold=True, italic=True, size=10)
    add_text_paragraph(doc, "Please contact the customer to review the application to ensure the proper equipment is quoted based upon the application requirements.", bold=True, italic=True, size=10)
    add_text_paragraph(doc, "If you have any questions, please let us know.", size=10)
    add_text_paragraph(doc, "Best Regards,", size=10)
    add_text_paragraph(doc, "Dorner Mfg. Corp.", size=10)
    add_text_paragraph(doc, "CustomerService@Dorner.com", size=10)
    add_text_paragraph(doc, "Tel: USA 800.397.8664   Global 262.367.7600", size=10)

    doc.add_paragraph()
    add_color_bar(doc)
    add_info_table(doc, [("Distributor:", lead.get("Distributor", ""))])
    doc.add_paragraph()
    add_color_bar(doc)
    add_info_table(doc, [
        ("Customer Contact Info:", ""),
        ("Name:", f"{lead.get('FirstName','')} {lead.get('LastName','')}".strip()),
        ("Title:", lead.get("ContactTitle", "")),
        ("Industry:", lead.get("Industry", "")),
        ("Company:", lead.get("Company", "")),
        ("Phone:", lead.get("PhoneSupplied", "")),
        ("Email:", lead.get("Email", "")),
    ])
    doc.add_paragraph()
    add_color_bar(doc)
    add_info_table(doc, [
        ("Customer Contact Info:", ""),
        ("Address:", lead.get("FullAddress", "")),
    ])
    doc.add_paragraph()
    add_quote_table(doc, lead)
    add_product_section_docx(doc, lead)
    add_notes_docx(doc, lead)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


def para_style(styles, name="normal", size=9, leading=12, bold=False, italic=False):
    return ParagraphStyle(
        name, parent=styles["Normal"], fontName="Helvetica-Bold" if bold else "Helvetica",
        fontSize=size, leading=leading, textColor=colors.black, spaceAfter=4
    )


def add_rl_table(story, data, col_widths, bg=colors.HexColor("#C6D9F1")):
    table = Table(data, colWidths=col_widths, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), bg),
        ("BOX", (0,0), (-1,-1), 0.25, colors.HexColor("#D9D9D9")),
        ("INNERGRID", (0,0), (-1,-1), 0.25, colors.HexColor("#D9D9D9")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("LEFTPADDING", (0,0), (-1,-1), 5),
        ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(table)


def generate_pdf(lead):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=0.35*inch, bottomMargin=0.35*inch, leftMargin=0.45*inch, rightMargin=0.45*inch)
    styles = getSampleStyleSheet()
    normal = para_style(styles, "n", 9, 12)
    bold = para_style(styles, "b", 9, 12, bold=True)
    italic_bold = ParagraphStyle("ib", parent=normal, fontName="Helvetica-BoldOblique", fontSize=9, leading=12, spaceAfter=5)
    small = para_style(styles, "s", 7.2, 9)

    story = []
    intro = [
        ("Dorner Distributor,", normal),
        ("Please find the following new RFQ and URGENT lead from Dorner CAD and process accordingly." if "Dorner CAD" in lead["CleanBody"] else "Please find the following new RFQ and URGENT lead from Dorner Config and process accordingly.", normal),
        ("The pricing information has been submitted to the customer.", italic_bold),
        ("Please contact the customer to review the application to ensure the proper equipment is quoted based upon the application requirements.", italic_bold),
        ("If you have any questions, please let us know.", normal),
        ("Best Regards,", normal),
        ("Dorner Mfg. Corp.", normal),
        ("CustomerService@Dorner.com", normal),
        ("Tel: USA 800.397.8664   Global 262.367.7600", normal),
    ]
    for txt, sty in intro:
        story.append(Paragraph(txt, sty))
    story.append(Spacer(1, 0.12*inch))
    story.append(Table([[""]], colWidths=[3.8*inch], rowHeights=[0.28*inch], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#FF6600"))])))
    add_rl_table(story, [["Distributor:", lead.get("Distributor", "")]], [1.05*inch, 2.75*inch])
    story.append(Spacer(1, 0.25*inch))
    story.append(Table([[""]], colWidths=[3.8*inch], rowHeights=[0.28*inch], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#FF6600"))])))
    add_rl_table(story, [
        ["Customer Contact Info:", ""],
        ["Name:", f"{lead.get('FirstName','')} {lead.get('LastName','')}"] ,
        ["Title:", lead.get("ContactTitle", "")],
        ["Industry:", lead.get("Industry", "")],
        ["Company:", lead.get("Company", "")],
        ["Phone:", lead.get("PhoneSupplied", "")],
        ["Email:", lead.get("Email", "")],
    ], [1.05*inch, 2.75*inch])
    story.append(Spacer(1, 0.18*inch))
    story.append(Table([[""]], colWidths=[4.0*inch], rowHeights=[0.28*inch], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#FF6600"))])))
    add_rl_table(story, [["Customer Contact Info:", ""], ["Address:", lead.get("FullAddress", "")]], [1.05*inch, 2.95*inch])
    story.append(Spacer(1, 0.2*inch))
    story.append(Table([[""]], colWidths=[7.0*inch], rowHeights=[0.28*inch], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#FF6600"))])))
    add_rl_table(story, [["Dorner Quote:", lead.get("Quote", ""), "Grand Total:", lead.get("GrandTotal", "")]], [1.2*inch, 2.0*inch, 1.2*inch, 2.6*inch])

    body = lead["CleanBody"]
    m = re.search(r"Grand Total:\s*\$?[\d,]+\.\d{2}\s*(.*?)\s*Qty\s+Part Number", body, flags=re.I | re.S)
    title_block = normalize_text(m.group(1)) if m else ""
    for line in title_block.splitlines():
        story.append(Paragraph(line, bold))
    product_rows = parse_product_rows(body)
    pdata = [[Paragraph("Qty", bold), Paragraph("Part Number / Description", bold), Paragraph("Unit Price", bold)]]
    for qty, part, desc, price in product_rows:
        pdata.append([Paragraph(qty, small), Paragraph(f"<b>{part}</b><br/>{desc}", small), Paragraph(price.replace("ea.", "<br/>ea."), small)])
    if product_rows:
        t = Table(pdata, colWidths=[0.4*inch, 5.45*inch, 1.0*inch], repeatRows=1, hAlign="LEFT")
        t.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"), ("FONTSIZE",(0,0),(-1,-1),7), ("TOPPADDING",(0,0),(-1,-1),5), ("BOTTOMPADDING",(0,0),(-1,-1),5)]))
        story.append(t)

    notes_match = re.search(r"Total Equipment Price:.*", body, flags=re.I | re.S)
    if notes_match:
        notes = notes_match.group(0)
        for para in notes.split("\n"):
            para = para.strip()
            if para:
                story.append(Paragraph(para, small if len(para) > 60 else bold))

    SimpleDocTemplate(buf, pagesize=letter, topMargin=0.35*inch, bottomMargin=0.35*inch, leftMargin=0.45*inch, rightMargin=0.45*inch).build(story)
    buf.seek(0)
    return buf.getvalue()


def build_excel(leads):
    if isinstance(leads, dict):
        leads = [leads]
    rows = []
    for lead in leads:
        data = {col: lead.get(col, "") for col in EXCEL_COLUMNS}
        data["Product"] = lead.get("Product", "")
        rows.append(data)
    df = pd.DataFrame(rows)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Leads")
        ws = writer.book["Leads"]
        for col_cells in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 12), 60)
        # Device column can be very long; keep it readable but not huge.
        try:
            device_col_idx = list(df.columns).index("device") + 1
            from openpyxl.utils import get_column_letter
            ws.column_dimensions[get_column_letter(device_col_idx)].width = 80
        except Exception:
            pass
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
    out.seek(0)
    return out.getvalue()


def build_output_zip(results, xlsx_bytes):
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in results:
            lead = item["lead"]
            base = lead["FileBase"]
            z.writestr(f"{base}.docx", item["docx_bytes"])
            z.writestr(f"{base}.pdf", item["pdf_bytes"])
            z.writestr(f"{base}.msg", item["raw"])
        z.writestr("Dorner_Leads_Output.xlsx", xlsx_bytes)
    out.seek(0)
    return out.getvalue()


def make_unique_file_bases(leads):
    seen = {}
    for lead in leads:
        base = lead.get("FileBase", "Dorner_output")
        if base not in seen:
            seen[base] = 1
            continue
        seen[base] += 1
        new_base = f"{base}_{seen[base]}"
        lead["FileBase"] = new_base
        lead["PDF"] = f"{new_base}.pdf, {new_base}.msg, {new_base}.docx"

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Upload one or many Dorner .msg/.txt leads. App creates styled DOCX, PDF, MSG copies and one Excel output with all rows.")

    uploaded_files = st.file_uploader(
        "Upload Dorner lead file(s)",
        type=["msg", "txt"],
        accept_multiple_files=True,
    )
    if not uploaded_files:
        st.info("Upload one or more Dorner lead messages to start.")
        return

    results = []
    errors = []
    for uploaded in uploaded_files:
        try:
            raw, subject, created, body = read_uploaded_file(uploaded)
            lead = parse_lead(subject, created, body)
            results.append({"uploaded_name": uploaded.name, "raw": raw, "lead": lead})
        except Exception as exc:
            errors.append(f"{uploaded.name}: {exc}")

    if errors:
        st.error("Some files could not be parsed:\n" + "\n".join(errors))
    if not results:
        return

    leads = [item["lead"] for item in results]
    make_unique_file_bases(leads)

    # Generate files after unique base names are finalized.
    for item in results:
        lead = item["lead"]
        item["docx_bytes"] = generate_docx(lead)
        item["pdf_bytes"] = generate_pdf(lead)

    xlsx_bytes = build_excel(leads)
    zip_bytes = build_output_zip(results, xlsx_bytes)

    st.success(f"Parsed {len(results)} lead(s). Excel will contain {len(results)} row(s).")

    preview_cols = ["Brand", "Product", "Distributor", "FirstName", "LastName", "Company", "Email", "PhoneSupplied", "Quote", "GrandTotal", "LeadTime", "FileBase"]
    st.dataframe(pd.DataFrame([{k: lead.get(k, "") for k in preview_cols} for lead in leads]), use_container_width=True)

    c1, c2 = st.columns(2)
    c1.download_button("Download Excel (all rows)", xlsx_bytes, file_name="Dorner_Leads_Output.xlsx")
    c2.download_button("Download All ZIP", zip_bytes, file_name="Dorner_All_Outputs.zip")

    st.subheader("Individual downloads")
    for item in results:
        lead = item["lead"]
        with st.expander(f"{lead.get('Company','')} / Quote {lead.get('Quote','')} / {lead.get('FileBase','')}"):
            d1, d2, d3 = st.columns(3)
            d1.download_button("Download DOCX", item["docx_bytes"], file_name=f"{lead['FileBase']}.docx", key=f"docx_{lead['FileBase']}")
            d2.download_button("Download PDF", item["pdf_bytes"], file_name=f"{lead['FileBase']}.pdf", key=f"pdf_{lead['FileBase']}")
            d3.download_button("Download MSG", item["raw"], file_name=f"{lead['FileBase']}.msg", key=f"msg_{lead['FileBase']}")
            st.text_area("Device", lead.get("device", ""), height=220, key=f"device_{lead['FileBase']}")

if __name__ == "__main__":
    main()
