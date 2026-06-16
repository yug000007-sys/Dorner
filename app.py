import io
import os
import re
import zipfile
from datetime import datetime
from email import policy
from email.parser import BytesParser
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

try:
    import extract_msg
except Exception:
    extract_msg = None

APP_TITLE = "Dorner Lead Automation"
EASTERN = ZoneInfo("America/New_York")

HEADERS = [
    "Referral", "ReferralEmail", "Brand", "ReceivedDateTime", "FirstName", "LastName",
    "ContactTitle", "Email", "Company", "Address", "County", "City", "State", "ZipCode",
    "Country", "LeadSource1", "LeadSource2", "LeadSource3", "LeadSource4", "LeadComments",
    "PhoneSupplied", "PhSuppliedExtension", "PhoneResearched", "CSRName", "PDF", "DUNS",
    "WebAddress", "SIC", "NAICS", "noOfEmployees", "ParentName", "LineOfBusiness", "Product",
    "Market", "PQ", "interestedIn", "crm_lead_id", "Latitude", "Longitude", "Keyword", "device",
    "DemoLead", "about_me", "college_1", "college_1_degree", "college_1_start", "college_1_end",
    "college_2", "college_2_degree", "college_2_start", "college_2_end", "month_of_joining",
    "about_experience", "Linkedin_Link", "Linkedin_Title", "searched_on_google", "linkedin_city",
    "linkedin_state", "linkedin_country", "GrandTotal"
]


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"<mailto:[^>]+>", "", text, flags=re.I)
    text = re.sub(r"<https?://[^>]+>", "", text, flags=re.I)
    text = re.sub(r"https?://\S+", "", text, flags=re.I)
    text = re.sub(r"\n{4,}", "\n\n\n", text)
    return text.strip()


def extract_msg_content(uploaded_file):
    data = uploaded_file.getvalue()
    subject = os.path.splitext(uploaded_file.name)[0]
    body = ""
    sender_date = None

    # First try Outlook MSG parser.
    if extract_msg is not None:
        try:
            msg = extract_msg.Message(io.BytesIO(data))
            subject = msg.subject or subject
            body = msg.body or msg.htmlBody or ""
            sender_date = getattr(msg, "date", None)
            if sender_date is None:
                sender_date = getattr(msg, "receivedTime", None)
        except Exception:
            pass

    # Fallback: parse as RFC email or decode raw bytes.
    if not body:
        try:
            eml = BytesParser(policy=policy.default).parsebytes(data)
            subject = eml.get("subject") or subject
            sender_date = eml.get("date") or sender_date
            if eml.is_multipart():
                parts = []
                for part in eml.walk():
                    ctype = part.get_content_type()
                    if ctype == "text/plain":
                        parts.append(part.get_content())
                body = "\n".join(parts)
            else:
                body = eml.get_content()
        except Exception:
            body = data.decode("utf-8", errors="ignore")

    body = clean_text(str(body))
    if not subject or subject == os.path.splitext(uploaded_file.name)[0]:
        m = re.search(r"Subject:\s*(.+)", body, re.I)
        if m:
            subject = m.group(1).strip()
    return subject, body, sender_date, data


def parse_datetime_from_any(raw_date, body):
    """Return Excel display time like 5/28/2026 12:15 PM.
    Priority:
    1) MSG/RFC Date header (actual received/sent time)
    2) Created line in body (UTC)
    3) current time fallback
    """
    dt = None
    if raw_date:
        if isinstance(raw_date, datetime):
            dt = raw_date
        else:
            from email.utils import parsedate_to_datetime
            try:
                dt = parsedate_to_datetime(str(raw_date))
            except Exception:
                dt = None

    if dt is None:
        m = re.search(r"Created:\s*([A-Za-z]{3}\s+\d{1,2}\s+[A-Za-z]{3}\s+\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[AP]M\s+UTC)", body, re.I)
        if m:
            try:
                dt = datetime.strptime(m.group(1), "%a %d %b %Y %I:%M:%S %p UTC").replace(tzinfo=ZoneInfo("UTC"))
            except Exception:
                dt = None

    if dt is None:
        dt = datetime.now(EASTERN)
    if dt.tzinfo is None:
        # If MSG parser returns naive, it generally reflects local received time. Keep as Eastern display.
        dt_eastern = dt.replace(tzinfo=EASTERN)
    else:
        dt_eastern = dt.astimezone(EASTERN)
    return f"{dt_eastern.month}/{dt_eastern.day}/{dt_eastern.year} {dt_eastern.strftime('%I:%M %p').lstrip('0')}"


def get_field(body, label):
    # Handles label: value until next common label or line end.
    pat = rf"{re.escape(label)}\s*:\s*(.+)"
    m = re.search(pat, body, re.I)
    if not m:
        return ""
    return m.group(1).strip()


def parse_address(body):
    m = re.search(r"Address:\s*(.+?)(?:\n\s*\n|\n\s*Dorner Quote:|\n\s*CAD Models|$)", body, re.I | re.S)
    block = m.group(1).strip() if m else ""
    lines = [x.strip() for x in block.splitlines() if x.strip()]
    address = lines[0] if lines else ""
    city = state = zip_code = country = ""
    if len(lines) >= 2:
        last = lines[-1]
        m2 = re.search(r"(.+?)\s+([A-Z]{2})\s+([A-Z0-9][A-Z0-9\s-]{2,10})\s+([A-Z]{2})$", last)
        if m2:
            city, state, zip_code, country = [v.strip() for v in m2.groups()]
    return address, city, state, zip_code, country


def split_name(full_name):
    parts = [p for p in full_name.split() if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def extract_grand_total(body):
    # Prefer the first quote header Grand Total, fallback to final total.
    patterns = [
        r"Dorner Quote:\s*[^\n\r]*?Grand\s+Total:\s*\$?\s*([\d,]+\.\d{2})",
        r"Grand\s+Total:\s*\$?\s*([\d,]+\.\d{2})",
    ]
    for pat in patterns:
        m = re.search(pat, body, re.I | re.S)
        if m:
            return "$" + m.group(1).strip()
    return ""


def extract_device(body):
    start = re.search(r"\bDistributor:\s*", body, re.I)
    if not start:
        return body.strip()
    end = len(body)
    # Keep through Created line, but remove tracking URLs after it.
    created = re.search(r"Created:\s*.*", body, re.I)
    if created and created.end() > start.start():
        end = created.end()
    return clean_text(body[start.start():end]).strip()


def determine_brand_product(device):
    d = device.lower()
    if "aquagard" in d or "aquapruf" in d:
        return "Dorner", "AquaX"
    if "garvey" in d:
        return "Garvey", ""
    if "montratec" in d:
        return "Montratec", ""
    return "Dorner", ""


def lead_comment(body):
    if re.search(r"Dorner\s+CAD", body, re.I):
        kind = "CAD"
    else:
        kind = "Config"
    return (f"Please find the following new RFQ and URGENT lead from Dorner {kind} and process accordingly. "
            "Kindly contact the customer to review the application to ensure the proper equipment is quoted based upon "
            "the application requirements. Please click on 'Click Here' below to view the lead details.")


def build_lead(uploaded_file):
    subject, body, raw_date, original_bytes = extract_msg_content(uploaded_file)
    device = extract_device(body)
    first, last = split_name(get_field(body, "Name"))
    address, city, state, zip_code, country = parse_address(body)
    brand, product = determine_brand_product(device)
    received = parse_datetime_from_any(raw_date, body)
    # File base name uses created/received timestamp if possible; safe fallback from current.
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    cm = re.search(r"Created:\s*([A-Za-z]{3}\s+\d{1,2}\s+[A-Za-z]{3}\s+\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[AP]M\s+UTC)", body, re.I)
    if cm:
        try:
            dtu = datetime.strptime(cm.group(1), "%a %d %b %Y %I:%M:%S %p UTC")
            ts = dtu.strftime("%Y%m%d_%H%M%S")
        except Exception:
            pass
    base_name = f"Dorner_{ts}"
    pdf_names = f"{base_name}.pdf, {base_name}.msg, {base_name}.doc"
    row = {h: "" for h in HEADERS}
    row.update({
        "Brand": brand,
        "Product": product,
        "ReceivedDateTime": received,
        "FirstName": first,
        "LastName": last,
        "ContactTitle": get_field(body, "Title"),
        "Email": get_field(body, "Email"),
        "Company": get_field(body, "Company"),
        "Address": address,
        "City": city,
        "State": state,
        "ZipCode": zip_code,
        "Country": "USA" if country.upper() in {"US", "USA"} else country,
        "LeadSource1": "Request For Quote",
        "LeadComments": lead_comment(body),
        "PhoneSupplied": get_field(body, "Phone"),
        "PDF": pdf_names,
        "Keyword": subject,
        "device": device,
        "GrandTotal": extract_grand_total(body),
    })
    return row, base_name, subject, body, device, original_bytes


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(str(text or ""))
    r.bold = bold
    r.font.size = Pt(9)
    if color:
        r.font.color.rgb = RGBColor(*color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def add_section_header(doc, title):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    cell = table.cell(0, 0)
    shade_cell(cell, "F4B183")
    set_cell_text(cell, title, bold=True)
    return table


def generate_docx(row, body, path):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.45)
    sec.bottom_margin = Inches(0.45)
    sec.left_margin = Inches(0.55)
    sec.right_margin = Inches(0.55)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(9)

    # Intro as target: no filename, no subject, no URLs.
    intro_end = body.find("Distributor:")
    intro = clean_text(body[:intro_end]) if intro_end > 0 else "Dorner Distributor,"
    for line in intro.splitlines():
        p = doc.add_paragraph(line)
        p.paragraph_format.space_after = Pt(1)

    add_section_header(doc, "Distributor")
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_text(t.cell(0, 0), "Distributor", True)
    shade_cell(t.cell(0, 0), "D9EAF7")
    distributor = get_field(body, "Distributor")
    set_cell_text(t.cell(0, 1), distributor)
    shade_cell(t.cell(0, 1), "D9EAF7")

    add_section_header(doc, "Customer Contact Info")
    info = [
        ("Name", f"{row['FirstName']} {row['LastName']}".strip()),
        ("Title", row["ContactTitle"]),
        ("Industry", get_field(body, "Industry")),
        ("Company", row["Company"]),
        ("Phone", row["PhoneSupplied"]),
        ("Email", row["Email"]),
        ("Address", row["Address"]),
        ("City/State/Zip/Country", " ".join([row["City"], row["State"], row["ZipCode"], row["Country"]]).strip()),
    ]
    t = doc.add_table(rows=len(info), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k, v) in enumerate(info):
        set_cell_text(t.cell(i, 0), k, True)
        set_cell_text(t.cell(i, 1), v)
        shade_cell(t.cell(i, 0), "D9EAF7")
        shade_cell(t.cell(i, 1), "D9EAF7")

    add_section_header(doc, "Quote Details")
    q = doc.add_table(rows=1, cols=3)
    q.alignment = WD_TABLE_ALIGNMENT.CENTER
    quote = get_field(body, "Dorner Quote") or re.search(r"Dorner Quote:\s*([\d]+)", body, re.I)
    quote_val = quote if isinstance(quote, str) else (quote.group(1) if quote else "")
    for i, (k, v) in enumerate([("Dorner Quote", quote_val), ("Grand Total", row["GrandTotal"]), ("Lead Time", get_field(body, "Lead Time (Business Days)") or "")]):
        set_cell_text(q.cell(0, i), f"{k}: {v}", True)
        shade_cell(q.cell(0, i), "D9EAF7")

    # Body content after customer info/quote as readable report. Remove raw URLs.
    doc.add_paragraph()
    add_section_header(doc, "Lead Details")
    detail = row["device"]
    # Avoid duplicating distributor/contact info too much? keep full text per your device rule.
    for block in detail.split("\n"):
        p = doc.add_paragraph(block)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0

    doc.save(path)


def rtf_escape(text):
    """Escape text for RTF while preserving readable Unicode where possible."""
    if text is None:
        return ""
    out = []
    for ch in str(text):
        code = ord(ch)
        if ch == "\\":
            out.append(r"\\")
        elif ch == "{":
            out.append(r"\{")
        elif ch == "}":
            out.append(r"\}")
        elif ch == "\n":
            out.append(r"\par
")
        elif code > 127:
            # RTF unicode escape. Word will display the correct character.
            if code > 32767:
                code -= 65536
            out.append(rf"\u{code}?")
        else:
            out.append(ch)
    return "".join(out)


def rtf_par(text="", bold=False, font_size=18, color=1, before=0, after=40):
    b1 = r"\b " if bold else ""
    b2 = r"\b0 " if bold else ""
    return rf"\pard\sa{after}\sb{before}\cf{color}\fs{font_size} {b1}{rtf_escape(text)}{b2}\par
"


def generate_rtf_doc(row, body, path):
    """Create a Word-openable .doc file using RTF content.

    This is the Streamlit Cloud compatible replacement for true binary .doc.
    Microsoft Word opens it as a normal document even though the file extension is .doc.
    """
    intro_end = body.find("Distributor:")
    intro = clean_text(body[:intro_end]) if intro_end > 0 else "Dorner Distributor,"

    parts = []
    parts.append(r"{\rtf1\ansi\deff0")
    parts.append(r"{\fonttbl{\f0 Calibri;}}")
    # color table: 1 black, 2 orange, 3 light blue, 4 white
    parts.append(r"{\colortbl ;\red0\green0\blue0;\red244\green177\blue131;\red217\green234\blue247;\red255\green255\blue255;}")
    parts.append(r"\paperw12240\paperh15840\margl792\margr792\margt648\margb648\f0\fs18 ")

    # Email intro exactly at top, no filename/subject.
    for line in intro.splitlines():
        parts.append(rtf_par(line, font_size=18, after=20))

    # Simple colored sections that Word can render reliably from RTF.
    def section(title):
        parts.append(rtf_par(title, bold=True, font_size=20, color=1, before=120, after=40))

    def kv(label, value):
        parts.append(rtf_par(f"{label}: {value or ''}", bold=False, font_size=18, after=20))

    section("Distributor")
    kv("Distributor", get_field(body, "Distributor"))

    section("Customer Contact Info")
    kv("Name", f"{row['FirstName']} {row['LastName']}".strip())
    kv("Title", row["ContactTitle"])
    kv("Industry", get_field(body, "Industry"))
    kv("Company", row["Company"])
    kv("Phone", row["PhoneSupplied"])
    kv("Email", row["Email"])
    kv("Address", row["Address"])
    kv("City/State/Zip/Country", " ".join([row["City"], row["State"], row["ZipCode"], row["Country"]]).strip())

    section("Quote Details")
    quote = get_field(body, "Dorner Quote") or ""
    if not quote:
        qm = re.search(r"Dorner Quote:\s*([\d]+)", body, re.I)
        quote = qm.group(1) if qm else ""
    kv("Dorner Quote", quote)
    kv("Grand Total", row.get("GrandTotal", ""))
    kv("Lead Time", get_field(body, "Lead Time (Business Days)") or "")

    section("Lead Details")
    # Preserve full device text, including spare parts and created line, but cleaned of URLs.
    for line in row["device"].splitlines():
        parts.append(rtf_par(line, font_size=18, after=20))

    parts.append("}")
    Path(path).write_text("".join(parts), encoding="utf-8")


def generate_pdf(row, body, path):
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("normal9", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, leading=10.5, alignment=TA_LEFT)
    header = ParagraphStyle("header", parent=normal, fontSize=9, leading=11, textColor=colors.black)
    story = []

    def P(text, style=normal):
        story.append(Paragraph((text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), style))

    intro_end = body.find("Distributor:")
    intro = clean_text(body[:intro_end]) if intro_end > 0 else "Dorner Distributor,"
    for line in intro.splitlines():
        P(line, normal)
    story.append(Spacer(1, 8))

    def section(title):
        tbl = Table([[Paragraph(f"<b>{title}</b>", header)]], colWidths=[7.0*inch])
        tbl.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4B183")), ("BOX", (0, 0), (-1, -1), 0.25, colors.grey)]))
        story.append(tbl)
        story.append(Spacer(1, 3))

    section("Customer Contact Info")
    rows = [
        ["Name", f"{row['FirstName']} {row['LastName']}".strip()],
        ["Title", row["ContactTitle"]],
        ["Industry", get_field(body, "Industry")],
        ["Company", row["Company"]],
        ["Phone", row["PhoneSupplied"]],
        ["Email", row["Email"]],
        ["Address", row["Address"]],
        ["City/State/Zip/Country", " ".join([row["City"], row["State"], row["ZipCode"], row["Country"]]).strip()],
    ]
    tbl = Table([[Paragraph(f"<b>{a}</b>", normal), Paragraph(str(b), normal)] for a, b in rows], colWidths=[2.0*inch, 5.0*inch])
    tbl.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#D9EAF7")), ("GRID", (0, 0), (-1, -1), 0.25, colors.white), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(tbl)
    story.append(Spacer(1, 8))

    section("Lead Details")
    for line in row["device"].splitlines():
        P(line, normal)

    doc = SimpleDocTemplate(path, pagesize=letter, rightMargin=0.55*inch, leftMargin=0.55*inch, topMargin=0.45*inch, bottomMargin=0.45*inch)
    doc.build(story)


def to_excel(rows):
    df = pd.DataFrame(rows)
    for h in HEADERS:
        if h not in df.columns:
            df[h] = ""
    df = df[HEADERS]
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Leads")
        ws = writer.book["Leads"]
        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col[:50])
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 10), 45)
        ws.freeze_panes = "A2"
    out.seek(0)
    return out.getvalue()


def process_files(files):
    rows = []
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            row, base_name, subject, body, device, original_bytes = build_lead(f)
            rows.append(row)
            zf.writestr(f"{base_name}.msg", original_bytes)
            tmp_doc = f"/tmp/{base_name}.doc"
            generate_rtf_doc(row, body, tmp_doc)
            with open(tmp_doc, "rb") as fh:
                zf.writestr(f"{base_name}.doc", fh.read())
            tmp_pdf = f"/tmp/{base_name}.pdf"
            generate_pdf(row, body, tmp_pdf)
            with open(tmp_pdf, "rb") as fh:
                zf.writestr(f"{base_name}.pdf", fh.read())
        excel_bytes = to_excel(rows)
        zf.writestr("Dorner_Leads_Output.xlsx", excel_bytes)
    zip_buffer.seek(0)
    return rows, zip_buffer.getvalue()


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.write("Upload one or more Dorner .msg files. The app creates Excel rows plus styled PDF/MSG/DOC attachments. The .doc file is RTF-based so it works on Streamlit Cloud and opens in Microsoft Word.")
files = st.file_uploader("Upload MSG files", type=["msg", "eml", "txt"], accept_multiple_files=True)

if files:
    if st.button("Process Leads", type="primary"):
        try:
            rows, output_zip = process_files(files)
            st.success(f"Processed {len(rows)} lead(s).")
            st.dataframe(pd.DataFrame(rows)[HEADERS], use_container_width=True)
            st.download_button(
                "Download ZIP",
                data=output_zip,
                file_name="dorner_processed_output.zip",
                mime="application/zip",
            )
        except Exception as e:
            st.error(f"Processing failed: {e}")
            st.exception(e)
else:
    st.info("Upload .msg files to begin.")
