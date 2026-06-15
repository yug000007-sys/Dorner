import io
import re
import zipfile
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

try:
    import extract_msg
except Exception:
    extract_msg = None

APP_HEADERS = [
    "Referral", "ReferralEmail", "Brand", "ReceivedDateTime", "FirstName", "LastName", "ContactTitle",
    "Email", "Company", "Address", "County", "City", "State", "ZipCode", "Country", "LeadSource1",
    "LeadSource2", "LeadSource3", "LeadSource4", "LeadComments", "GrandTotal", "PhoneSupplied",
    "PhSuppliedExtension", "PhoneResearched", "CSRName", "PDF", "DUNS", "WebAddress", "SIC", "NAICS",
    "noOfEmployees", "ParentName", "LineOfBusiness", "Product", "Market", "PQ", "interestedIn",
    "crm_lead_id", "Latitude", "Longitude", "Keyword", "device", "DemoLead", "about_me", "college_1",
    "college_1_degree", "college_1_start", "college_1_end", "college_2", "college_2_degree", "college_2_start",
    "college_2_end", "month_of_joining", "about_experience", "Linkedin_Link", "Linkedin_Title", "searched_on_google",
    "linkedin_city", "linkedin_state", "linkedin_country"
]

ORANGE = "FF6600"
BLUE = "C7DDF6"
DARK_BLUE = "004B83"


def clean_html(value: str) -> str:
    if not value:
        return ""
    soup = BeautifulSoup(value, "html.parser")
    for a in soup.find_all("a"):
        a.replace_with(a.get_text(" "))
    text = soup.get_text("\n")
    return normalize_text(text)


def normalize_text(text: str) -> str:
    if text is None:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{4,}", "\n\n\n", text)
    text = remove_urls(text)
    return text.strip()


def remove_urls(text: str) -> str:
    text = re.sub(r"<mailto:[^>]+>", "", text, flags=re.I)
    text = re.sub(r"<https?://[^>]+>", "", text, flags=re.I)
    text = re.sub(r"https?://\S+", "", text, flags=re.I)
    text = re.sub(r"www\.\S+", "", text, flags=re.I)
    text = re.sub(r"\n{4,}", "\n\n\n", text)
    return text.strip()


def decode_bytes(raw: bytes) -> str:
    for enc in ("utf-8", "utf-16", "cp1252", "latin1"):
        try:
            return raw.decode(enc, errors="ignore")
        except Exception:
            pass
    return raw.decode("latin1", errors="ignore")


def read_uploaded_file(uploaded) -> Tuple[str, str, str, bytes]:
    raw = uploaded.read()
    uploaded.seek(0)
    filename = uploaded.name
    subject = Path(filename).stem
    body = ""
    headers = ""

    if filename.lower().endswith(".msg") and extract_msg is not None:
        try:
            tmp = Path("/tmp") / filename
            tmp.write_bytes(raw)
            msg = extract_msg.Message(str(tmp))
            subject = msg.subject or subject
            headers = getattr(msg, "header", "") or getattr(msg, "headers", "") or ""
            if getattr(msg, "htmlBody", None):
                html = msg.htmlBody
                if isinstance(html, bytes):
                    html = decode_bytes(html)
                body = clean_html(html)
            else:
                body = msg.body or ""
            # Some MSG files expose delivery time but not raw Date header.
            if not headers:
                for attr in ("date", "receivedTime", "sent_date"):
                    val = getattr(msg, attr, None)
                    if val:
                        headers += f"\nDate: {val}"
        except Exception:
            text = decode_bytes(raw)
            body = text
            headers = text[:5000]
    else:
        text = decode_bytes(raw)
        body = text
        headers = text[:5000]

    return subject, normalize_text(body), headers, raw


def to_eastern(dt: datetime) -> datetime:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    if ZoneInfo is not None:
        return dt.astimezone(ZoneInfo("America/New_York"))
    return dt.astimezone(timezone.utc)


def format_excel_dt(dt: Optional[datetime]) -> str:
    if not dt:
        return ""
    e = to_eastern(dt)
    return f"{e.month}/{e.day}/{e.year} {e.strftime('%I').lstrip('0')}:{e.strftime('%M')} {e.strftime('%p')}"


def filebase_from_dt(dt: Optional[datetime]) -> str:
    e = to_eastern(dt) if dt else datetime.now()
    return f"Dorner_{e.strftime('%Y%m%d_%H%M%S')}"


def parse_email_datetime(headers: str, body: str) -> Optional[datetime]:
    candidates = []
    text = (headers or "") + "\n" + (body or "")
    for m in re.finditer(r"(?im)^Date:\s*(.+)$", text):
        candidates.append(m.group(1).strip())
    for m in re.finditer(r"(?im)^Created:\s*(.+)$", text):
        candidates.append(m.group(1).strip())
    for m in re.finditer(r"(?im)^Sent:\s*(.+)$", text):
        candidates.append(m.group(1).strip())
    for item in candidates:
        try:
            dt = parsedate_to_datetime(item)
            if dt:
                return dt
        except Exception:
            pass
        for fmt in (
            "%a %d %b %Y %I:%M:%S %p %Z",
            "%a %d %b %Y %H:%M:%S %Z",
            "%m/%d/%Y %I:%M %p",
            "%m/%d/%Y %H:%M",
        ):
            try:
                dt = datetime.strptime(item.replace("UTC", "GMT"), fmt)
                return dt.replace(tzinfo=timezone.utc)
            except Exception:
                pass
    return None


def rx(text: str, pattern: str, default: str = "", flags=re.I | re.S) -> str:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else default


def line_value(text: str, label: str) -> str:
    return rx(text, rf"(?im)^\s*{re.escape(label)}\s*:\s*(.+?)\s*$")


def parse_address(text: str) -> Dict[str, str]:
    addr_block = rx(text, r"(?im)^\s*Address\s*:\s*(.*?)(?:\n\s*\n|\n\s*Dorner Quote:|\n\s*CAD Models|\Z)")
    lines = [l.strip() for l in addr_block.splitlines() if l.strip()]
    address = lines[0] if lines else ""
    city = state = zipcode = country = ""
    if len(lines) >= 2:
        last = lines[-1]
        m = re.search(r"(.+?)\s+([A-Z]{2})\s+([A-Z0-9 -]{4,10})\s+([A-Z]{2,3})$", last)
        if m:
            city, state, zipcode, country = m.group(1).strip(), m.group(2), m.group(3).strip(), m.group(4)
        else:
            city = last
    if country == "US":
        country = "USA"
    return {"Address": address, "City": city, "State": state, "ZipCode": zipcode, "Country": country}


def parse_name(name: str) -> Tuple[str, str]:
    parts = name.split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def parse_grand_total(text: str) -> str:
    matches = re.findall(r"Grand\s+Total\s*:?\s*\$?\s*([\d,]+(?:\.\d{1,2})?)", text, flags=re.I)
    if not matches:
        return ""
    amount = matches[-1]
    try:
        val = float(amount.replace(",", ""))
        return f"${val:,.2f}"
    except Exception:
        return "$" + amount


def parse_lead(subject: str, body: str, headers: str, original_name: str, raw: bytes) -> Dict[str, str]:
    full = normalize_text(body)
    lower = full.lower()
    brand = "Dorner"
    product = ""
    if "aquagard" in lower or "aquapruf" in lower:
        product = "AquaX"
    if "garvey" in lower:
        brand, product = "Garvey", ""
    if "montratec" in lower:
        brand, product = "Montratec", ""

    name = line_value(full, "Name")
    first, last = parse_name(name)
    addr = parse_address(full)
    created_dt = parse_email_datetime(headers, full)
    received = format_excel_dt(created_dt)
    filebase = filebase_from_dt(created_dt)

    comment_type = "CAD" if re.search(r"Dorner\s+CAD", full, re.I) else "Config" if re.search(r"Dorner\s+Config", full, re.I) else "Config"
    lead_comment = (
        f"Please find the following new RFQ and URGENT lead from Dorner {comment_type} and process accordingly. "
        "Kindly contact the customer to review the application to ensure the proper equipment is quoted based upon the application requirements. "
        "Please click on 'Click Here' below to view the lead details."
    )

    start = re.search(r"(?im)^\s*Distributor\s*:", full)
    device = full[start.start():].strip() if start else full
    device = remove_urls(device)

    pdf_cell = f"{filebase}.pdf, {filebase}.msg, {filebase}.docx"

    row = {h: "" for h in APP_HEADERS}
    row.update({
        "Brand": brand,
        "ReceivedDateTime": received,
        "FirstName": first,
        "LastName": last,
        "ContactTitle": line_value(full, "Title"),
        "Email": line_value(full, "Email"),
        "Company": line_value(full, "Company"),
        "Address": addr["Address"],
        "City": addr["City"],
        "State": addr["State"],
        "ZipCode": addr["ZipCode"],
        "Country": addr["Country"],
        "LeadSource1": "Request For Quote",
        "LeadComments": lead_comment,
        "GrandTotal": parse_grand_total(full),
        "PhoneSupplied": line_value(full, "Phone"),
        "PDF": pdf_cell,
        "Product": product,
        "Keyword": subject,
        "device": device,
    })
    row["_Quote"] = rx(full, r"Dorner\s+Quote\s*:\s*([A-Z0-9-]+)")
    row["_LeadTime"] = rx(full, r"Lead\s+Time\s*\(Business\s+Days\)\s*([0-9]+)")
    row["_Distributor"] = line_value(full, "Distributor")
    row["_FileBase"] = filebase
    row["_Subject"] = subject
    row["_Raw"] = raw
    return row


def shade_cell(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text: str, bold=False, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(str(text or ""))
    run.bold = bold
    run.font.size = Pt(size)


def add_bar(doc: Document, width_cols=1):
    t = doc.add_table(rows=1, cols=width_cols)
    t.autofit = True
    for c in t.rows[0].cells:
        shade_cell(c, ORANGE)
        set_cell_text(c, " ")
    return t


def add_blue_table(doc: Document, rows: List[Tuple[str, str]], col_widths=(1.2, 4.3)):
    add_bar(doc)
    table = doc.add_table(rows=len(rows), cols=2)
    table.autofit = False
    for i, (label, val) in enumerate(rows):
        for cell in table.rows[i].cells:
            shade_cell(cell, BLUE)
        set_cell_text(table.cell(i, 0), label, bold=True)
        set_cell_text(table.cell(i, 1), val)
    return table


def generate_docx(row: Dict[str, str]) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.45)
    sec.bottom_margin = Inches(0.45)
    sec.left_margin = Inches(0.45)
    sec.right_margin = Inches(0.45)

    for line in [
        "Dorner Distributor,",
        f"Please find the following new RFQ and URGENT lead from Dorner {'CAD' if 'Dorner CAD' in row.get('device','') else 'Config'} and process accordingly.",
        "The pricing information has been submitted to the customer.",
        "Please contact the customer to review the application to ensure the proper equipment is quoted based upon the application requirements.",
        "If you have any questions, please let us know.",
        "Best Regards,",
        "Dorner Mfg. Corp.",
        "CustomerService@Dorner.com",
        "Tel: USA 800.397.8664    Global 262.367.7600",
    ]:
        p = doc.add_paragraph(line)
        if line.startswith("The pricing") or line.startswith("Please contact"):
            p.runs[0].bold = True
            p.runs[0].italic = True

    doc.add_paragraph("")
    add_blue_table(doc, [("Distributor:", row.get("_Distributor", ""))])
    doc.add_paragraph("")
    add_blue_table(doc, [
        ("Customer Contact Info:", ""),
        ("Name:", f"{row.get('FirstName','')} {row.get('LastName','')}").strip(),
        ("Title:", row.get("ContactTitle", "")),
        ("Industry:", ""),
        ("Company:", row.get("Company", "")),
        ("Phone:", row.get("PhoneSupplied", "")),
        ("Email:", row.get("Email", "")),
    ])
    doc.add_paragraph("")
    add_blue_table(doc, [("Customer Contact Info:", ""), ("Address:", f"{row.get('Address','')}\n{row.get('City','')} {row.get('State','')} {row.get('ZipCode','')} {row.get('Country','')}")])
    doc.add_paragraph("")

    add_bar(doc)
    qtable = doc.add_table(rows=2, cols=3)
    for r in qtable.rows:
        for c in r.cells:
            shade_cell(c, BLUE)
    set_cell_text(qtable.cell(0, 0), f"Dorner Quote: {row.get('_Quote','')}", bold=True)
    set_cell_text(qtable.cell(0, 1), "Grand Total:", bold=True)
    set_cell_text(qtable.cell(0, 2), row.get("GrandTotal", ""), bold=True)
    set_cell_text(qtable.cell(1, 0), "")
    set_cell_text(qtable.cell(1, 1), "")
    set_cell_text(qtable.cell(1, 2), "")

    # Add full device content after the designed header blocks.
    device = row.get("device", "")
    # Skip duplicated customer block where possible.
    m = re.search(r"(?im)^\s*Dorner\s+Quote\s*:", device)
    rest = device[m.start():] if m else device
    for para in rest.split("\n"):
        txt = para.strip()
        if not txt:
            doc.add_paragraph("")
            continue
        p = doc.add_paragraph(txt)
        if txt.lower().startswith(("2200", "3200", "general notes", "additional spare parts", "grand total", "total equipment price", "lead time")):
            p.runs[0].bold = True

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def generate_pdf(row: Dict[str, str]) -> bytes:
    buf = io.BytesIO()
    pdf = SimpleDocTemplate(buf, pagesize=letter, leftMargin=0.45*inch, rightMargin=0.45*inch, topMargin=0.45*inch, bottomMargin=0.45*inch)
    styles = getSampleStyleSheet()
    normal = styles["Normal"]
    normal.fontSize = 9
    normal.leading = 12
    bold = ParagraphStyle("bold", parent=normal, fontName="Helvetica-Bold")
    story = []

    def p(text, style=normal):
        story.append(Paragraph(str(text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), style))
        story.append(Spacer(1, 4))

    def blue(rows):
        data = [[Paragraph(f"<b>{a}</b>", normal), Paragraph(str(b or ""), normal)] for a,b in rows]
        t = Table(data, colWidths=[1.4*inch, 4.6*inch])
        t.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#C7DDF6")), ("BOX", (0,0), (-1,-1), 0, colors.white), ("VALIGN", (0,0), (-1,-1), "TOP")]))
        story.append(Table([[""]], colWidths=[6*inch], rowHeights=[0.25*inch], style=TableStyle([("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#FF6600"))])))
        story.append(t)
        story.append(Spacer(1, 10))

    p("Dorner Distributor,")
    p("Please find the following new RFQ and URGENT lead from Dorner and process accordingly.")
    p("The pricing information has been submitted to the customer.", bold)
    p("Please contact the customer to review the application to ensure the proper equipment is quoted based upon the application requirements.", bold)
    p("If you have any questions, please let us know.")
    p("Best Regards,")
    p("Dorner Mfg. Corp.")
    p("CustomerService@Dorner.com")
    p("Tel: USA 800.397.8664    Global 262.367.7600")
    blue([("Distributor:", row.get("_Distributor", ""))])
    blue([("Customer Contact Info:", ""), ("Name:", f"{row.get('FirstName','')} {row.get('LastName','')}") , ("Title:", row.get("ContactTitle","")), ("Industry:", ""), ("Company:", row.get("Company","")), ("Phone:", row.get("PhoneSupplied","")), ("Email:", row.get("Email",""))])
    blue([("Customer Contact Info:", ""), ("Address:", f"{row.get('Address','')}<br/>{row.get('City','')} {row.get('State','')} {row.get('ZipCode','')} {row.get('Country','')}")])
    q = Table([[Paragraph(f"<b>Dorner Quote: {row.get('_Quote','')}</b>", normal), Paragraph("<b>Grand Total:</b>", normal), Paragraph(f"<b>{row.get('GrandTotal','')}</b>", normal)]], colWidths=[2.5*inch, 1.7*inch, 1.8*inch])
    q.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#C7DDF6")), ("VALIGN", (0,0), (-1,-1), "TOP")]))
    story.append(Table([[""]], colWidths=[6*inch], rowHeights=[0.25*inch], style=TableStyle([("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#FF6600"))])))
    story.append(q)
    story.append(Spacer(1, 10))
    rest = row.get("device", "")
    m = re.search(r"(?im)^\s*Dorner\s+Quote\s*:", rest)
    rest = rest[m.start():] if m else rest
    for line in rest.split("\n"):
        if line.strip():
            p(line.strip())
        else:
            story.append(Spacer(1, 6))
    pdf.build(story)
    return buf.getvalue()


def build_excel(rows: List[Dict[str, str]]) -> bytes:
    df = pd.DataFrame([{h: r.get(h, "") for h in APP_HEADERS} for r in rows], columns=APP_HEADERS)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dorner Leads")
        ws = writer.book["Dorner Leads"]
        ws.freeze_panes = "A2"
        widths = {
            "A": 12, "B": 18, "C": 14, "D": 22, "E": 14, "F": 14, "G": 18,
            "H": 28, "I": 28, "J": 30, "K": 14, "L": 18, "M": 10, "N": 12, "O": 12,
            "P": 20, "T": 60, "U": 16, "V": 18, "Y": 45, "AH": 18, "AO": 45, "AP": 80,
        }
        for col in range(1, len(APP_HEADERS)+1):
            letter = ws.cell(1, col).column_letter
            ws.column_dimensions[letter].width = widths.get(letter, 16)
            ws.cell(1, col).font = ws.cell(1, col).font.copy(bold=True)
        # Force all columns to text format so currency strings and dates appear exactly.
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.number_format = "@"
        # Explicitly set GrandTotal cells as text, never formulas/numbers.
        if "GrandTotal" in APP_HEADERS:
            gt_col = APP_HEADERS.index("GrandTotal") + 1
            for r_idx, r in enumerate(rows, start=2):
                c = ws.cell(r_idx, gt_col)
                c.value = str(r.get("GrandTotal", ""))
                c.number_format = "@"
    return buf.getvalue()


def main():
    st.set_page_config(page_title="Dorner Lead Automation", layout="wide")
    st.title("Dorner Lead Automation")
    st.caption("Upload one or many Dorner .msg/.txt leads. App creates styled DOCX, PDF, MSG copies and one Excel output with all rows.")
    uploaded_files = st.file_uploader("Upload Dorner lead file(s)", type=["msg", "txt", "eml"], accept_multiple_files=True)
    if not uploaded_files:
        return

    rows = []
    output_files = []
    for up in uploaded_files:
        subject, body, headers, raw = read_uploaded_file(up)
        row = parse_lead(subject, body, headers, up.name, raw)
        rows.append(row)
        fb = row.get("_FileBase") or Path(up.name).stem
        docx = generate_docx(row)
        pdf = generate_pdf(row)
        output_files.append((f"{fb}.docx", docx))
        output_files.append((f"{fb}.pdf", pdf))
        output_files.append((f"{fb}.msg", raw))

    st.success(f"Parsed {len(rows)} lead(s). Excel will contain {len(rows)} row(s).")
    preview_cols = ["Brand", "ReceivedDateTime", "FirstName", "LastName", "Company", "Email", "PhoneSupplied", "GrandTotal", "PDF"]
    st.dataframe(pd.DataFrame([{c: r.get(c, "") for c in preview_cols} for r in rows]), use_container_width=True)

    excel = build_excel(rows)
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Dorner_Leads_Output.xlsx", excel)
        for name, data in output_files:
            z.writestr(name, data)
    zip_bytes = zip_buf.getvalue()

    c1, c2 = st.columns(2)
    c1.download_button("Download Excel (all rows)", excel, file_name="Dorner_Leads_Output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("Download All ZIP", zip_bytes, file_name="Dorner_Lead_Output_All.zip", mime="application/zip")

    with st.expander("Device text preview"):
        for r in rows:
            st.subheader(r.get("Company", "Lead"))
            st.text_area("device", r.get("device", ""), height=250, key=r.get("_FileBase", "device"))

if __name__ == "__main__":
    main()
