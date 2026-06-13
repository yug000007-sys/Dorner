import io
import os
import re
import zipfile
import tempfile
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from html import escape

import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch

APP_TITLE = "Dorner Lead Automation"
TEMPLATE_FILE = "template.xlsx"
DEFAULT_SUBJECT = "URGENT DORNER LEAD - Quote EQUIPMENT"

CONFIG_COMMENT = (
    "Please find the following new RFQ and URGENT lead from Dorner Config and process accordingly. "
    "Kindly contact the customer to review the application to ensure the proper equipment is quoted based upon "
    "the application requirements. Please click on 'Click Here' below to view the lead details."
)
CAD_COMMENT = (
    "Please find the following new RFQ and URGENT lead from Dorner CAD and process accordingly. "
    "Kindly contact the customer to review the application to ensure the proper equipment is quoted based upon "
    "the application requirements. Please click on 'Click Here' below to view the lead details."
)

HEADER_MAP = {
    "Brand": "Brand",
    "ReceivedDateTime": "ReceivedDateTime",
    "FirstName": "FirstName",
    "LastName": "LastName",
    "ContactTitle": "ContactTitle",
    "Email": "Email",
    "Company": "Company",
    "Address": "Address",
    "City": "City",
    "State": "State",
    "ZipCode": "ZipCode",
    "Country": "Country",
    "LeadSource1": "LeadSource1",
    "LeadComments": "LeadComments",
    "PhoneSupplied": "PhoneSupplied",
    "PhoneResearched": "PhoneResearched",
    "PDF": "PDF",
    "Product": "Product",
    "Keyword": "Keyword",
    "device": "device",
}


def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[\t\u00a0]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def html_to_text(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "lxml")
    return normalize_text(soup.get_text("\n"))


def parse_msg_bytes(msg_bytes: bytes, original_name: str):
    """Return subject, body_text, created_dt. Uses extract_msg with a temp file."""
    try:
        import extract_msg
        with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
            tmp.write(msg_bytes)
            tmp_path = tmp.name
        try:
            msg = extract_msg.Message(tmp_path)
            subject = (getattr(msg, "subject", None) or original_name or DEFAULT_SUBJECT).strip()
            body = getattr(msg, "body", None) or ""
            if not body and getattr(msg, "htmlBody", None):
                html_body = msg.htmlBody
                if isinstance(html_body, bytes):
                    html_body = html_body.decode("utf-8", errors="ignore")
                body = html_to_text(html_body)
            date_value = getattr(msg, "date", None) or getattr(msg, "delivery_time", None)
            created_dt = parse_any_datetime(str(date_value)) if date_value else None
            try:
                msg.close()
            except Exception:
                pass
            return subject, normalize_text(body), created_dt, None
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    except Exception as exc:
        return original_name or DEFAULT_SUBJECT, "", None, str(exc)


def parse_any_datetime(value: str):
    if not value:
        return None
    value = value.strip()
    patterns = [
        "%a %d %b %Y %I:%M:%S %p %Z",
        "%a, %d %b %Y %H:%M:%S %z",
        "%d %b %Y %H:%M:%S %z",
        "%Y-%m-%d %H:%M:%S%z",
        "%Y-%m-%d %H:%M:%S",
    ]
    cleaned = value.replace(" UTC", " +0000")
    for fmt in patterns:
        try:
            return datetime.strptime(cleaned, fmt)
        except Exception:
            pass
    try:
        return parsedate_to_datetime(value)
    except Exception:
        return None


def extract_created_date(text: str, fallback=None):
    m = re.search(r"Created:\s*(.+)", text or "", flags=re.I)
    if m:
        dt = parse_any_datetime(m.group(1).strip())
        if dt:
            return dt
    return fallback or datetime.now(timezone.utc)


def field_after_label(text: str, label: str):
    pattern = rf"^{re.escape(label)}\s*:\s*(.+)$"
    m = re.search(pattern, text or "", flags=re.I | re.M)
    return m.group(1).strip() if m else ""


def parse_name(full_name: str):
    parts = [p for p in re.split(r"\s+", (full_name or "").strip()) if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def parse_address(text: str):
    m = re.search(r"^Address\s*:\s*(.+?)(?:\n\s*\n|\n\s*Dorner Quote:|\Z)", text or "", flags=re.I | re.M | re.S)
    block = normalize_text(m.group(1)) if m else ""
    lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
    address = ""
    city = state = zipcode = country = ""
    if lines:
        last = lines[-1]
        cm = re.search(r"(.+?)\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s+([A-Z]{2,3}|USA|US)$", last)
        if cm:
            city, state, zipcode, country = cm.group(1).strip(), cm.group(2), cm.group(3), cm.group(4)
            address = ", ".join(lines[:-1]).strip()
        else:
            address = ", ".join(lines).strip()
    if country.upper() == "US":
        country = "USA"
    return address, city, state, zipcode, country


def extract_quote_total_lead_time(text: str):
    quote = field_after_label(text, "Dorner Quote")
    if quote:
        quote = re.split(r"\s+Grand Total", quote, flags=re.I)[0].strip()
    total = ""
    mt = re.search(r"Grand Total\s*:\s*\$?\s*([\d,]+\.\d{2})", text or "", flags=re.I)
    if mt:
        total = "$" + mt.group(1)
    lead = ""
    ml = re.search(r"Lead Time \(Business Days\)\s*(\d+)", text or "", flags=re.I)
    if ml:
        lead = ml.group(1)
    return quote, total, lead


def extract_device(text: str):
    """
    Device column rule:
    keep the full Dorner lead detail from the first Distributor line through
    all quote/product/general notes/spare-parts sections. Do NOT stop at
    Grand Total, because some leads include Additional Spare Parts after the
    quote total. Only remove the Dorner footer and Created line when present.
    """
    if not text:
        return ""

    start = re.search(r"^Distributor\s*:", text, flags=re.I | re.M)
    start_idx = start.start() if start else 0

    tail = text[start_idx:]
    end_candidates = []
    for pattern in [
        r"^\s*©\s*\d{4}\s+Dorner\s+Mfg\.\s+Corp\.",
        r"^\s*Created\s*:",
        r"^\s*USA\s*:\s*800\.397\.8664",
    ]:
        m = re.search(pattern, tail, flags=re.I | re.M)
        if m:
            end_candidates.append(m.start())

    end_idx = start_idx + min(end_candidates) if end_candidates else len(text)
    device = normalize_text(text[start_idx:end_idx]).strip()

    # Excel supports a maximum of 32,767 characters in one cell. Keep the full
    # device content whenever possible; if a very rare lead exceeds Excel's
    # hard limit, keep the maximum Excel can store rather than failing.
    return device[:32767]


def apply_brand_product_rules(device_text: str):
    lower = (device_text or "").lower()
    if "garvey" in lower:
        return "Garvey", ""
    if "montratec" in lower:
        return "Montratec", ""
    if "aquagard" in lower or "aquapruf" in lower:
        return "Dorner", "AquaX"
    return "Dorner", ""


def build_lead_record(subject: str, body_text: str, created_dt: datetime, base_filename: str):
    text = normalize_text(body_text)
    created_dt = extract_created_date(text, created_dt)
    device = extract_device(text)
    brand, product = apply_brand_product_rules(device)
    full_name = field_after_label(text, "Name")
    first_name, last_name = parse_name(full_name)
    address, city, state, zipcode, country = parse_address(text)
    quote, total, lead_time = extract_quote_total_lead_time(text)

    if re.search(r"Dorner\s+CAD", text, flags=re.I):
        lead_comment = CAD_COMMENT
    else:
        lead_comment = CONFIG_COMMENT

    pdf_col = f"{base_filename}.pdf, {base_filename}.msg, {base_filename}.doc"
    phone = field_after_label(text, "Phone")

    return {
        "Brand": brand,
        "Product": product,
        "ReceivedDateTime": created_dt.replace(tzinfo=None) if created_dt else None,
        "FirstName": first_name,
        "LastName": last_name,
        "ContactTitle": field_after_label(text, "Title"),
        "Email": field_after_label(text, "Email"),
        "Company": field_after_label(text, "Company"),
        "Address": address,
        "City": city,
        "State": state,
        "ZipCode": zipcode,
        "Country": country or "USA",
        "LeadSource1": "Request For Quote",
        "LeadComments": lead_comment,
        "PhoneSupplied": phone,
        "PhoneResearched": format_phone(phone),
        "PDF": pdf_col,
        "Keyword": subject.strip() if subject else DEFAULT_SUBJECT,
        "device": device,
        "DornerQuote": quote,
        "GrandTotal": total,
        "LeadTimeBusinessDays": lead_time,
        "Industry": field_after_label(text, "Industry"),
        "Distributor": field_after_label(text, "Distributor"),
    }


def format_phone(phone: str):
    digits = re.sub(r"\D", "", phone or "")
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return phone or ""


def filename_from_dt(dt: datetime):
    if not dt:
        dt = datetime.now(timezone.utc)
    return "Dorner_" + dt.strftime("%Y%m%d_%H%M%S")


def safe_para(text: str):
    return escape(text or "").replace("\n", "<br/>")


def make_pdf_bytes(title: str, body_text: str):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=0.5*inch, leftMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
    styles = getSampleStyleSheet()
    story = [Paragraph(escape(title), styles["Title"]), Spacer(1, 0.15*inch)]
    for chunk in normalize_text(body_text).split("\n"):
        story.append(Paragraph(escape(chunk) if chunk else " ", styles["BodyText"]))
        story.append(Spacer(1, 0.04*inch))
    doc.build(story)
    return buffer.getvalue()


def make_doc_bytes(title: str, body_text: str):
    # Word can open this HTML content even with .doc extension.
    html = f"""<html><head><meta charset='utf-8'><title>{escape(title)}</title></head>
<body><h1>{escape(title)}</h1><div>{safe_para(body_text)}</div></body></html>"""
    return html.encode("utf-8")


def append_records_to_workbook(records):
    if os.path.exists(TEMPLATE_FILE):
        wb = load_workbook(TEMPLATE_FILE)
    else:
        raise FileNotFoundError(f"{TEMPLATE_FILE} not found. Keep it beside app.py.")
    ws = wb.active
    headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    header_to_col = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h}

    device_col = header_to_col.get("device") or header_to_col.get("Device")

    for record in records:
        row = ws.max_row + 1
        for excel_header, record_key in HEADER_MAP.items():
            if excel_header in header_to_col:
                cell = ws.cell(row=row, column=header_to_col[excel_header])
                cell.value = record.get(record_key, "")
                if excel_header.lower() == "device":
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
        if device_col:
            ws.row_dimensions[row].height = 120

    if device_col:
        # Make the long device text visible/readable in Excel. This does not
        # truncate content; users can expand the row height or read formula bar.
        col_letter = ws.cell(row=1, column=device_col).column_letter
        ws.column_dimensions[col_letter].width = 80
        ws.cell(row=1, column=device_col).alignment = Alignment(wrap_text=True, vertical="top")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def process_one_file(uploaded_file, pasted_text=""):
    msg_bytes = uploaded_file.getvalue() if uploaded_file else b""
    original_name = uploaded_file.name if uploaded_file else "pasted_text.msg"
    subject, body, msg_dt, error = parse_msg_bytes(msg_bytes, original_name) if msg_bytes else (DEFAULT_SUBJECT, "", None, None)
    if pasted_text.strip():
        body = normalize_text(pasted_text)
    if not body:
        raise ValueError(f"Could not read message body. Parser error: {error or 'unknown'}")
    created_dt = extract_created_date(body, msg_dt)
    base_name = filename_from_dt(created_dt)
    record = build_lead_record(subject, body, created_dt, base_name)
    return {
        "base_name": base_name,
        "subject": subject,
        "body": body,
        "msg_bytes": msg_bytes,
        "record": record,
        "pdf_bytes": make_pdf_bytes(subject, body),
        "doc_bytes": make_doc_bytes(subject, body),
        "parse_warning": error,
    }


def build_output_zip(processed_items):
    records = [item["record"] for item in processed_items]
    excel_bytes = append_records_to_workbook(records)
    zbuf = io.BytesIO()
    used_names = set()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Dorner_Leads_Output.xlsx", excel_bytes)
        for item in processed_items:
            base = item["base_name"]
            if base in used_names:
                suffix = 2
                while f"{base}_{suffix}" in used_names:
                    suffix += 1
                base = f"{base}_{suffix}"
            used_names.add(base)
            z.writestr(f"{base}.pdf", item["pdf_bytes"])
            z.writestr(f"{base}.doc", item["doc_bytes"])
            z.writestr(f"{base}.msg", item["msg_bytes"] or item["body"].encode("utf-8"))
    zbuf.seek(0)
    return zbuf.getvalue()


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption("Upload Dorner .msg leads, apply mapping rules, and download Excel + PDF/MSG/DOC files.")

with st.expander("Rules used by this app", expanded=False):
    st.markdown("""
- `LeadSource1` is always `Request For Quote`.
- `Keyword` is the email subject line.
- If body contains `Dorner CAD`, CAD lead comment is used; otherwise Config comment is used.
- If Device contains `AquaGard` or `AquaPruf`: Brand = `Dorner`, Product = `AquaX`.
- If Device contains `Garvey`: Brand = `Garvey`, Product = blank.
- If Device contains `Montratec`: Brand = `Montratec`, Product = blank.
- File names use `Dorner_YYYYMMDD_HHMMSS` from the `Created:` line when present.
- PDF column contains `.pdf`, `.msg`, and `.doc` with the same base name.
""")

uploaded_files = st.file_uploader("Upload Dorner .msg file(s)", type=["msg"], accept_multiple_files=True)
pasted_text = st.text_area("Optional: paste message body here if MSG parsing fails or for testing", height=260)

if st.button("Process Lead(s)", type="primary"):
    if not uploaded_files and not pasted_text.strip():
        st.error("Please upload at least one .msg file or paste lead text.")
    else:
        items = []
        errors = []
        if uploaded_files:
            for uf in uploaded_files:
                try:
                    items.append(process_one_file(uf, pasted_text if len(uploaded_files) == 1 else ""))
                except Exception as exc:
                    errors.append(f"{uf.name}: {exc}")
        elif pasted_text.strip():
            try:
                items.append(process_one_file(None, pasted_text))
            except Exception as exc:
                errors.append(str(exc))

        for e in errors:
            st.error(e)

        if items:
            st.success(f"Processed {len(items)} lead(s).")
            preview_rows = []
            for item in items:
                r = item["record"]
                preview_rows.append({
                    "FileBase": item["base_name"],
                    "Brand": r.get("Brand"),
                    "Product": r.get("Product"),
                    "FirstName": r.get("FirstName"),
                    "LastName": r.get("LastName"),
                    "Company": r.get("Company"),
                    "Email": r.get("Email"),
                    "Phone": r.get("PhoneSupplied"),
                    "Quote": r.get("DornerQuote"),
                    "GrandTotal": r.get("GrandTotal"),
                    "LeadSource1": r.get("LeadSource1"),
                    "PDF Column": r.get("PDF"),
                })
                if item.get("parse_warning"):
                    st.warning(f"Parser warning for {item['base_name']}: {item['parse_warning']}")
            st.dataframe(preview_rows, use_container_width=True)
            zip_bytes = build_output_zip(items)
            st.download_button(
                "Download Output ZIP",
                data=zip_bytes,
                file_name="dorner_processed_output.zip",
                mime="application/zip",
            )
