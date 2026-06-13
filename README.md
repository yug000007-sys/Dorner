# Dorner Lead Automation

Clean Streamlit app for Dorner lead processing.

## Files in this ZIP

- `app.py` - Streamlit application and all automation rules
- `requirements.txt` - Python dependencies
- `template.xlsx` - Excel header/template used for output
- `README.md` - setup notes

## What it does

1. Upload one or more Dorner `.msg` files.
2. Extract lead/customer/quote information.
3. Apply Dorner rules:
   - LeadSource1 = Request For Quote
   - Keyword = subject line
   - Dorner CAD vs Config comment logic
   - AquaGard/AquaPruf => Brand Dorner, Product AquaX
   - Garvey => Brand Garvey, Product blank
   - Montratec => Brand Montratec, Product blank
4. The Device column keeps the full long lead section, including CAD files, quote details, General Notes, Additional Spare Parts, footer, and Created line where available.
5. Generates three files with the same base filename from Created date/time:
   - `Dorner_YYYYMMDD_HHMMSS.pdf`
   - `Dorner_YYYYMMDD_HHMMSS.msg`
   - `Dorner_YYYYMMDD_HHMMSS.docx`
6. Creates an updated Excel output ZIP.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

Upload these files directly to GitHub root, then select `app.py` as the Streamlit entry point.
