# Dorner Lead Automation

A simple Streamlit app to process Dorner `.msg` lead emails and append the mapped row to the Excel header/template.

## Files

- `app.py` - full Streamlit app and business logic
- `requirements.txt` - Python dependencies
- `template.xlsx` - Excel template with your required headers

## How to run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## How to deploy on Streamlit Cloud

1. Create a GitHub repository.
2. Upload `app.py`, `requirements.txt`, and `template.xlsx` to the repository root.
3. In Streamlit Cloud, choose the repo and set main file path to `app.py`.

## What the app does

- Upload one or more Dorner `.msg` emails.
- Extract customer contact fields, quote, total, lead time, subject, date, and device text.
- Apply Dorner rules:
  - CAD vs Config comments.
  - AquaGard/AquaPruf -> Brand Dorner and Product AquaX.
  - Garvey -> Brand Garvey and Product blank.
  - Montratec -> Brand Montratec and Product blank.
  - LeadSource1 = Request For Quote.
  - Keyword = email subject line.
  - PDF column contains generated `.pdf`, `.msg`, and `.doc` names.
- File names use `Dorner_YYYYMMDD_HHMMSS` from the message Created date when available.
- Creates a ZIP containing generated PDF, MSG, DOC, and updated Excel workbook.

## Notes

- `.doc` output is HTML-based for Word compatibility.
- If `.msg` parsing fails for a message, paste the lead text into the text box in the app and process from text.
