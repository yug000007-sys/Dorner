# Dorner Lead Automation

Flat Streamlit project for Dorner lead processing.

## What it does

- Upload `.msg` or `.txt` Dorner lead
- Extract customer, quote, product and device details
- Apply Dorner rules:
  - AquaGard/AquaPruf -> Brand `Dorner`, Product `AquaX`
  - Garvey -> Brand `Garvey`, Product blank
  - Montratec -> Brand `Montratec`, Product blank
  - LeadSource1 -> `Request For Quote`
  - Keyword -> subject line
  - LeadComments -> CAD or Config comment
- Creates files named like `Dorner_YYYYMMDD_HHMMSS`
  - `.docx`
  - `.pdf`
  - `.msg`
- Creates Excel output with the `PDF` column containing all three filenames
- Device column keeps full long section, including spare parts and Created line

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Cloud

Upload these files to GitHub root:

- `app.py`
- `requirements.txt`
- `README.md`

Then set Streamlit Cloud main file path to:

```text
app.py
```

## Important

The DOCX and PDF are template-style outputs with orange bars, blue info sections and formatted product tables, not plain text dumps.
