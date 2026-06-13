# Dorner Lead Automation

Clean Streamlit app for Dorner lead processing.

## Files
- `app.py` - Streamlit app and all processing logic
- `requirements.txt` - Python dependencies
- `template.xlsx` - Excel output header template
- `README.md` - this guide

## DOCX fix in this version
The generated DOCX now matches the user's sample style:
- starts directly with `Dorner Distributor,`
- no `DORNER LEAD DOCUMENT` heading
- no file-name block
- no subject block
- no visible URL/mailto/tracking lines
- keeps the full original Dorner email body including quote details, general notes, spare parts, footer, and Created line

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud
Upload all four files to the root of your GitHub repository, then deploy `app.py`.
