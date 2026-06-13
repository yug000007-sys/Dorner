# Dorner Lead Automation

Streamlit app for Dorner lead automation.

## Fix in this version
- Fixes the DOCX generation crash caused by `.strip()` being applied to a tuple.
- Fixes Distributor parsing so it uses the real `Distributor:` field and does not capture the greeting text.
- Keeps the styled DOCX/PDF format with orange bars and blue information blocks.

## Files
- `app.py` - Streamlit app and all Dorner rules
- `requirements.txt` - Python dependencies
- `README.md` - setup notes

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
