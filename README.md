# Dorner Lead Automation

Clean flat Streamlit app.

## Files
- `app.py` - full app and rules
- `requirements.txt` - Streamlit dependencies
- `README.md` - setup

## Features
- Upload multiple `.msg` files at once
- Creates one Excel row per lead
- Uses required header order
- Adds `GrandTotal` as the final column
- Extracts `GrandTotal` as text like `$14,370.00`
- Extracts `ReceivedDateTime` from MSG/RFC Date header, fallback to `Created:` line
- Creates same base filenames: `Dorner_YYYYMMDD_HHMMSS.pdf/.msg/.docx`
- Generates styled DOCX/PDF attachments
- Device field keeps full long lead section through `Created:`

## Streamlit Cloud
1. Upload these files to GitHub root.
2. In Streamlit Cloud, set main file path to `app.py`.
3. Deploy.
