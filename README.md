# Dorner Lead Automation

Streamlit app for Dorner lead MSG processing.

## What it does
- Upload one or more `.msg` files
- Creates Excel rows using the required header order
- Adds `GrandTotal` as the last Excel column
- Generates attachments:
  - `.pdf`
  - `.msg`
  - `.doc`
- The `.doc` file is RTF-based and opens in Microsoft Word. This works on Streamlit Cloud without Microsoft Word installed.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy
Upload these flat files to GitHub and deploy `app.py` on Streamlit Cloud.
