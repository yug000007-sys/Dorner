# Dorner Lead Automation

Streamlit app for processing one or many Dorner `.msg` / `.txt` leads.

Outputs:
- One Excel file with all rows
- One `.pdf` per lead
- One `.msg` copy per lead
- One `.doc` per lead using Streamlit-compatible RTF content
- One ZIP containing all generated files

Run:
```bash
pip install -r requirements.txt
streamlit run app.py
```
