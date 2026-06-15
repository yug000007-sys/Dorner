# Dorner Lead Automation

Upload one or many Dorner `.msg` or `.txt` lead files. The app creates:

- One Excel output with all rows
- Styled DOCX for each lead
- PDF for each lead
- MSG copy for each uploaded file
- ZIP containing all outputs

Important fixes in this version:

- Downloaded Excel uses the requested header order.
- GrandTotal is included in the downloaded Excel and formatted as text like `$14,370.00`.
- ReceivedDateTime is taken from the raw MSG/email Date header first, converted to Eastern time, and formatted like `5/27/2026 1:54 PM`.
