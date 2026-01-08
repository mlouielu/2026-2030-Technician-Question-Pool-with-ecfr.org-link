#!/bin/sh

source .venv/bin/activate

python add_ecfr_link.py
soffice --convert-to pdf:writer_pdf_Export --outdir . 2026_Technician_Pool_Linked.docx

deactivate
