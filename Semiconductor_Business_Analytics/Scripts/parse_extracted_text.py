"""
parse_extracted_text.py: Print exact text from relevant pages of the new PDF
handling unicode cleanly.
"""
from pathlib import Path

with open("extracted_new_pdf_text.txt", "r", encoding="utf-8", errors="ignore") as f:
    text = f.read()

pages = text.split("--- PAGE ")
target_pages = [3, 6, 7, 8, 21]

for p in pages:
    if not p.strip():
        continue
    try:
        p_num = int(p.split(" ---")[0])
    except Exception:
        continue
    if p_num in target_pages:
        out_file = f"page_{p_num}_clean.txt"
        with open(out_file, "w", encoding="utf-8") as out:
            out.write(p)
        print(f"Wrote page {p_num} to {out_file} ({len(p)} characters)")
