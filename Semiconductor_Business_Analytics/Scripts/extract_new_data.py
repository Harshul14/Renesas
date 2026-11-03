"""
extract_new_data.py: Inspect and extract all FY2021 and FY2022 data from
'Renesas Industrial Strategy Analysis_with 2021 and 2022 data.pdf'
"""
import sys
from pathlib import Path
from pypdf import PdfReader

PDF_PATH = Path("../Renesas Industrial Strategy Analysis_with 2021 and 2022 data.pdf")
if not PDF_PATH.exists():
    PDF_PATH = Path("Renesas Industrial Strategy Analysis_with 2021 and 2022 data.pdf")

print(f"Reading PDF from {PDF_PATH.resolve()}...")
reader = PdfReader(str(PDF_PATH))
print(f"Total pages: {len(reader.pages)}")

# Extract all text and dump pages mentioning 2021 or 2022
matches = []
full_text = []

for idx, page in enumerate(reader.pages):
    txt = page.extract_text() or ""
    full_text.append(f"--- PAGE {idx+1} ---\n{txt}")
    if any(k in txt for k in ["2021", "2022", "FY21", "FY22", "FY2021", "FY2022"]):
        matches.append((idx + 1, txt))

print(f"\nFound {len(matches)} pages mentioning 2021 or 2022:")
for p_num, _ in matches:
    print(f"  - Page {p_num}")

# Save full text to file for thorough inspection
with open("extracted_new_pdf_text.txt", "w", encoding="utf-8") as f:
    f.write("\n\n".join(full_text))

print("\nFull text saved to extracted_new_pdf_text.txt.")

# Print sections on pages with 2021/2022
print("\n" + "="*80)
print("DETAILED CONTENT FROM MATCHING PAGES:")
print("="*80)

for p_num, txt in matches:
    print(f"\n>>> PAGE {p_num} <<<")
    for line in txt.split("\n"):
        if any(k in line.lower() for k in ["2021", "2022", "fy21", "fy22", "fy2021", "fy2022", "automotive", "industrial", "revenue", "operating", "gross", "margin", "segment", "table"]):
            print(f"  {line}")
