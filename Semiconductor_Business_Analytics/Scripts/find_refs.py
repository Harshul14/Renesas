import re
from pathlib import Path

scripts = ["xl_part1.py", "xl_part2.py", "common.py", "build_excel.py", "recalc_excel.py"]
for fn in scripts:
    p = Path("Semiconductor_Business_Analytics/Scripts") / fn
    txt = p.read_text(encoding="utf-8")
    matches = re.findall(r"Raw_Data![A-Z0-9$:]+", txt)
    print(f"=== {fn} has {len(matches)} Raw_Data references ===")
    for m in sorted(set(matches)):
        print(f"  {m}")
