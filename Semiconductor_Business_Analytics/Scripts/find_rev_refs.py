import re
from pathlib import Path

scripts = ["xl_part1.py", "xl_part2.py", "build_excel.py", "recalc_excel.py"]
for fn in scripts:
    p = Path("Semiconductor_Business_Analytics/Scripts") / fn
    txt = p.read_text(encoding="utf-8")
    matches = re.findall(r"Revenue_Analysis![A-Z0-9$:]+", txt)
    print(f"=== {fn} has {len(matches)} Revenue_Analysis references ===")
    for m in sorted(set(matches)):
        print(f"  {m}")
