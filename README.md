# Renesas Electronics: Semiconductor Business Analytics Case

> **Comprehensive Excel Financial Model & 3-Slide Executive Strategy Presentation**  
> **Central Management Question:** *"Where should the semiconductor company focus to maximize revenue growth while managing business risk?"*  
> **Key Recommendation:** Focus incremental commercial resources, FAE deployment, and software ecosystem on **Industrial (Robotics & Physical AI)**, while defending the **Automotive** cash cow and resolving the margin paradox via **Altium 365**.

---

## 📁 Repository Structure

```
.
├── Semiconductor_Business_Analytics/
│   ├── Semiconductor_Business_Analytics.xlsx             # 12-sheet dynamic financial model (FY2021–FY2026E, 0 errors, 34 audit checks)
│   ├── Semiconductor_Business_Analytics_Presentation.pptx # 3-slide executive PowerPoint (16:9 widescreen, 6-yr baseline)
│   ├── Semiconductor_Business_Analytics_Presentation.pdf  # Vector PDF presentation export
│   ├── README_Setup_and_Methodology.pdf                  # 7-page comprehensive setup & methodology manual with FY21-22 extension
│   ├── Final_Deliverables/                               # Mirrored verified deliverables
│   ├── Backup_Pre_FY21_22/                               # Pre-extension backup of deliverables
│   ├── Scripts/                                          # Python build & recalculation automation suite
│   │   ├── build_excel.py                                # Builds 12-sheet Excel model
│   │   ├── recalc_excel.py                               # Recalculates via Excel COM & audits formulas
│   │   ├── build_powerpoint.py                           # Builds 3-slide presentation & exports to PDF
│   │   ├── build_readme_pdf.py                           # Builds methodology PDF manual
│   │   ├── common.py                                     # Shared styles, palettes & constants
│   │   ├── xl_part1.py                                   # Sheets: Raw_Data, Revenue, Profit, Growth, Risk, Forecast
│   │   └── xl_part2.py                                   # Sheets: Scorecard, Resource, Sensitivity, Validation, Dashboard
│   ├── Source_Files/                                     # Original and updated research dossiers (PDF & Markdown)
│   └── Exports/                                          # Build artifacts
├── Renesas Industrial Strategy Analysis.pdf
├── Renesas Industrial Strategy Analysis.md
├── Renesas Industrial Strategy Analysis_with 2021 and 2022 data.pdf
├── .gitignore
└── README.md
```

---

## 🚀 60-Second Reproducibility Runbook

### Prerequisites
* Windows 10 / 11 (64-bit)
* Python 3.10+ (tested on Python 3.14.6)
* Microsoft 365 Excel & PowerPoint (Desktop)

### Installation
```powershell
pip install openpyxl xlsxwriter python-pptx reportlab pywin32 pypdf pandas pillow
```

### Build Everything From Source
From `Semiconductor_Business_Analytics/`:
```powershell
# 1. Generate full 12-sheet Excel workbook
py Scripts/build_excel.py

# 2. Recalculate all formulas in Excel COM and verify 0 errors
py Scripts/recalc_excel.py

# 3. Generate 3-slide executive PowerPoint presentation and export to PDF
py Scripts/build_powerpoint.py

# 4. Generate 6-page comprehensive methodology PDF manual
py Scripts/build_readme_pdf.py
```

---

## 📊 Strategic Highlights & 6-Year Historical Baseline

* **Continuous 6-Year Timeline:** Extended from FY2021 through FY2026E run-rate. Consolidated revenue grew at **+7.37% full-cycle CAGR** (FY21-25) across 4 elapsed intervals (from 994.2 bn to 1,321.2 bn JPY).
* **The Post-COVID Supercycle (FY21–FY22):** Revenue surged +51.3% YoY in FY22 to 1,503.8 bn JPY with peak operating margin of 37.2%. I&I exploded +64.1% YoY (515.5 to 845.9 bn JPY).
* **Inventory Reset vs Structural Decline:** Proves the FY24 downcycle (-16.7% I&I) was an industry-wide channel digestion after the 2022 peak, not structural decay. I&I is now firmly rebounding (+5.5% FY25, +14.9% FY26E run-rate).
* **Structural Gross Margin Floor:** Non-GAAP gross margin expanded permanently by **+440 bps** from 53.2% (FY21) to 57.6% (FY25) and 58.1% (FY26E), proving catalog MCU and Dialog PMIC integration sustained pricing power.
* **Scorecard Winner:** **Industrial (7.25 / 10)** beats IoT (6.28), Automotive (5.33), and Consumer (3.95).
* **Top Submarket:** **Robotics & Physical AI (8.00 / 10)** leads AI Data Centres (7.15) and Medical (6.70). TAM expanding from $10.9B to $38.9B by 2030 (~29% CAGR).
* **The Margin Paradox:** Industrial commands a **+580 bps Gross Margin premium (60.6% vs ~54.8%)**, but historically lagged in Operating Margin (27.2% vs ~31.5%) due to high direct FAE overhead.
* **The Altium Multiplier:** The $5.9B Altium acquisition provides a cloud-based self-service reference design funnel that collapses customer PCB cycles from 6 months to 2 weeks, converting high gross margin into 29%+ operating margin.
* **Capital Rebalance:** 55% Industrial (45% of which goes to Robotics), 25% Automotive (defensive SDV cash cow), 15% IoT, 5% Consumer.
* **FY2027 Target:** 365.0 bn JPY Base Case Industrial Revenue (+14.1% YoY).

