# WISE — Water & Irrigation Systems Evaluation

**A web application that replicates Karthik's Excel irrigation upgrade tool for all 50 states — with real EIA, USDA, and climate data.**

Built for the [GLBRC AgLands Project](https://atlas.glbrc.org/) · UW-Madison

---

## What This Does

A farmer opens the app in any browser. They pick their state, county, current irrigation system, crop, and water situation. The app compares three upgrade options — **Center Pivot, Improved Sprinkler, and Drip** — and returns:

- Which system is recommended (ranked by NPV)
- Annual net savings
- Water savings %
- Payback period
- 25-year Net Present Value

All in under 1 second. No Excel required.

---

## Real Data Sources

| Data | Source | Coverage |
|---|---|---|
| Electricity rates | EIA March 2026 | All 51 states |
| Crop market prices | USDA NASS 5-yr avg | 44 states |
| Climate data (ET, Precip, Temp, AWS) | Excel Climate_data sheet | 3,094 counties |
| System specs & costs | Excel SystemTech sheet | 3 upgrade options |
| Production costs | USDA ERS 2025 | 6 crops |
| All calculation formulas | Excel Results_calc sheet | NPV, TDH, energy, payback |

---

## Formulas

All formulas translated directly from Karthik's `Results_calc` sheet:

```
1. TDH = Well Depth + (Pressure × 2.31)
2. Applied Water = Net Crop Water ÷ Efficiency
3. Energy Cost = Water × Acres × TDH × 0.0853 ÷ Pump Efficiency × Elec Rate
4. Energy Savings = Baseline Energy − New Energy
5. Annual Net Savings = Energy Savings − (O&M × Acres)
6. Net Investment = Upgrade Cost × Acres × (1 − Incentive %)
7. Payback = Investment ÷ Annual Savings
8. NPV = Savings × n ÷ (1 + r) − Investment  [growing annuity, r=g=4%]
```

---

## Quick Start

```bash
# Clone the repo
git clone https://github.com/YOUR_USERNAME/wise-irrigation-tool.git
cd wise-irrigation-tool

# Install dependencies
pip install -r requirements.txt

# Run the app
python3 app.py

# Open in browser
# http://localhost:5000
```

---

## Project Structure

```
wise-irrigation-tool/
├── app.py                  # Flask backend — all routes and calculation engine
├── requirements.txt        # Python dependencies (flask, pandas)
├── data/
│   └── app_data.json       # All data: counties, elec prices, crop prices, climate
├── templates/
│   └── index.html          # Frontend — HTML form + JavaScript state management
├── static/
│   └── style.css           # Styling
└── scripts/
    └── generate_data.py    # Regenerate app_data.json from Karthik's Excel file
```

---

## How It Works

```
Browser (HTML form)
    ↓ fetch() — background request, no page reload
Flask routes (app.py)
    ↓ reads from memory (app_data.json loaded at startup)
calculate_wise() — Python calculation engine
    ↓ returns ranked results as JSON
JavaScript — updates results panel live
```

Three routes:
- `GET /get_counties?state=Kansas` → returns county list + electricity rate
- `GET /get_county_data?state=Kansas&county=Finney&crop=Corn` → returns climate data, crop price, production cost
- `POST /calculate` → runs full WISE analysis, returns ranked results

---

## Regenerating Data From Excel

If Karthik updates the Excel file, run:

```bash
# Put the Excel file in the project root, then:
python3 scripts/generate_data.py
```

This extracts electricity prices, crop prices, county lists, and climate data and saves to `data/app_data.json`.

---

## What's Still Manual (Next Steps)

| Input | Current | Next Step |
|---|---|---|
| Expected Yield | Farmer types manually | Connect county_irrigated_yield sheet (45k rows) |
| Farmer inputs | Lost on browser close | Add browser localStorage or PostgreSQL |
| Map interface | Dropdown only | React + vector tiles (Version 2) |
| BE-SMART tool | Not started | Same approach as WISE |

---

## Tech Stack

- **Backend:** Python 3 + Flask
- **Frontend:** HTML + CSS + vanilla JavaScript (fetch API)
- **Data:** JSON file loaded into memory at startup
- **Future:** PostgreSQL (GCP Cloud SQL) + React frontend

---

## Contributing

1. Fork the repo
2. Create a branch: `git checkout -b feature/your-feature-name`
3. Make your changes
4. Push and open a Pull Request

---

## Team

- **Karthik Ramaswamy** — tool design and Excel model
- **Tyler Lark** — research lead, GLBRC Atlas
- **Arav Bhanushali** — web application development

---

## Related

- [GLBRC Atlas](https://atlas.glbrc.org/) — geospatial land use data
- [Atlas Analysis Platform](https://atlasanalysis.glbrc.org/) — React analysis frontend
- BE-SMART — bioenergy crop transition tool (coming soon)
