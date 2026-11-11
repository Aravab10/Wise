# WISE — Water & Irrigation Systems Evaluation

**A web application that replicates Karthik's Excel irrigation upgrade tool for all 50 states — with real EIA, USDA, and climate data.**

Built for the [GLBRC AgLands Project](https://atlas.glbrc.org/) · UW-Madison

---

## What This Does

A farmer opens the app in any browser, picks their state, county, current irrigation system
(or **No Irrigation System**), crop, and water-supply scenario. WISE screens three upgrade
options — **Center Pivot, Improved Sprinkler, and Drip** — and returns:

- The recommended upgrade (or "No upgrade recommended") with a status line and priority
- Annual net savings, water savings %, user-paid investment, payback
- NPV over the chosen horizon, including equipment replacement and salvage

---

## Model of Record

The calculation is a line-for-line port of Karthik's workbook,
[`reference/System_upgrade_v6.xlsm`](reference/System_upgrade_v6.xlsm). **The workbook and
Karthik's v6 calculation spec (`reference/spec.pdf`) are the authoritative formula
reference** — this README intentionally does not restate the formulas, so it can't drift.
In code, every step in `calculate_wise()` (`app.py`) cites the workbook cell it mirrors
(e.g. `Results_calc!M3`), and `CLAUDE.md` has a step → cell map.

Parity is enforced by `tests/test_golden.py`: five Excel reference cases must match within
$1 on NPV, annual net savings, and investment, with decision, rank, and recommended system
matching exactly.

---

## Data

All model parameters live in `data/app_data.json`, exported from the workbook by
`scripts/generate_data.py`. The file carries a `source` block (workbook filename + SHA-256)
that is shown in the site footer, so you can always tell which workbook the site runs on.

| Data | Workbook sheet |
|---|---|
| System efficiency, cost, O&M, life | `SystemTech` |
| Engineering constants (2.31 ft/psi, 0.0853 kWh/ac-in-ft, drive efficiencies) | `INPUTS` |
| Dropdown lists, water scenarios | `Lists` (named ranges) |
| Electricity prices (EIA) | `ElecPrice` |
| Natural gas prices (EIA, latest year) | `NGP` |
| County / state yields (USDA NASS) | `Yield_county`, `Yield_state` |
| State / national prices (USDA NASS) | `Price_state`, `Price_national` |
| Production cost (display only) | `CropCost` |
| Climate, slope, counties | `Climate_data`, `Slope_data`, `FIPS` |

---

## Quick Start

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements-dev.txt
.venv/bin/python app.py          # http://localhost:5000
.venv/bin/python -m pytest tests # Excel parity + engine tests
```

## Regenerating Data From Excel

When Karthik updates the workbook, save it in Excel (so cached values are current), replace
`reference/System_upgrade_v6.xlsm`, then:

```bash
.venv/bin/python scripts/generate_data.py
.venv/bin/python -m pytest tests
```

Commit the workbook and `data/app_data.json` together.

---

## Project Structure

```
├── app.py                    # Flask routes + pure calculation engine
├── data/app_data.json        # Exported model parameters (generated — don't hand-edit)
├── reference/                # System_upgrade_v6.xlsm (model of record), spec
├── scripts/generate_data.py  # Workbook → app_data.json exporter
├── templates/index.html      # Form, results dashboard, JS
├── static/style.css
└── tests/                    # golden_cases.json, test_golden.py, test_engine.py
```

Routes: `GET /get_counties?state=…`, `POST /resolve` (auto-fill, messages, required-field
check), `POST /calculate` (full analysis).

---

## Tech Stack

- **Backend:** Python 3 + Flask
- **Frontend:** HTML + CSS + vanilla JavaScript (fetch API)
- **Data:** JSON exported from the workbook, loaded into memory at startup
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
