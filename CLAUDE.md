# WISE — Water & Irrigation Systems Evaluation

Flask irrigation-economics app: `app.py` (engine + routes), `data/app_data.json` (all model
parameters), `templates/index.html` (form + JS), `static/style.css`. Deployed on Render
from `main`. Working branch: `excel-v6-sync`. Live demo: **Dec 11, 2026** (parity needed by
early December).

## Source of truth
- `reference/System_upgrade_v6.xlsm` is the **model of record**. If the app and the workbook
  disagree, the workbook wins. Cite cells in comments, e.g. `# Results_calc!M3`.
- `reference/spec.pdf` (Karthik's P1/P2/P3 fix list + v6 calc spec) is **not in the repo yet**.
  Until it arrives, the calculation is implemented from the workbook formulas directly; every
  judgement call is logged in `NOTES_FOR_KARTHIK.md`. When the PDF lands, re-audit against it.
- `tests/golden_cases.json`: 5 Excel reference cases. **Never edit expected values to make a
  test pass — fix the code.**

## Definition of done
1. `scripts/generate_data.py` (openpyxl, `data_only=True`) exports everything model-related
   into `data/app_data.json`, including `source` = workbook filename + SHA-256 (shown in the
   site footer). Nothing model-related hard-coded in `app.py`.
2. `calculate_wise(inputs, data)` is a pure function (no globals, no I/O, no mutation).
3. `tests/test_golden.py`: all 5 cases within $1 on NPV, annual net savings, investment;
   decision, rank and recommended system exact.
4. UI uses Excel's dropdowns/labels, supports the No Irrigation System path, custom upgrade
   cost, required-field gate, yield/price data-level messages, gross-vs-net 20% warning.
5. README formula section points at the spec/workbook; export script committed; `wise_v4.zip`
   deleted. Landing copy per Dr. Lark (headline / subheading / "Use the Free Tool").

## Calculation map (workbook → engine)
| Step | Workbook | Notes |
|---|---|---|
| Energy price | `INPUTS_UI!I16`, `INPUTS!AI2:AI9` | Elec ¢/kWh÷100 ÷ 0.9 motor eff; NG latest-year $/Mcf ÷ (302 kWh/Mcf × 0.27) |
| TDH | `INPUTS_UI!I17` | 0 well depth for Surface Water, else full depth; + psi × 2.31 |
| Baseline | `Results_calc!A2:I2` | applied = net ÷ current eff (0 for No Irrigation) |
| Upgrades | `Results_calc!A3:T5` | O&M is incremental vs baseline; revenue benefit only for No Irrigation |
| Reinvest/salvage | `Results_calc!X3:X5` | on gross cost (before incentive) |
| NPV | `Results_calc!M3:M5` | growing annuity on energy, flat annuity on O&M and revenue, minus (L + X) |
| Screens | `Results_calc!O:R`, `W` | eligible → water cap → NPV → Candidate |
| Rank / rec / status | `Results_calc!S`, `V2:V11` | |
| Priority / rationale | `RESULTS!D24`, `D26` | |
| Readiness | `INPUTS!AE52:AE81` | required-field checklist |
| Yield / price / cost msgs | `INPUTS!AE6:AE47` | |
| Gross-vs-net warning | `INPUTS!AE84:AE87` | 20% tolerance |

## Commands
- Setup: `python3 -m venv .venv && .venv/bin/pip install -r requirements-dev.txt`
- Regenerate data: `.venv/bin/python scripts/generate_data.py`
- Parity test: `.venv/bin/python -m pytest tests -v`
- Run app: `.venv/bin/python app.py` → http://localhost:5000

## Conventions
- Small, logical commits. Never push to `main` (auto-deploys to Render).
- Re-run the parity test after every engine change.
