# Notes for Karthik — WISE v6 parity

Branch `excel-v6-sync`. The engine now reproduces `System_upgrade_v6.xlsm` on all five golden
cases. The largest difference on NPV, annual net savings or investment is about $4×10⁻¹⁰, and
decision, rank, recommendation, status and priority all match exactly.

## 0. The spec PDF never reached the repo

`reference/spec.pdf` (the P1/P2/P3 list and the v6 calculation spec) was not in the repository
during this work. I implemented the calculation **directly from the workbook formulas**
(Results_calc, RESULTS, INPUTS, INPUTS_UI), using the golden cases as the acceptance test. So:

- The calculation should match the spec wherever the spec matches the workbook.
- I could not check the P1/P2/P3 items one by one against your list. Any P-item that asks for
  something the workbook doesn't do (for example a UI behaviour or wording) may still be open.
  Please send the PDF so I can audit against it before Dec 11.

## 1. Questions about workbook behaviour (I followed the workbook in every case)

1. **No Irrigation System ignores the water scenario.** `Results_calc!W3:W5` caps water at
   `INPUTS_UI!D36` (planned water). It does not use `I26` = D36 × (1 − reduction). With
   "Mild reduction (10%)" and 12 in planned, the cap is 12.0, not 10.8. Golden case 4 confirms
   this is what Excel does. Is it intended?
2. **Reinvestment and salvage use the gross cost** (`X3:X5` uses C × acres), so the cost
   share does not reduce replacement purchases. Payback (`N`) uses user-paid investment `L`
   only and ignores replacements. Is that intended?
3. **Escalation applies only to energy.** Energy savings are a growing annuity, while O&M and
   crop revenue are level annuities. Just confirming.
4. **Status line when nothing is eligible.** If the current system is Drip, no upgrade is
   more efficient, so every row is "Not eligible" and `P` shows "—". `V3` then reports
   "No evaluated upgrade meets the selected water constraint", which is misleading. Its last
   branch, "No eligible upgrade available", can never be reached, because a row that meets the
   constraint with positive NPV is always a Candidate. Should the "Not eligible" case get its
   own message?
5. **Water savings for a rainfed field show 0%.** `H` = IFERROR(G/F2, 0) with F2 = 0, so the
   recommendation text says "reduces applied water use by 0%". The app shows the same text and
   adds "0.0 → 11.8 in/ac/yr" under the metric. I hid the "best-option shortfall" note (from
   helper cell `V7`) on this path because it read as nonsense.
6. **Priority when payback is "Not recovered".** `RESULTS!D24` compares text to a number.
   Excel treats text as larger, so a positive-NPV candidate with no payback inside the horizon
   gets **Medium**. I replicated that.
7. **Gross-irrigation dropdown with No Irrigation System.** `Lists!U2` is empty unless
   Irrigated? = Yes. On the No Irrigation path (Irrigated? = No), "Planned irrigation water
   available" is required but its dropdown is empty in Excel. The app enables the 0–48 in list
   whenever Irrigated? = Yes **or** the current system is No Irrigation System.
8. **Gross-vs-net check is strictly greater than 20%** (`INPUTS!AE87`). Exactly 20% gives no
   warning.
9. **Price fallback order** (`INPUTS!AE31`): state recent → national recent → state old. A
   "National old" price is never used.

## 2. Judgement calls in the web app

- **Irrigation source** now uses the workbook list (Surface Water / Groundwater /
  Surface + Groundwater). Surface + Groundwater uses the full well depth (`INPUTS_UI!I17`), so
  I removed the earlier groundwater-blend slider. The "distance from surface water" field is
  kept as optional and marked "not used in calculations". Tell me if it should be removed or
  should enter TDH.
- **Defaults:** operating pressure (40 psi) and pump efficiency (75%) come preselected, per
  the notes in `INPUTS!H25/H26`. Every other required field starts blank, as in Excel, and
  must be chosen before Calculate runs.
- **System age and pumping capacity** are collected but unused, as in the workbook
  (`INPUTS!AE56` "not used in calculations").
- **Production cost** is display-only (`INPUTS!AE66`).
- **"Other (manual entry)" crop** has no NASS data, so it requires a custom yield and a custom
  price, which is the Excel behaviour.
- **Natural gas** uses the latest non-blank year in `NGP` for each state (2025 for all states
  today), matching `INPUTS_UI!I16`.
- **Counties** are validated against `FIPS`, which reproduces the "RE-PICK" logic in
  `INPUTS_UI!E8`.

## 3. Data refresh

`scripts/generate_data.py` reads the cached values (openpyxl `data_only=True`), so **save the
workbook in Excel before exporting**. The footer shows the workbook filename and the first 12
characters of its SHA-256, so we can always tell which workbook the site is running.
