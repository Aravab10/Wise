#!/usr/bin/env python3
"""
Export every model parameter WISE needs from the v6 workbook into data/app_data.json.

    python3 scripts/generate_data.py [path/to/System_upgrade_v6.xlsm] [-o data/app_data.json]

The workbook is the model of record. Values are read with openpyxl data_only=True (Excel's
cached results), so the workbook must have been saved by Excel after its last recalculation.
Nothing model-related should be hard-coded in app.py; if the engine needs a number, it
belongs here.
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_XLSM = os.path.join(ROOT, "reference", "System_upgrade_v6.xlsm")
DEFAULT_OUT = os.path.join(ROOT, "data", "app_data.json")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def rows(ws, min_row=2):
    for r in ws.iter_rows(min_row=min_row, values_only=True):
        if r and any(v is not None for v in r):
            yield r


def clean(v):
    """Normalise cell values: strip strings, turn integral floats into ints."""
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def put_first(d, key, value):
    """XLOOKUP returns the first match, so the first row for a key wins."""
    if key not in d:
        d[key] = value


def named_list(wb, name):
    """Values of a workbook named range (dropdown source), blanks dropped."""
    dn = wb.defined_names[name]
    out = []
    for sheet, ref in dn.destinations:
        cells = wb[sheet][ref.replace("$", "")]
        if not isinstance(cells, tuple):
            cells = ((cells,),)
        for row in cells:
            for c in (row if isinstance(row, tuple) else (row,)):
                v = clean(c.value)
                if v is not None:
                    out.append(v)
    return out


def sequence_from_formula(wb_formulas, sheet, cell):
    """Evaluate a cell holding SEQUENCE(rows,1,start,step) (e.g. Lists!U2). Its cached
    value depends on other inputs, so read the formula instead."""
    v = wb_formulas[sheet][cell].value
    text = getattr(v, "text", v) or ""
    m = re.search(r"SEQUENCE\((\d+),\s*1,\s*([\d.]+),\s*([\d.]+)\)", text)
    if not m:
        sys.exit(f"Could not parse SEQUENCE in {sheet}!{cell}: {text!r}")
    n, start, step = int(m.group(1)), float(m.group(2)), float(m.group(3))
    return [clean(round(start + i * step, 10)) for i in range(n)]


def export(xlsm_path):
    wb = openpyxl.load_workbook(xlsm_path, data_only=True)
    wbf = openpyxl.load_workbook(xlsm_path, read_only=True)  # formulas, for SEQUENCE lists

    # ── SystemTech (efficiency, cost, O&M, life, role) ─────────────────────
    system_tech = {}
    for name, role, eff, cost, om, life, *_ in rows(wb["SystemTech"]):
        if not name:
            continue
        system_tech[clean(name)] = {
            "role": clean(role),
            "efficiency": clean(eff) or 0,
            "upgrade_cost": clean(cost) or 0,
            "om": clean(om) or 0,
            "life": clean(life),  # None when blank (current-only systems)
        }
    # Lists!V2 / W2: FILTER(SystemTech!A, ISNUMBER(SEARCH("Current"/"Upgrade", Role)))
    current_systems = [s for s, t in system_tech.items() if "current" in (t["role"] or "").lower()]
    upgrade_systems = [s for s, t in system_tech.items() if "upgrade" in (t["role"] or "").lower()]
    # Results_calc!A3:A5: the systems evaluated, in table order
    rc = wb["Results_calc"]
    evaluated = [clean(rc[f"A{r}"].value) for r in range(3, 6)]

    # ── Engineering / economic constants ───────────────────────────────────
    inp = wb["INPUTS"]
    constants = {
        "electric_motor_efficiency": inp["AI2"].value,      # INPUTS!AI2
        "gas_engine_efficiency": inp["AI3"].value,          # INPUTS!AI3
        "ng_kwh_per_mcf": inp["AI4"].value,                 # INPUTS!AI4
        "pressure_head_ft_per_psi": inp["Z54"].value,       # INPUTS!Z54
        "pumping_energy_kwh_per_acin_ft": inp["Z55"].value, # INPUTS!Z55 (= AI5)
        "gross_net_tolerance": inp["AE86"].value,           # INPUTS!AE86
        "elec_price_divisor": 100,                          # INPUTS_UI!I16: ElecPrice!B/100 (cents → $)
    }

    # ── Dropdown lists (named ranges on Lists) ─────────────────────────────
    lists = {
        "states": [s for s in named_list(wb, "State_List") if s != "Select…"],
        "irrigation_sources": named_list(wb, "Irrigation_Source"),
        "energy_sources": named_list(wb, "Energy_Source"),
        "crops": named_list(wb, "Crop_Types"),
        "binary": named_list(wb, "Binary"),
        "system_age": named_list(wb, "System_Age"),
        "well_capacity_gpm": named_list(wb, "Well_Capacity"),
        "operating_pressure_psi": named_list(wb, "Operating_Pressure"),
        "pump_efficiency": named_list(wb, "Pump_Efficiency"),
        "incentive": named_list(wb, "Incentive"),
        "analysis_horizon": named_list(wb, "Analysis_Horizon"),
        "discount_rate": named_list(wb, "Discount_Rate"),
        "energy_escalation": [round(x, 6) for x in named_list(wb, "Energy_Escalation")],
        "gross_irrigation_in": sequence_from_formula(wbf, "Lists", "U2"),
        "current_systems": current_systems,
        "upgrade_systems": upgrade_systems,
        "water_scenarios": [],
    }
    scenarios = {}
    ls = wb["Lists"]
    for r in range(2, 7):  # Lists!K2:L6 (VLOOKUP table for INPUTS_UI!I25)
        label, red = clean(ls[f"K{r}"].value), ls[f"L{r}"].value
        if label:
            lists["water_scenarios"].append(label)
            scenarios[label] = red if isinstance(red, (int, float)) else None  # Custom → user input

    # ── Location ────────────────────────────────────────────────────────────
    state_counties = {}
    for r in rows(wb["FIPS"]):
        st, county = clean(r[0]), clean(r[2])
        if st and county:
            lst = state_counties.setdefault(st, [])
            if county not in lst:
                lst.append(county)
    state_counties = {s: sorted(c) for s, c in sorted(state_counties.items())}

    climate = {}
    for st, county, _fips, et, precip, temp, aws, *_ in rows(wb["Climate_data"]):
        if st and county:
            put_first(climate, f"{st}|{county}",
                      {"et": et, "precip": precip, "temp": temp, "aws": aws})
    slope = {}
    for st, county, _fips, sl, *_ in rows(wb["Slope_data"]):
        if st and county:
            put_first(slope, f"{st}|{county}", sl)

    # ── Energy prices ──────────────────────────────────────────────────────
    ep = wb["ElecPrice"]
    elec_label = clean(ep["B1"].value)
    elec_cents = {clean(r[0]): r[1] for r in rows(ep) if r[0] and isinstance(r[1], (int, float))}

    ngp = wb["NGP"]
    years = [c.value for c in next(ngp.iter_rows(min_row=1, max_row=1))]
    ng_prices = {}
    for r in rows(ngp):
        st = clean(r[0])
        if not st:
            continue
        # INPUTS_UI!I16: LOOKUP(2, 1/(row<>""), row) → last non-blank value in the row
        for idx in range(len(r) - 1, 0, -1):
            if r[idx] not in (None, ""):
                ng_prices[st] = {"price": r[idx], "year": clean(years[idx])}
                break

    # ── Yield, price, production cost ──────────────────────────────────────
    yield_county = {}
    for _fips, st, county, crop, y, unit, yrs, latest, _lvl, status, *_ in rows(wb["Yield_county"]):
        if st and county and crop:
            put_first(yield_county, f"{st}|{county}|{crop}", {
                "yield": y, "unit": clean(unit), "years": clean(yrs),
                "latest_year": clean(latest), "status": clean(status)})
    yield_state = {}
    for st, crop, y, unit, yrs, latest, _lvl, status, *_ in rows(wb["Yield_state"]):
        if st and crop:
            put_first(yield_state, f"{st}|{crop}", {
                "yield": y, "unit": clean(unit), "years": clean(yrs),
                "latest_year": clean(latest), "status": clean(status)})
    price_state = {}
    for st, crop, p, unit, yrs, latest, _lvl, status, *_ in rows(wb["Price_state"]):
        if st and crop:
            put_first(price_state, f"{st}|{crop}", {
                "price": p, "unit": clean(unit), "years": clean(yrs),
                "latest_year": clean(latest), "status": clean(status)})
    price_national = {}
    for crop, p, unit, yrs, latest, _lvl, status, *_ in rows(wb["Price_national"]):
        if crop:
            put_first(price_national, clean(crop), {
                "price": p, "unit": clean(unit), "years": clean(yrs),
                "latest_year": clean(latest), "status": clean(status)})
    crop_cost = {}
    for crop, region, cost, source, year, *_ in rows(wb["CropCost"]):
        if crop:
            put_first(crop_cost, clean(crop), {
                "cost": cost, "unit": "$/acre",  # INPUTS!AE45
                "region": clean(region), "source": clean(source), "year": clean(year)})

    return {
        "source": {
            "workbook": os.path.basename(xlsm_path),
            "sha256": sha256(xlsm_path),
            "generated_at": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "generator": "scripts/generate_data.py",
            "elec_price_label": elec_label,
        },
        "constants": constants,
        "system_tech": system_tech,
        "evaluated_systems": evaluated,
        "water_scenarios": scenarios,
        "lists": lists,
        "state_counties": state_counties,
        "climate": climate,
        "slope": slope,
        "elec_prices_cents": elec_cents,
        "ng_prices": ng_prices,
        "yield_county": yield_county,
        "yield_state": yield_state,
        "price_state": price_state,
        "price_national": price_national,
        "crop_cost": crop_cost,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("xlsm", nargs="?", default=DEFAULT_XLSM)
    ap.add_argument("-o", "--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    data = export(args.xlsm)
    with open(args.out, "w") as f:
        json.dump(data, f, indent=1, sort_keys=False, ensure_ascii=False)
        f.write("\n")
    print(f"Wrote {args.out}")
    print(f"  source   {data['source']['workbook']}  sha256 {data['source']['sha256'][:12]}…")
    for k in ("system_tech", "state_counties", "climate", "slope", "elec_prices_cents",
              "ng_prices", "yield_county", "yield_state", "price_state", "price_national",
              "crop_cost"):
        print(f"  {k:<18} {len(data[k]):>6}")


if __name__ == "__main__":
    main()
