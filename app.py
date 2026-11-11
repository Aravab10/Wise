"""
WISE — Water & Irrigation Systems Evaluation
Flask backend, v6 parity edition.

The calculation engine mirrors reference/System_upgrade_v6.xlsm (the model of record)
cell by cell; comments cite the workbook cell each step comes from. Every model
parameter is read from data/app_data.json, which scripts/generate_data.py exports from
the workbook. Nothing model-related is hard-coded here.

calculate_wise(inputs, data) is a pure function: same inputs + data → same result, no
globals, no I/O, no mutation.
"""
import json
import math
import os
from decimal import ROUND_HALF_UP, Decimal

from flask import Flask, jsonify, render_template, request

NO_IRRIGATION = "No Irrigation System"
SURFACE_WATER = "Surface Water"
READY = "READY FOR ANALYSIS"
NO_UPGRADE = "No upgrade recommended"
NOT_RECOVERED = "Not recovered"
NOT_RECOMMENDED = "Not recommended"
CUSTOM_SCENARIO = "Custom"


def load_data(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ── Excel semantics helpers ───────────────────────────────────────────────

def _text(v):
    """A text cell: None when blank."""
    if v is None:
        return None
    v = str(v).strip()
    return v or None


def _num(v):
    """ISNUMBER-style coercion: a finite float, or None for blank / non-numeric."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None
    try:
        f = float(str(v).strip().replace(",", ""))
    except ValueError:
        return None
    return f if math.isfinite(f) else None


def _round_half_up(x, places=0):
    """Excel rounds half away from zero (Python rounds half to even)."""
    q = Decimal(1).scaleb(-places)
    return float(Decimal(repr(x)).copy_abs().quantize(q, rounding=ROUND_HALF_UP)) * (1 if x >= 0 else -1)


def fmt_fixed(x, places):          # TEXT(x, "0.0")
    return f"{_round_half_up(x, places):.{places}f}"


def fmt_pct(x):                    # TEXT(x, "0%")
    return f"{_round_half_up(x * 100):.0f}%"


def fmt_signed_pct(x):             # TEXT(x, "+0%;-0%")
    r = _round_half_up(x * 100)
    return f"-{abs(r):.0f}%" if r < 0 else f"+{r:.0f}%"


def fmt_money(x):                  # TEXT(x, "$#,##0")
    r = _round_half_up(x)
    return f"-${abs(r):,.0f}" if r < 0 else f"${r:,.0f}"


def fmt_price(x):                  # TEXT(x, "$0.0000")
    return f"${_round_half_up(x, 4):.4f}"


# ── Input resolution (INPUTS_UI auto fields, INPUTS!AE resolution blocks) ──

def _resolve_energy(state, energy_source, data):
    """INPUTS_UI!I16 (energy price in its native unit) and INPUTS!AI6:AI9."""
    c = data["constants"]
    price = unit = label = None
    brake_kwh_per_unit = None
    if energy_source == "Natural Gas":
        rec = data["ng_prices"].get(state or "")
        if rec and _num(rec["price"]) is not None:
            price, unit = float(rec["price"]), "$/Mcf"
            label = f"EIA natural gas price {rec['year']} · {state}"
        brake_kwh_per_unit = c["ng_kwh_per_mcf"] * c["gas_engine_efficiency"]      # AI7
    elif energy_source == "Electricity":
        cents = data["elec_prices_cents"].get(state or "")
        if _num(cents) is not None:
            price, unit = cents / c["elec_price_divisor"], "$/kWh"
            label = f"EIA {data['source'].get('elec_price_label') or 'electricity rate'} · {state}"
        brake_kwh_per_unit = c["electric_motor_efficiency"]                         # AI6
    cost_per_brake_kwh = (price / brake_kwh_per_unit
                          if price is not None and brake_kwh_per_unit else None)   # AI9
    return {"price": price, "unit": unit, "label": label,
            "brake_kwh_per_unit": brake_kwh_per_unit,
            "cost_per_brake_kwh": cost_per_brake_kwh}


def _resolve_yield(state, county, crop, custom_yield, data):
    """INPUTS!AE6:AE19 → effective yield (INPUTS_UI!D25) and its message."""
    cty = data["yield_county"].get(f"{state}|{county}|{crop}") if state and county and crop else None
    sta = data["yield_state"].get(f"{state}|{crop}") if state and crop else None
    if cty and cty["status"] == "County recent":
        source = "County recent"
    elif sta and sta["status"] == "State recent":
        source = "State recent"
    elif cty and cty["yield"] is not None:
        source = "County old"
    elif sta and sta["yield"] is not None:
        source = "State old"
    else:
        source = "None"                                                            # AE14
    rec = cty if source.startswith("County") else sta if source.startswith("State") else None
    auto = _num(rec["yield"]) if rec else None                                    # AE15
    unit = rec["unit"] if rec else None                                           # AE16
    years = rec["years"] if rec else None                                         # AE17

    detail = f"({years}; {unit})"
    if source == "County recent":                                                 # AE18
        data_msg = f"Recent county NASS irrigated yield for {county} {detail}."
    elif source == "State recent":
        data_msg = f"No recent county data - using recent {state} state NASS yield {detail}."
    elif source == "County old":
        data_msg = f"Only older county data available {detail} - review or enter a custom yield."
    elif source == "State old":
        data_msg = f"Only older state data available {detail} - review or enter a custom yield."
    else:
        data_msg = f"No county or state yield data for {crop} in {state} - enter a custom yield below."

    if not (state and county and crop):                                          # AE19
        message = "Select state, county and crop to load yield data."
    elif custom_yield is not None:
        message = (f"Using your custom yield of {fmt_fixed(custom_yield, 1)}"
                   + (f" {unit}" if unit else "") + ". "
                   + (f"No automatic yield is available for {crop} in {state}." if auto is None else
                      f"Automatic yield for {county}, {crop} is {fmt_fixed(auto, 1)} {unit} ({source})"
                      " - clear the custom yield to use it instead."))
    else:
        message = data_msg
    value = custom_yield if custom_yield is not None else auto                    # INPUTS_UI!D25
    return {"value": value, "auto": auto, "unit": unit, "years": years, "source": source,
            "level": "OVERRIDE" if custom_yield is not None else ("NO DATA" if auto is None else "AUTO"),
            "message": message}


def _resolve_price(state, crop, custom_price, data):
    """INPUTS!AE23:AE36 → effective market price (INPUTS_UI!D28) and its message."""
    sta = data["price_state"].get(f"{state}|{crop}") if state and crop else None
    nat = data["price_national"].get(crop) if crop else None
    if sta and sta["status"] == "State recent":
        source = "State recent"
    elif nat and nat["status"] == "National recent":
        source = "National recent"
    elif sta and sta["price"] is not None:
        source = "State old"
    else:
        source = "None"                                                            # AE31
    rec = nat if source == "National recent" else sta if source.startswith("State") else None
    auto = _num(rec["price"]) if rec else None                                    # AE32
    unit = rec["unit"] if rec else None                                           # AE33
    years = rec["years"] if rec else None                                         # AE34

    detail = f"({years}; {unit})"
    if source == "State recent":                                                  # AE35
        data_msg = f"State 5-year average marketing price for {state} {detail}."
    elif source == "National recent":
        data_msg = f"No recent state price - using national 5-year average {detail}."
    elif source == "State old":
        data_msg = f"Only older state price data available {detail} - review or enter a custom price."
    else:
        data_msg = f"No state or national price for {crop} - enter a custom price below."

    if not crop:                                                                  # AE36
        message = "Select a crop to load price data."
    elif custom_price is not None:
        message = f"Using your custom market price of {fmt_price(custom_price)}" + (f" ({unit})" if unit else "") + "."
    else:
        message = data_msg
    value = custom_price if custom_price is not None else auto                    # INPUTS_UI!D28
    return {"value": value, "auto": auto, "unit": unit, "years": years, "source": source,
            "level": "OVERRIDE" if custom_price is not None else ("NO DATA" if auto is None else "AUTO"),
            "message": message}


def _resolve_prod_cost(crop, custom_cost, data):
    """INPUTS!AE39:AE47 (display only — not used in the economics)."""
    rec = data["crop_cost"].get(crop) if crop else None
    auto = _num(rec["cost"]) if rec else None
    if not crop:
        message = "Select a crop to load production cost."
    elif custom_cost is not None:
        message = f"Using your custom production cost of {fmt_money(custom_cost)} per acre."
    elif rec is None:
        message = f"No published production cost for {crop} - enter a custom cost below."
    else:
        message = (f"Budget production cost for {crop}, {rec['region']} region "
                   f"({rec['source']}, {rec['year']}; {rec['unit']}).")
    return {"value": custom_cost if custom_cost is not None else auto, "auto": auto,
            "unit": "$/acre",
            "level": "OVERRIDE" if custom_cost is not None else ("NO DATA" if auto is None else "AUTO"),
            "message": message}


def resolve_inputs(inputs, data):
    """Everything INPUTS_UI / INPUTS work out before Results_calc runs: auto-filled values,
    data-level messages, warnings, and the required-field checklist (INPUTS!AE52:AE81)."""
    t = lambda k: _text(inputs.get(k))          # noqa: E731
    n = lambda k: _num(inputs.get(k))           # noqa: E731
    c, tech = data["constants"], data["system_tech"]

    state, county, crop = t("state"), t("county"), t("crop")
    current, upgrade = t("current_system"), t("upgrade_scenario")
    irr_src, energy_src = t("irrigation_source"), t("energy_source")
    irrigated, scenario = t("irrigated"), t("water_scenario")
    acres, well, psi, pump_eff = n("acres"), n("well_depth_ft"), n("operating_pressure_psi"), n("pump_efficiency")
    gross, net = n("gross_irrigation_in"), n("net_crop_water_in")
    incentive, horizon = n("incentive"), n("horizon_yr")
    disc, esc = n("discount_rate"), n("energy_escalation")
    rainfed = n("rainfed_yield")
    custom_cost = n("custom_upgrade_cost")
    is_no_irr = current == NO_IRRIGATION

    energy = _resolve_energy(state, energy_src, data)
    tdh = (None if psi is None else                                                # INPUTS_UI!I17
           (0.0 if irr_src == SURFACE_WATER else (well or 0.0)) + psi * c["pressure_head_ft_per_psi"])
    yld = _resolve_yield(state, county, crop, n("custom_yield"), data)
    price = _resolve_price(state, crop, n("custom_price"), data)
    prod = _resolve_prod_cost(crop, n("custom_prod_cost"), data)

    if scenario == CUSTOM_SCENARIO:                                                # INPUTS_UI!I25
        reduction = n("custom_reduction")
    else:
        reduction = _num(data["water_scenarios"].get(scenario)) if scenario else None

    uses_gross = irrigated == "Yes" or is_no_irr
    available = (gross * (1 - reduction)                                           # INPUTS_UI!I26
                 if uses_gross and gross is not None and reduction is not None else None)

    upgrade_tech = tech.get(upgrade) if upgrade else None
    upgrade_cost = (custom_cost if custom_cost is not None else                     # INPUTS_UI!I29
                    (_num(upgrade_tech["upgrade_cost"]) if upgrade_tech else None))

    current_eff = _num(tech[current]["efficiency"]) if current in tech else None
    implied_net = (gross * current_eff                                             # INPUTS_UI!D37
                   if uses_gross and gross is not None and current_eff is not None else None)

    gross_net_warning = None                                                       # INPUTS!AE85:AE87
    if not is_no_irr and implied_net is not None and net is not None and net > 0:
        rel = (implied_net - net) / net
        if abs(rel) > c["gross_net_tolerance"]:
            gross_net_warning = (
                f"⚠ Gross irrigation x efficiency ({fmt_fixed(implied_net, 1)} in/ac) differs from "
                f"the net crop water requirement ({fmt_fixed(net, 1)} in/ac) by {fmt_signed_pct(rel)}. "
                "The analysis uses net requirement / efficiency - review both entries.")

    # ── Required-field checklist (INPUTS!AE52:AE78; 1 = missing) ──
    pumping_exempt = is_no_irr and upgrade is None
    counties = data["state_counties"].get(state or "", [])
    checks = [
        ("state", "State", state is None),
        ("county", "County", county is None or county not in counties),
        ("acres", "Field area", acres is None or acres <= 0),
        ("current_system", "Current system", current is None),
        ("irrigation_source", "Irrigation source", not pumping_exempt and irr_src is None),
        ("well_depth_ft", "Well depth",
         not pumping_exempt and irr_src != SURFACE_WATER and (well is None or well < 0)),
        ("operating_pressure_psi", "Operating pressure", not pumping_exempt and (psi is None or psi < 0)),
        ("pump_efficiency", "Pump efficiency", not pumping_exempt and (pump_eff is None or pump_eff <= 0)),
        ("energy_source", "Energy source", not pumping_exempt and energy_src is None),
        ("energy_source", "Energy cost (unsupported state/source)",
         not pumping_exempt and energy["price"] is None),
        ("crop", "Crop type", crop is None),
        ("custom_yield", "Expected yield", yld["value"] is None or yld["value"] <= 0),
        ("custom_price", "Market price", price["value"] is None or price["value"] <= 0),
        ("irrigated", "Irrigated?", irrigated is None),
        ("gross_irrigation_in",
         "Planned irrigation water available" if is_no_irr else "Gross irrigation applied",
         uses_gross and (gross is None or gross <= 0)),
        ("net_crop_water_in", "Net crop water requirement", net is None or net <= 0),
        ("water_scenario", "Water supply scenario", scenario is None),
        ("custom_reduction", "Custom reduction", scenario == CUSTOM_SCENARIO and reduction is None),
        ("upgrade_scenario", "Upgrade scenario", upgrade is None),
        ("custom_upgrade_cost", "Estimated upgrade cost", upgrade_cost is None or upgrade_cost < 0),
        ("incentive", "Incentive / cost share", incentive is None or incentive < 0 or incentive > 1),
        ("horizon_yr", "Analysis horizon", horizon is None or horizon <= 0),
        ("discount_rate", "Discount rate", disc is None or disc < 0),
        ("energy_escalation", "Energy escalation", esc is None or esc < 0),
        ("rainfed_yield", "Rainfed yield", is_no_irr and (rainfed is None or rainfed < 0)),
    ]
    missing = [{"field": f, "label": label} for f, label, bad in checks if bad]
    readiness = READY if not missing else f"⚠ {len(missing)} REQUIRED FIELD(S) MISSING"   # AE81

    # ── Section notes (INPUTS_UI G22, G35, G36, B40) ──
    if scenario is None:
        scenario_msg = "Select a water supply scenario to set the reduction from current supply."
    elif scenario == CUSTOM_SCENARIO:
        scenario_msg = "Custom scenario - enter the reduction percentage below."
    else:
        scenario_msg = (f"Scenario applies a {fmt_pct(reduction or 0)} reduction to current gross "
                        "applied irrigation (Lists scenario table).")
    budget = _num(upgrade_tech["upgrade_cost"]) if upgrade_tech else None
    if custom_cost is not None:
        cost_msg = (f"Using your custom upgrade cost of {fmt_money(custom_cost)}/ac (SystemTech budget "
                    f"for {upgrade or ''} is {fmt_money(budget or 0)}/ac).")
    elif upgrade_cost is None:
        cost_msg = f"No SystemTech upgrade cost for {upgrade or ''} - enter a custom cost above."
    else:
        cost_msg = (f"SystemTech budget cost for {upgrade}. Net of {fmt_pct(incentive or 0)} incentive = "
                    f"{fmt_money(upgrade_cost * (1 - (incentive or 0)))}/ac.")
    if gross_net_warning:
        net_msg = gross_net_warning
    elif net is None:
        net_msg = "Manual entry required - no source dataset for net crop water requirement."
    else:
        net_msg = (f"Manual entry ({fmt_fixed(net, 1)} in/ac) - no source dataset; revise if local "
                   "ET-based guidance is available.")

    climate = data["climate"].get(f"{state}|{county}") if state and county else None
    return {
        "readiness": readiness,
        "ready": not missing,
        "missing": missing,
        "missing_message": ("Configure the farm, irrigation system, crop, and scenario - then compare "
                            "upgrade options." if not missing else
                            "Missing required inputs: " + ", ".join(m["label"] for m in missing)),
        "auto": {
            "energy_price": energy["price"], "energy_unit": energy["unit"],
            "energy_label": energy["label"], "cost_per_brake_kwh": energy["cost_per_brake_kwh"],
            "tdh": tdh,
            "yield": yld["value"], "yield_auto": yld["auto"], "yield_unit": yld["unit"],
            "yield_source": yld["source"], "yield_level": yld["level"],
            "price": price["value"], "price_auto": price["auto"], "price_unit": price["unit"],
            "price_source": price["source"], "price_level": price["level"],
            "prod_cost": prod["value"], "prod_cost_auto": prod["auto"], "prod_cost_level": prod["level"],
            "reduction": reduction, "available_water": available,
            "upgrade_cost": upgrade_cost, "upgrade_cost_budget": budget,
            "current_net_irrigation": implied_net, "current_efficiency": current_eff,
            "climate": climate,
            "slope": data["slope"].get(f"{state}|{county}") if state and county else None,
        },
        "messages": {
            "yield": yld["message"], "price": price["message"], "prod_cost": prod["message"],
            "gross_net_warning": gross_net_warning, "scenario": scenario_msg,
            "upgrade_cost": cost_msg, "net_water": net_msg,
            "same_system": ("Upgrade scenario matches the current system - no efficiency gain will be "
                            "modelled." if upgrade and upgrade == current else None),
        },
    }


# ── Economics (Results_calc) ───────────────────────────────────────────────

def _annuity(x, r, n):
    """PV of a level annual amount (Results_calc!M pvOM / pvRev)."""
    return x * n if r == 0 else x * (1 - (1 + r) ** (-n)) / r


def _growing_annuity(x, r, g, n):
    """PV of an amount escalating at g (Results_calc!M pvEnergy)."""
    if abs(r - g) < 1e-7:
        return x * n / (1 + r)
    return x * (1 - ((1 + g) / (1 + r)) ** n) / (r - g)


def _reinvest_minus_salvage(life, n, r, cost):
    """Results_calc!X3:X5 — PV of replacements inside the horizon minus PV of salvage at n."""
    if life is None or life <= 0 or cost is None:
        return 0.0
    k = math.ceil(n / life) - 1
    pv_rep = sum(cost / (1 + r) ** (i * life) for i in range(1, k + 1)) if k > 0 else 0.0
    salvage = cost * ((k + 1) * life - n) / life / (1 + r) ** n
    return pv_rep - salvage


def calculate_wise(inputs, data):
    """Run the full v6 analysis. Pure: depends only on `inputs` and `data`."""
    res = resolve_inputs(inputs, data)
    out = dict(res, baseline=None, systems=[], recommended_system="Complete required inputs",
               status=res["readiness"], priority="Pending", payback=None,
               rationale="Complete all required inputs before using the upgrade recommendation.",
               note="Warning: Complete required inputs before relying on upgrade results.",
               best_water_option=None, max_water_savings=None,
               required_reduction=res["auto"]["reduction"], shortfall=None, key_metrics=None)
    if not res["ready"]:
        return out

    n_ = lambda k: _num(inputs.get(k))          # noqa: E731
    c, tech, a = data["constants"], data["system_tech"], res["auto"]
    current, upgrade = _text(inputs["current_system"]), _text(inputs["upgrade_scenario"])
    is_no_irr = current == NO_IRRIGATION
    acres, pump_eff, net = n_("acres"), n_("pump_efficiency"), n_("net_crop_water_in")
    incentive, n = n_("incentive"), n_("horizon_yr")
    r, g = n_("discount_rate"), n_("energy_escalation")
    tdh, cost_kwh = a["tdh"], a["cost_per_brake_kwh"]
    k_energy = c["pumping_energy_kwh_per_acin_ft"]

    def energy_cost(applied):                                                      # Results_calc!I
        return applied * acres * tdh * k_energy / pump_eff * cost_kwh

    # Baseline row (Results_calc row 2)
    bt = tech[current]
    b_eff, b_om = _num(bt["efficiency"]) or 0.0, _num(bt["om"]) or 0.0
    b_water = 0.0 if b_eff == 0 else net / b_eff                                   # F2
    b_energy = energy_cost(b_water)                                                # I2
    out["baseline"] = {"system": current, "eff": b_eff, "om": b_om,
                       "applied_water": b_water, "energy_cost": b_energy}

    # Crop revenue benefit — only when irrigating a rainfed field (Results_calc!T)
    if is_no_irr:
        revenue = max(0.0, a["yield"] - n_("rainfed_yield")) * a["price"] * acres
    else:
        revenue = 0.0
    water_cap = n_("gross_irrigation_in") if is_no_irr else b_water * (1 - a["reduction"])   # W

    rows = []
    for name in data["evaluated_systems"]:
        st = tech[name]
        eff, om, life = _num(st["efficiency"]), _num(st["om"]) or 0.0, _num(st["life"])
        cost = a["upgrade_cost"] if name == upgrade else _num(st["upgrade_cost"])   # C
        applied = net / eff                                                         # F
        saved = b_water - applied                                                   # G
        saved_pct = saved / b_water if b_water else 0.0                             # H
        e_cost = energy_cost(applied)                                               # I
        e_save = b_energy - e_cost                                                  # J
        om_delta = (om - b_om) * acres
        ans = e_save - om_delta + revenue                                           # K
        invest = cost * acres * (1 - incentive)                                     # L
        x = _reinvest_minus_salvage(life, n, r, cost * acres)                       # X
        npv = (_growing_annuity(e_save, r, g, n) - _annuity(om_delta, r, n)         # M
               + _annuity(revenue, r, n) - (invest + x))
        if ans <= 0 or invest / ans > n:                                            # N
            payback = NOT_RECOVERED
        else:
            payback = invest / ans
        eligible = "Eligible" if eff > b_eff else "Not eligible"                    # O
        if eligible != "Eligible":
            meets, econ = "—", "—"
        else:
            meets = "Meets" if applied <= water_cap else "Does not meet"            # P
            econ = "Positive NPV" if npv > 0 else "Negative NPV"                    # Q
        if eligible != "Eligible":                                                  # R
            decision = "Not eligible"
        elif meets != "Meets":
            decision = "Fails water constraint"
        elif npv <= 0:
            decision = "Economically unattractive"
        else:
            decision = "Candidate"
        rows.append({
            "system": name, "eff": eff, "upgrade_cost": cost, "om": om, "life": life,
            "applied_water": applied, "water_saved": saved, "water_saved_pct": saved_pct,
            "energy_cost": e_cost, "energy_savings": e_save, "om_change": om_delta,
            "revenue_benefit": revenue, "annual_net_savings": ans, "investment": invest,
            "gross_investment": cost * acres, "reinvest_minus_salvage_pv": x, "npv": npv,
            "payback": payback, "eligible": eligible, "meets_water": meets,
            "econ_status": econ, "decision": decision, "water_cap": water_cap,
        })

    candidates = [row for row in rows if row["decision"] == "Candidate"]
    for row in rows:                                                                # S
        row["rank"] = (1 + sum(1 for o in candidates if o["npv"] > row["npv"])
                       if row["decision"] == "Candidate" else NOT_RECOMMENDED)
    rec = max(candidates, key=lambda row: row["npv"]) if candidates else None       # V2 (first max)
    for row in rows:
        row["recommended"] = row is rec

    max_h = max(row["water_saved_pct"] for row in rows)                             # V5
    best_water = next(row["system"] for row in rows if row["water_saved_pct"] == max_h)   # V4
    if rec:                                                                         # V3
        status = "Recommended"
    elif not any(row["meets_water"] == "Meets" for row in rows):
        status = "No evaluated upgrade meets the selected water constraint"
    elif not any(row["meets_water"] == "Meets" and row["econ_status"] == "Positive NPV" for row in rows):
        status = "Water-feasible alternatives are not economically attractive"
    else:
        status = "No eligible upgrade available"

    if rec is None:                                                                 # RESULTS!D24
        priority = "Low"
    elif rec["npv"] <= 0:
        priority = "Low"
    elif isinstance(rec["payback"], float) and rec["payback"] <= n / 2:
        priority = "High"
    else:
        priority = "Medium"   # Excel compares text > number, so "Not recovered" → Medium

    if rec is None or rec["npv"] <= 0:                                              # RESULTS!D26
        rationale = ("No upgrade has a positive NPV under the selected assumptions. Keeping the "
                     "current system is recommended unless water conservation is the primary objective.")
    else:
        pb = rec["payback"] if isinstance(rec["payback"], str) else fmt_fixed(rec["payback"], 1)
        rationale = (f"{rec['system']} provides the highest positive NPV among the evaluated systems. "
                     f"It reduces applied water use by {fmt_pct(rec['water_saved_pct'])} and generates "
                     f"estimated annual net savings of {fmt_money(rec['annual_net_savings'])} per year, "
                     f"with a payback period of {pb} years.")

    out.update(
        baseline=out["baseline"], systems=rows,
        recommended_system=rec["system"] if rec else NO_UPGRADE,
        status=status, priority=priority, rationale=rationale,
        payback=rec["payback"] if rec else NOT_RECOVERED,                           # RESULTS!D25
        note=("Warning: No evaluated upgrade has positive NPV under the selected assumptions."  # P14
              if rec is None else
              "Results are based on current inputs and selected economic assumptions."),
        best_water_option=best_water, max_water_savings=max_h,
        shortfall=max(0.0, a["reduction"] - max_h),                                 # V7
        key_metrics=None if rec is None else {                                      # RESULTS!B7:L10
            "water_saved_pct": rec["water_saved_pct"],
            "annual_net_savings": rec["annual_net_savings"],
            "investment": rec["investment"], "npv": rec["npv"], "payback": rec["payback"],
            "current_water": b_water, "upgraded_water": rec["applied_water"],
            "energy_savings": rec["energy_savings"],
        },
    )
    return out


# ── Flask app ─────────────────────────────────────────────────────────────

app = Flask(__name__)
DATA = load_data(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "app_data.json"))


@app.route("/")
def index():
    return render_template("index.html", lists=DATA["lists"], source=DATA["source"],
                           system_tech=DATA["system_tech"],
                           evaluated_systems=DATA["evaluated_systems"])


@app.route("/get_counties")
def get_counties():
    state = request.args.get("state", "")
    return jsonify({"counties": DATA["state_counties"].get(state, [])})


@app.route("/resolve", methods=["POST"])
def resolve():
    """Live auto-fill, data-level messages and the required-field gate for the form."""
    return jsonify(resolve_inputs(request.get_json(silent=True) or {}, DATA))


@app.route("/calculate", methods=["POST"])
def calculate():
    inputs = request.get_json(silent=True) or {}
    result = calculate_wise(inputs, DATA)
    if not result["ready"]:
        return jsonify(result), 400
    return jsonify(result)


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
