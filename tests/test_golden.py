"""
Excel-parity test: every golden case in tests/golden_cases.json must reproduce
reference/System_upgrade_v6.xlsm.

Tolerances (definition of done):
  * NPV, annual net savings, investment: within $1
  * decision, rank, recommended system: exact

Expected values come from Excel and are NEVER edited to make this pass.
"""
import copy
import json
import math
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from app import calculate_wise, load_data  # noqa: E402

with open(os.path.join(ROOT, "tests", "golden_cases.json")) as f:
    CASES = json.load(f)

DATA = load_data(os.path.join(ROOT, "data", "app_data.json"))

# golden_cases.json input name -> form field / engine input name.
# The web form posts the workbook's own field names, so this is 1:1 today; the map
# is kept explicit so a rename on either side fails loudly here.
GOLDEN_TO_FORM = {
    "state": "state",                                   # INPUTS_UI!D7
    "county": "county",                                 # INPUTS_UI!D8
    "acres": "acres",                                   # INPUTS_UI!D10
    "current_system": "current_system",                 # INPUTS_UI!I7
    "pumping_capacity_gpm": "pumping_capacity_gpm",     # INPUTS_UI!I9 (not used in calc)
    "irrigation_source": "irrigation_source",           # INPUTS_UI!I10
    "well_depth_ft": "well_depth_ft",                   # INPUTS_UI!I11
    "operating_pressure_psi": "operating_pressure_psi", # INPUTS_UI!I12
    "pump_efficiency": "pump_efficiency",               # INPUTS_UI!I13
    "energy_source": "energy_source",                   # INPUTS_UI!I15
    "crop": "crop",                                     # INPUTS_UI!D23
    "rainfed_yield": "rainfed_yield",                   # INPUTS_UI!D24
    "custom_price": "custom_price",                     # INPUTS_UI!D29
    "irrigated": "irrigated",                           # INPUTS_UI!D34
    "gross_irrigation_in": "gross_irrigation_in",       # INPUTS_UI!D36
    "net_crop_water_in": "net_crop_water_in",           # INPUTS_UI!D39
    "water_scenario": "water_scenario",                 # INPUTS_UI!I23
    "upgrade_scenario": "upgrade_scenario",             # INPUTS_UI!I28
    "incentive": "incentive",                           # INPUTS_UI!I30
    "horizon_yr": "horizon_yr",                         # INPUTS_UI!I32
    "discount_rate": "discount_rate",                   # INPUTS_UI!I33
    "energy_escalation": "energy_escalation",           # INPUTS_UI!I34
}

UPGRADES = ["Center Pivot", "Improved Sprinkler", "Drip"]
DOLLAR_TOL = 1.0


def to_form(golden_inputs):
    unknown = set(golden_inputs) - set(GOLDEN_TO_FORM)
    assert not unknown, f"golden input(s) with no form mapping: {sorted(unknown)}"
    return {GOLDEN_TO_FORM[k]: v for k, v in golden_inputs.items()}


def norm_rank(r):
    """Excel stores numeric ranks as floats (1.0); text ranks stay text."""
    if isinstance(r, (int, float)) and not isinstance(r, bool):
        return int(r)
    return r


def run(case):
    return calculate_wise(to_form(case["inputs"]), DATA)


def by_system(result):
    return {row["system"]: row for row in result["systems"]}


ids = [c["name"] for c in CASES]


@pytest.mark.parametrize("case", CASES, ids=ids)
def test_recommended_system(case):
    assert run(case)["recommended_system"] == case["excel_auto"]["recommended"]


@pytest.mark.parametrize("case", CASES, ids=ids)
@pytest.mark.parametrize("system", UPGRADES)
def test_dollars_within_one(case, system):
    exp = case["excel_results"][system]
    got = by_system(run(case))[system]
    for field in ("npv", "annual_net_savings", "investment"):
        assert math.isclose(got[field], exp[field], abs_tol=DOLLAR_TOL), (
            f"{system} {field}: app {got[field]:.2f} vs Excel {exp[field]:.2f} "
            f"(diff {got[field] - exp[field]:+.2f})")


@pytest.mark.parametrize("case", CASES, ids=ids)
@pytest.mark.parametrize("system", UPGRADES)
def test_decision_and_rank_exact(case, system):
    exp = case["excel_results"][system]
    got = by_system(run(case))[system]
    assert got["decision"] == exp["decision"]
    assert norm_rank(got["rank"]) == norm_rank(exp["rank"])


# ── Supporting parity checks (diagnostics for the headline checks above) ──

@pytest.mark.parametrize("case", CASES, ids=ids)
def test_auto_values(case):
    exp = case["excel_auto"]
    res = run(case)
    assert res["readiness"] == exp["readiness"]
    assert math.isclose(res["auto"]["energy_price"], exp["energy_price"], rel_tol=1e-9)
    assert math.isclose(res["auto"]["tdh"], exp["tdh"], rel_tol=1e-9)
    assert math.isclose(res["auto"]["yield"], exp["yield_"], rel_tol=1e-9)
    assert math.isclose(res["auto"]["price"], exp["price"], rel_tol=1e-9)
    assert res["status"] == exp["status"]
    assert (res["messages"]["gross_net_warning"] or "") == exp["gross_net_warning"]


@pytest.mark.parametrize("case", CASES, ids=ids)
def test_intermediates(case):
    """Row-level intermediates, so a dollar mismatch can be traced to its cause."""
    res = run(case)
    rows = by_system(res)
    b_exp = case["excel_results"]["baseline"]
    assert math.isclose(res["baseline"]["applied_water"], b_exp["applied_water"], abs_tol=1e-9)
    assert math.isclose(res["baseline"]["energy_cost"], b_exp["energy_cost"], abs_tol=0.01)
    for name in UPGRADES:
        exp, got = case["excel_results"][name], rows[name]
        for f in ("eff", "upgrade_cost", "om", "life", "applied_water", "water_saved_pct"):
            assert math.isclose(got[f], exp[f], abs_tol=1e-9), f"{name} {f}"
        for f in ("energy_cost", "energy_savings", "revenue_benefit",
                  "reinvest_minus_salvage_pv"):
            assert math.isclose(got[f], exp[f], abs_tol=0.01), f"{name} {f}"
        assert math.isclose(got["water_cap"], exp["water_cap"], abs_tol=1e-9)
        assert got["eligible"] == exp["eligible"]
        assert got["meets_water"] == exp["meets_water"]
        assert got["econ_status"] == exp["econ_status"]
        if isinstance(exp["payback"], str):
            assert got["payback"] == exp["payback"]
        else:
            assert math.isclose(got["payback"], exp["payback"], abs_tol=1e-6)


@pytest.mark.parametrize("case", CASES, ids=ids)
def test_pure_function(case):
    inputs = to_form(case["inputs"])
    inputs_before = copy.deepcopy(inputs)
    data_before = copy.deepcopy(DATA)
    first = calculate_wise(inputs, DATA)
    second = calculate_wise(inputs, DATA)
    assert first == second
    assert inputs == inputs_before, "calculate_wise mutated its inputs"
    assert DATA == data_before, "calculate_wise mutated the data"
