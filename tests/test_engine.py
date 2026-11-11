"""Behavioural tests for workbook rules the golden cases don't exercise directly."""
import copy
import json
import math
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from app import calculate_wise, load_data, resolve_inputs  # noqa: E402

DATA = load_data(os.path.join(ROOT, "data", "app_data.json"))
with open(os.path.join(ROOT, "tests", "golden_cases.json")) as f:
    BASE = json.load(f)[0]["inputs"]   # NE flood electric


def case(**overrides):
    i = copy.deepcopy(BASE)
    i.update(overrides)
    return i


def rows(result):
    return {r["system"]: r for r in result["systems"]}


def test_blank_form_is_not_ready_and_lists_every_required_field():
    res = calculate_wise({}, DATA)
    assert not res["ready"]
    assert res["recommended_system"] == "Complete required inputs"
    assert res["priority"] == "Pending"
    labels = [m["label"] for m in res["missing"]]
    assert labels[:3] == ["State", "County", "Field area"]
    assert "Pumping capacity" not in labels          # optional (INPUTS!AE56)
    assert res["readiness"] == f"⚠ {len(labels)} REQUIRED FIELD(S) MISSING"


def test_county_must_belong_to_state():
    res = resolve_inputs(case(county="Finney"), DATA)   # Finney is in Kansas
    assert {"field": "county", "label": "County"} in res["missing"]


def test_custom_upgrade_cost_applies_only_to_selected_upgrade():
    res = calculate_wise(case(upgrade_scenario="Drip", custom_upgrade_cost=700), DATA)
    r = rows(res)
    assert r["Drip"]["upgrade_cost"] == 700
    assert r["Center Pivot"]["upgrade_cost"] == 600
    assert r["Drip"]["investment"] == 700 * BASE["acres"]


def test_current_drip_makes_every_upgrade_ineligible():
    res = calculate_wise(case(current_system="Drip"), DATA)
    assert all(r["decision"] == "Not eligible" for r in res["systems"])
    assert res["recommended_system"] == "No upgrade recommended"


def test_natural_gas_uses_brake_kwh_per_mcf():
    res = resolve_inputs(case(energy_source="Natural Gas"), DATA)
    c = DATA["constants"]
    assert math.isclose(res["auto"]["cost_per_brake_kwh"],
                        8.45 / (c["ng_kwh_per_mcf"] * c["gas_engine_efficiency"]))


def test_custom_reduction_scenario():
    missing = resolve_inputs(case(water_scenario="Custom"), DATA)["missing"]
    assert {"field": "custom_reduction", "label": "Custom reduction"} in missing
    res = calculate_wise(case(water_scenario="Custom", custom_reduction=0.5), DATA)
    assert res["ready"] and math.isclose(rows(res)["Drip"]["water_cap"], 10.0)


def test_no_irrigation_requires_rainfed_yield_and_uses_planned_water_label():
    res = resolve_inputs(case(current_system="No Irrigation System", irrigated="No",
                              rainfed_yield=None, gross_irrigation_in=None), DATA)
    labels = [m["label"] for m in res["missing"]]
    assert "Rainfed yield" in labels
    assert "Planned irrigation water available" in labels


def test_other_crop_needs_custom_yield_and_price():
    res = resolve_inputs(case(crop="Other (manual entry)"), DATA)
    labels = [m["label"] for m in res["missing"]]
    assert "Expected yield" in labels and "Market price" in labels
    assert res["messages"]["yield"].startswith("No county or state yield data")
    ok = calculate_wise(case(crop="Other (manual entry)", custom_yield=100, custom_price=3), DATA)
    assert ok["ready"]


def test_equal_discount_and_escalation_rates_and_zero_discount():
    for r, g in ((0.03, 0.03), (0.0, 0.02), (0.0, 0.0)):
        res = calculate_wise(case(discount_rate=r, energy_escalation=g), DATA)
        assert all(math.isfinite(x["npv"]) for x in res["systems"])


def test_gross_net_warning_threshold():
    # 20 in x 0.60 = 12.0 vs net 12 -> 0% -> no warning; 24 x 0.6 = 14.4 -> +20% exactly -> none
    assert resolve_inputs(case(), DATA)["messages"]["gross_net_warning"] is None
    assert resolve_inputs(case(gross_irrigation_in=24), DATA)["messages"]["gross_net_warning"] is None
    w = resolve_inputs(case(gross_irrigation_in=25), DATA)["messages"]["gross_net_warning"]
    assert w and "by +25%" in w


def test_string_inputs_from_form_are_coerced():
    as_strings = {k: (str(v) if isinstance(v, (int, float)) else v) for k, v in BASE.items()}
    a, b = calculate_wise(as_strings, DATA), calculate_wise(BASE, DATA)
    assert [r["npv"] for r in a["systems"]] == [r["npv"] for r in b["systems"]]
