"""
WISE — Water & Irrigation Systems Evaluation
Flask Backend — Version 3
Updates per Karthik Ramaswamy email (Jul 15 2026):
  1. Irrigation Source: Groundwater + Surface Water
  2. Energy Source: Electricity or Natural Gas (with state prices)
  3. County-level crop yield from Yield_county sheet
  4. Custom yield input with fallback message if no county data
"""

from flask import Flask, render_template, request, jsonify
import json, os

app = Flask(__name__)

# ── Load all data at startup ──────────────────────────────
DATA_PATH = os.path.join(os.path.dirname(__file__), 'data', 'app_data.json')
with open(DATA_PATH) as f:
    APP_DATA = json.load(f)

STATE_COUNTIES    = APP_DATA['state_counties']
ELEC_PRICES       = APP_DATA['elec_prices']       # $/kWh — EIA March 2026
NG_PRICES         = APP_DATA.get('ng_prices', {}) # $/mcf — NGP sheet latest year
STATE_CROP_PRICES = APP_DATA['state_crop_prices']
NATIONAL_PRICES   = APP_DATA['national_prices']
CLIMATE           = APP_DATA.get('climate', {})
SLOPE             = APP_DATA.get('slope', {})
COUNTY_YIELDS     = APP_DATA.get('county_yields', {})  # {state|county: {crop: {yield,unit,status}}}

# ── System specs from SystemTech sheet ───────────────────
UPGRADE_SYSTEMS = {
    "Center Pivot":       {"efficiency": 0.75, "upgrade_cost": 350,  "annual_om": 10},
    "Improved Sprinkler": {"efficiency": 0.85, "upgrade_cost": 1000, "annual_om": 25},
    "Drip":               {"efficiency": 0.95, "upgrade_cost": 1500, "annual_om": 40},
}

CURRENT_SYSTEM_EFFICIENCY = {
    "Flood":     0.60,
    "Sprinkler": 0.85,
    "Drip":      0.95,
}

WATER_CONSTRAINTS = {
    "No Constraint":       0.00,
    "Mild Constraint":     0.10,
    "Moderate Constraint": 0.25,
    "Severe Constraint":   0.40,
}

CROP_COSTS = {
    "Corn":     620.0,
    "Soybeans": 410.0,
    "Sorghum":  340.0,
    "Wheat":    290.0,
    "Cotton":   850.0,
    "Potatoes": 2500.0,
}

# Natural gas conversion: 1 mcf ≈ 293 kWh (thermal equivalent)
# Used to normalize NG price to $/kWh equivalent for energy cost formula
MCF_TO_KWH = 293.0


# ── Helpers ───────────────────────────────────────────────

def get_crop_price(state, crop):
    return STATE_CROP_PRICES.get(state, {}).get(crop) or NATIONAL_PRICES.get(crop, 6.0)

def get_climate(state, county):
    return CLIMATE.get(f"{state}|{county}")

def get_county_yield(state, county, crop):
    """
    Returns dict with yield, unit, status, and available flag.
    If not available, returns None so frontend can show message.
    """
    data = COUNTY_YIELDS.get(f"{state}|{county}", {}).get(crop)
    return data  # {yield, unit, status} or None

def get_energy_rate(state, energy_source):
    """
    Returns energy price in $/kWh equivalent.
    Electricity: direct $/kWh from EIA
    Natural Gas: convert from $/mcf to $/kWh equivalent
    """
    if energy_source == "Natural Gas":
        ng_price = NG_PRICES.get(state)  # $/mcf
        if ng_price:
            return round(ng_price / MCF_TO_KWH, 6), f"${ng_price}/mcf → ${round(ng_price/MCF_TO_KWH,5)}/kWh equiv."
        return 0.04, "Natural gas price not available for this state"
    else:
        elec = ELEC_PRICES.get(state, 0.12)
        return elec, f"EIA March 2026 · {state}"


# ── Core Calculation Engine ───────────────────────────────

def calculate_wise(inputs):
    """
    Formulas from Results_calc sheet.
    Updated per Karthik email Jul 15 2026:
    - irrigation_source affects well depth in TDH
    - energy_source selects electricity or natural gas price
    - yield from county data or custom entry
    """

    state          = inputs['state']
    county         = inputs['county']
    current_sys    = inputs['current_system']
    crop           = inputs['crop']
    acres          = float(inputs['acres'])
    water_scenario = inputs['water_scenario']
    irrigation_src = inputs.get('irrigation_source', 'Groundwater')
    gw_blend       = float(inputs.get('gw_blend', 1.0))  # fraction from groundwater
    energy_src     = inputs.get('energy_source', 'Electricity')
    well_depth     = float(inputs.get('well_depth', 200))
    op_pressure    = float(inputs.get('op_pressure', 40))
    pump_eff       = float(inputs.get('pump_efficiency', 0.75))
    incentive_pct  = float(inputs.get('incentive_pct', 0))
    horizon        = int(inputs.get('horizon', 25))
    discount_rate  = float(inputs.get('discount_rate', 0.04))
    energy_escal   = float(inputs.get('energy_escalation', 0.04))
    net_crop_water = float(inputs.get('net_crop_water', 18))

    # Yield — use county data if available, else use custom entry
    yield_ac = float(inputs.get('yield_per_acre') or 0)

    # Effective well depth depends on water source:
    # - Surface Water: no well, depth = 0
    # - Groundwater: full well depth
    # - Both: scale well depth by the groundwater fraction, since only
    #   the groundwater portion needs deep pumping. Surface water is
    #   pumped from ground level so contributes no well-lift energy.
    if irrigation_src == 'Surface Water':
        well_depth = 0
    elif irrigation_src == 'Groundwater + Surface Water':
        well_depth = well_depth * gw_blend

    # Energy rate — electricity or natural gas
    energy_rate, energy_rate_label = get_energy_rate(state, energy_src)

    # Other real data
    crop_price  = float(inputs.get('crop_price_override') or get_crop_price(state, crop))
    crop_cost   = float(inputs.get('prod_cost_override') or CROP_COSTS.get(crop, 500))
    current_eff = CURRENT_SYSTEM_EFFICIENCY.get(current_sys, 0.75)
    water_red   = WATER_CONSTRAINTS.get(water_scenario, 0.0)

    # Climate data
    climate_data = get_climate(state, county)
    et     = climate_data['et']     if climate_data else None
    precip = climate_data['precip'] if climate_data else None
    temp   = climate_data['temp']   if climate_data else None
    aws    = climate_data['aws']    if climate_data else None

    # TDH — Excel: =E24+(E25*2.31)
    tdh = well_depth + (op_pressure * 2.31)

    # Baseline energy cost — Excel Results_calc H2
    baseline_water  = net_crop_water / current_eff
    baseline_energy = baseline_water * acres * tdh * 0.0853 / pump_eff * energy_rate

    results = []
    for sys_name, specs in UPGRADE_SYSTEMS.items():

        # Applied water — Excel: =INPUTS!E53/efficiency
        applied_water = net_crop_water / specs['efficiency']

        # Water savings
        water_saved     = baseline_water - applied_water
        water_saved_pct = water_saved / baseline_water if baseline_water > 0 else 0

        # New energy cost
        new_energy = applied_water * acres * tdh * 0.0853 / pump_eff * energy_rate

        # Energy savings
        energy_savings = baseline_energy - new_energy

        # Annual net savings — Excel: =energy_savings-(om*area)
        annual_om_total    = specs['annual_om'] * acres
        annual_net_savings = energy_savings - annual_om_total

        # Investment — Excel: =cost*area*(1-incentive)
        gross_investment = specs['upgrade_cost'] * acres
        net_investment   = gross_investment * (1 - incentive_pct)

        # NPV — growing annuity — Excel L col
        r, g, n, s = discount_rate, energy_escal, horizon, annual_net_savings
        if abs(r - g) < 0.0001:
            pv_savings = s * n / (1 + r)
        else:
            pv_savings = s * (1 - ((1 + g) / (1 + r)) ** n) / (r - g)
        npv = pv_savings - net_investment

        # Payback — Excel M col
        if annual_net_savings <= 0:
            payback, payback_label = None, "Not recovered"
        elif net_investment / annual_net_savings > horizon:
            payback, payback_label = None, "Not recovered"
        else:
            payback = round(net_investment / annual_net_savings, 1)
            payback_label = f"{payback} yrs"

        gross_revenue = yield_ac * crop_price
        net_revenue   = gross_revenue - crop_cost

        results.append({
            "system":             sys_name,
            "efficiency_pct":     int(specs['efficiency'] * 100),
            "applied_water":      round(applied_water, 1),
            "water_saved":        round(water_saved, 1),
            "water_saved_pct":    round(water_saved_pct * 100, 1),
            "energy_cost":        round(new_energy),
            "energy_savings":     round(energy_savings),
            "annual_om":          round(annual_om_total),
            "annual_net_savings": round(annual_net_savings),
            "gross_investment":   round(gross_investment),
            "net_investment":     round(net_investment),
            "npv":                round(npv),
            "payback":            payback,
            "payback_label":      payback_label,
            "gross_revenue":      round(gross_revenue),
            "net_revenue":        round(net_revenue),
        })

    # Rank by NPV
    pos    = sorted([r for r in results if r['npv'] > 0],  key=lambda x: x['npv'], reverse=True)
    neg    = sorted([r for r in results if r['npv'] <= 0], key=lambda x: x['npv'], reverse=True)
    ranked = pos + neg
    for i, r in enumerate(ranked):
        r['rank']        = i + 1
        r['recommended'] = (i == 0 and r['npv'] > 0)
        r['rank_label']  = "Not recommended" if r['npv'] <= 0 else str(i + 1)

    baseline = {
        "system":           current_sys,
        "irrigation_src":   irrigation_src,
        "gw_blend_pct":     round(gw_blend * 100) if irrigation_src == 'Groundwater + Surface Water' else None,
        "energy_src":       energy_src,
        "energy_rate":      energy_rate,
        "energy_rate_label":energy_rate_label,
        "applied_water":    round(baseline_water, 1),
        "energy_cost":      round(baseline_energy),
        "crop_price":       crop_price,
        "crop_cost":        crop_cost,
        "tdh":              round(tdh, 1),
        "water_reduction":  round(water_red * 100),
        "net_crop_water":   net_crop_water,
        "et": et, "precip": precip, "temp": temp, "aws": aws,
        "climate_source":   "Excel Climate_data (PRISM/NOAA)" if climate_data else "Default estimate",
    }

    return ranked, baseline


def build_rationale(rec, inputs, baseline):
    if not rec or rec['npv'] <= 0:
        return ("No upgrade system produces a positive NPV under these assumptions. "
                "Keeping the current system is recommended unless water conservation "
                "is the primary goal.")
    return (
        f"{rec['system']} provides the highest positive NPV for {inputs['crop']} "
        f"in {inputs['county']} County, {inputs['state']}. "
        f"Reduces applied water by {rec['water_saved_pct']}%, "
        f"annual net savings ${rec['annual_net_savings']:,}/yr, "
        f"payback {rec['payback_label']}. "
        f"Energy source: {baseline['energy_src']} at "
        f"${baseline['energy_rate']:.4f}/kWh equiv. · "
        f"Irrigation: {baseline['irrigation_src']} · "
        f"Scenario: {inputs['water_scenario']}."
    )


# ── Flask Routes ──────────────────────────────────────────

@app.route("/")
def index():
    states = sorted(STATE_COUNTIES.keys())
    return render_template("index.html", states=states)


@app.route("/get_counties")
def get_counties():
    state    = request.args.get("state", "")
    counties = sorted(STATE_COUNTIES.get(state, []))
    elec     = ELEC_PRICES.get(state)
    ng       = NG_PRICES.get(state)
    return jsonify({"counties": counties, "elec_rate": elec, "ng_price": ng})


@app.route("/get_county_data")
def get_county_data():
    state  = request.args.get("state", "")
    county = request.args.get("county", "")
    crop   = request.args.get("crop", "")
    energy = request.args.get("energy_source", "Electricity")

    elec_rate    = ELEC_PRICES.get(state)
    ng_price     = NG_PRICES.get(state)
    crop_price   = get_crop_price(state, crop) if crop else None
    prod_cost    = CROP_COSTS.get(crop) if crop else None
    climate_data = get_climate(state, county)
    slope_val    = SLOPE.get(f"{state}|{county}")

    # County yield — returns data or None
    county_yield_data = get_county_yield(state, county, crop) if crop else None

    # Build yield response
    if county_yield_data:
        yield_info = {
            "available":   True,
            "value":       county_yield_data['yield'],
            "unit":        county_yield_data['unit'],
            "status":      county_yield_data['status'],
            "message":     None
        }
    else:
        yield_info = {
            "available":   False,
            "value":       None,
            "unit":        None,
            "status":      None,
            "message":     "NASS County-level yield data are unavailable. Please enter a custom yield value based on your best knowledge or local experience."
        }

    price_source = (
        f"USDA NASS 5-yr avg · {state}"
        if state in STATE_CROP_PRICES and crop in STATE_CROP_PRICES.get(state, {})
        else "USDA NASS national average"
    )

    return jsonify({
        "elec_rate":    elec_rate,
        "ng_price":     ng_price,
        "crop_price":   crop_price,
        "prod_cost":    prod_cost,
        "yield_info":   yield_info,
        "climate": {
            "et":     climate_data['et'],
            "precip": climate_data['precip'],
            "temp":   climate_data['temp'],
            "aws":    climate_data['aws'],
        } if climate_data else None,
        "slope": slope_val,
        "data_sources": {
            "elec":      f"EIA March 2026 · {state}",
            "ng":        f"EIA Natural Gas Prices 2025 · {state}" if ng_price else "Not available",
            "price":     price_source,
            "prod_cost": f"USDA ERS 2025 · {crop}" if crop else "Select crop first",
            "climate":   "Excel Climate_data sheet" if climate_data else "Not available",
        }
    })


@app.route("/calculate", methods=["POST"])
def calculate():
    inputs = request.get_json()

    required = ["state", "county", "current_system", "crop",
                "acres", "water_scenario"]
    for f in required:
        if not inputs.get(f):
            return jsonify({"error": f"Missing required field: {f}"}), 400

    # yield_per_acre must be present (either from county data or custom)
    if not inputs.get('yield_per_acre') or float(inputs.get('yield_per_acre', 0)) <= 0:
        return jsonify({"error": "Please enter a yield value (bu/ac)"}), 400

    try:
        results, baseline = calculate_wise(inputs)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    rec       = next((r for r in results if r.get('recommended')), results[0])
    rationale = build_rationale(rec, inputs, baseline)

    horizon = int(inputs.get('horizon', 25))
    if rec['npv'] <= 0:       priority = "Low"
    elif rec['payback'] and rec['payback'] <= horizon / 2: priority = "High"
    else:                     priority = "Medium"

    return jsonify({
        "recommended_system": rec['system'] if rec['recommended'] else "Keep Current System",
        "priority":           priority,
        "rationale":          rationale,
        "baseline":           baseline,
        "results":            results,
        "key_metrics": {
            "annual_net_savings": rec['annual_net_savings'],
            "water_saved_pct":    rec['water_saved_pct'],
            "payback_label":      rec['payback_label'],
            "npv":                rec['npv'],
        }
    })


if __name__ == "__main__":
    app.run(debug=True, port=5000)
