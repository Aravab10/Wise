"""
WISE — Water & Irrigation Systems Evaluation
Flask Backend — Version 2 (All 50 States + Real Climate Data)

What's in here:
- All 51 states, 3094 counties with real climate data
- Real EIA March 2026 electricity prices per state
- Real USDA NASS crop prices per state
- Irrigation Source (Groundwater vs Surface Water) — affects TDH
- Base Production Cost from USDA ERS — shown for context
- Formulas translated directly from Karthik's Results_calc sheet
"""

from flask import Flask, render_template, request, jsonify
import json, os

app = Flask(__name__)

# ─────────────────────────────────────────────────────────────
# LOAD ALL DATA AT STARTUP
# Extracted from Karthik's Excel file into one JSON.
# ─────────────────────────────────────────────────────────────

DATA_PATH = os.path.join(os.path.dirname(__file__), 'data', 'app_data.json')
with open(DATA_PATH) as f:
    APP_DATA = json.load(f)

STATE_COUNTIES    = APP_DATA['state_counties']       # {state: [county, ...]}
ELEC_PRICES       = APP_DATA['elec_prices']          # {state: $/kWh}
STATE_CROP_PRICES = APP_DATA['state_crop_prices']    # {state: {crop: price}}
NATIONAL_PRICES   = APP_DATA['national_prices']      # {crop: price} fallback
CLIMATE           = APP_DATA.get('climate', {})      # {state|county: {et,precip,temp,aws}}
SLOPE             = APP_DATA.get('slope', {})        # {state|county: slope_pct}
COUNTY_YIELDS     = APP_DATA.get('county_yields', {})# {state|county: {crop: yield}}


# ─────────────────────────────────────────────────────────────
# SYSTEM SPECS — from SystemTech sheet in Karthik's Excel
# efficiency    = % of water that reaches the crop
# upgrade_cost  = $/acre to install
# annual_om     = $/acre/year extra maintenance vs current
# ─────────────────────────────────────────────────────────────

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

# Water constraint reduction percentages — from Lists sheet
WATER_CONSTRAINTS = {
    "No Constraint":       0.00,
    "Mild Constraint":     0.10,
    "Moderate Constraint": 0.25,
    "Severe Constraint":   0.40,
}

# Base production costs — from CropCost sheet (USDA ERS 2025)
# These are shown for context but do NOT affect NPV/payback ranking
CROP_COSTS = {
    "Corn":     620.0,
    "Soybeans": 410.0,
    "Sorghum":  340.0,
    "Wheat":    290.0,
    "Cotton":   850.0,
    "Potatoes": 2500.0,
}


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def get_crop_price(state, crop):
    """State price first (USDA NASS 5-yr avg), national fallback."""
    return STATE_CROP_PRICES.get(state, {}).get(crop) or NATIONAL_PRICES.get(crop, 6.0)

def get_climate(state, county):
    """Real ET, precip, temp, AWS from Excel Climate_data sheet."""
    return CLIMATE.get(f"{state}|{county}")

def get_county_yield(state, county, crop):
    """County-level irrigated yield if available."""
    return COUNTY_YIELDS.get(f"{state}|{county}", {}).get(crop)


# ─────────────────────────────────────────────────────────────
# CORE CALCULATION ENGINE
# Translated directly from Results_calc sheet in Karthik's Excel.
# Formula references included as comments.
# ─────────────────────────────────────────────────────────────

def calculate_wise(inputs):
    """
    Run WISE analysis for all 3 upgrade systems.
    Returns (ranked_results, baseline_info).

    Formula sources (Results_calc sheet):
    E col: Applied Water  = INPUTS!E53 / efficiency
    H col: Energy Cost    = E_water * area * TDH * 0.0853 / pump_eff * elec_rate
    I col: Energy Savings = H_baseline - H_new
    J col: Annual Net Savings = I - (D * area)
    K col: Total Investment = C * area * (1 - incentive)
    L col: NPV = growing annuity formula with energy escalation
    M col: Payback = K / J
    N col: Rank = RANK.EQ by NPV descending
    """

    # ── Unpack inputs ─────────────────────────────────────────
    state          = inputs['state']
    county         = inputs['county']
    current_sys    = inputs['current_system']
    crop           = inputs['crop']
    acres          = float(inputs['acres'])
    yield_ac       = float(inputs['yield_per_acre'])
    water_scenario = inputs['water_scenario']
    irrigation_src = inputs.get('irrigation_source', 'Groundwater')
    well_depth     = float(inputs.get('well_depth', 200))
    op_pressure    = float(inputs.get('op_pressure', 40))
    pump_eff       = float(inputs.get('pump_efficiency', 0.75))
    incentive_pct  = float(inputs.get('incentive_pct', 0))
    horizon        = int(inputs.get('horizon', 25))
    discount_rate  = float(inputs.get('discount_rate', 0.04))
    energy_escal   = float(inputs.get('energy_escalation', 0.04))
    net_crop_water = float(inputs.get('net_crop_water', 18))  # INPUTS!E53 = 18 (USGS)

    # Surface water = no well, zero the well depth component
    if irrigation_src == 'Surface Water':
        well_depth = 0

    # ── Get real data from memory ─────────────────────────────
    elec_rate   = ELEC_PRICES.get(state, 0.12)
    # Use override if farmer provided one, otherwise use USDA NASS state avg
    crop_price  = float(inputs.get('crop_price_override') or get_crop_price(state, crop))
    # Production cost — shown for context, does NOT affect NPV ranking
    crop_cost   = float(inputs.get('prod_cost_override') or CROP_COSTS.get(crop, 500))
    current_eff = CURRENT_SYSTEM_EFFICIENCY.get(current_sys, 0.75)
    water_red   = WATER_CONSTRAINTS.get(water_scenario, 0.0)

    # ── Real climate data for this county ─────────────────────
    climate_data = get_climate(state, county)
    et     = climate_data['et']     if climate_data else None
    precip = climate_data['precip'] if climate_data else None
    temp   = climate_data['temp']   if climate_data else None
    aws    = climate_data['aws']    if climate_data else None

    # ── Total Dynamic Head — Excel INPUTS: =E24+(E25*2.31) ───
    # Well depth (ft) + operating pressure converted to feet of head
    # Surface water farms have well_depth=0 so TDH is just pressure head
    tdh = well_depth + (op_pressure * 2.31)

    # ── Baseline: Keep Current System ─────────────────────────
    # Applied water = net crop water requirement / efficiency
    # Excel Results_calc E2: =IFERROR(INPUTS!E53/B2,"")
    baseline_water = net_crop_water / current_eff

    # Energy cost for baseline system
    # Excel Results_calc H2: =E2*INPUTS!E17*INPUTS!E27*0.0853/INPUTS!E26*INPUTS!E22
    baseline_energy = baseline_water * acres * tdh * 0.0853 / pump_eff * elec_rate

    results = []

    for sys_name, specs in UPGRADE_SYSTEMS.items():

        # 1. Applied Water — Excel: =IFERROR(INPUTS!E53/efficiency,"")
        applied_water = net_crop_water / specs['efficiency']

        # 2. Water Savings — Excel: =E2-E3
        water_saved     = baseline_water - applied_water
        water_saved_pct = water_saved / baseline_water if baseline_water > 0 else 0

        # 3. Energy Cost for upgraded system — same formula, new applied_water
        new_energy = applied_water * acres * tdh * 0.0853 / pump_eff * elec_rate

        # 4. Energy Savings — Excel: =H2-H3
        energy_savings = baseline_energy - new_energy

        # 5. Annual Net Savings — Excel: =I3-(D3*INPUTS!E17)
        annual_om_total    = specs['annual_om'] * acres
        annual_net_savings = energy_savings - annual_om_total

        # 6. Total Investment after incentive — Excel: =C3*INPUTS!E17*(1-INPUTS!E47)
        gross_investment = specs['upgrade_cost'] * acres
        net_investment   = gross_investment * (1 - incentive_pct)

        # 7. NPV — growing annuity formula — Excel L col
        # Accounts for energy prices escalating each year
        # IF(r==g): savings * n / (1+r) - investment
        # ELSE:     savings * (1-((1+g)/(1+r))^n) / (r-g) - investment
        r, g, n, s = discount_rate, energy_escal, horizon, annual_net_savings
        if abs(r - g) < 0.0001:
            pv_savings = s * n / (1 + r)
        else:
            pv_savings = s * (1 - ((1 + g) / (1 + r)) ** n) / (r - g)
        npv = pv_savings - net_investment

        # 8. Payback — Excel: =IF(J<=0,"Not recovered",IF(K/J>horizon,"Not recovered",K/J))
        if annual_net_savings <= 0:
            payback, payback_label = None, "Not recovered"
        elif net_investment / annual_net_savings > horizon:
            payback, payback_label = None, "Not recovered"
        else:
            payback = round(net_investment / annual_net_savings, 1)
            payback_label = f"{payback} yrs"

        # Context: gross and net revenue per acre (display only, not in ranking)
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

    # ── Ranking — Excel: =IF(NPV<=0,"Not recommended",RANK.EQ(...)) ──
    pos    = sorted([r for r in results if r['npv'] > 0],  key=lambda x: x['npv'], reverse=True)
    neg    = sorted([r for r in results if r['npv'] <= 0], key=lambda x: x['npv'], reverse=True)
    ranked = pos + neg

    for i, r in enumerate(ranked):
        r['rank']        = i + 1
        r['recommended'] = (i == 0 and r['npv'] > 0)
        r['rank_label']  = "Not recommended" if r['npv'] <= 0 else str(i + 1)

    # ── Baseline info for display ──────────────────────────────
    baseline = {
        "system":           current_sys,
        "irrigation_src":   irrigation_src,
        "applied_water":    round(baseline_water, 1),
        "energy_cost":      round(baseline_energy),
        "elec_rate":        elec_rate,
        "crop_price":       crop_price,
        "crop_cost":        crop_cost,
        "tdh":              round(tdh, 1),
        "water_reduction":  round(water_red * 100),
        "net_crop_water":   net_crop_water,
        "et":               et,
        "precip":           precip,
        "temp":             temp,
        "aws":              aws,
        "climate_source":   "Excel Climate_data (PRISM/NOAA)" if climate_data else "Default estimate",
    }

    return ranked, baseline


def build_rationale(rec, inputs, baseline):
    county   = inputs['county']
    state    = inputs['state']
    crop     = inputs['crop']
    scenario = inputs['water_scenario']

    if not rec or rec['npv'] <= 0:
        return (
            "No upgrade system produces a positive NPV under these assumptions. "
            "Keeping the current system is recommended unless water conservation "
            "is the primary goal rather than economic return."
        )
    return (
        f"{rec['system']} provides the highest positive NPV for {crop} production "
        f"in {county} County, {state}. "
        f"It reduces applied water use by {rec['water_saved_pct']}% "
        f"and generates estimated annual net savings of ${rec['annual_net_savings']:,}/year "
        f"with a payback period of {rec['payback_label']}. "
        f"Irrigation source: {baseline['irrigation_src']} · "
        f"Elec: ${baseline['elec_rate']:.4f}/kWh (EIA 2026) · "
        f"Crop price: ${baseline['crop_price']}/bu (USDA NASS) · "
        f"Scenario: {scenario}."
    )


# ─────────────────────────────────────────────────────────────
# FLASK ROUTES
# ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    states = sorted(STATE_COUNTIES.keys())
    return render_template("index.html", states=states)


@app.route("/get_counties")
def get_counties():
    state    = request.args.get("state", "")
    counties = sorted(STATE_COUNTIES.get(state, []))
    elec     = ELEC_PRICES.get(state)
    return jsonify({"counties": counties, "elec_rate": elec})


@app.route("/get_county_data")
def get_county_data():
    state  = request.args.get("state", "")
    county = request.args.get("county", "")
    crop   = request.args.get("crop", "")

    elec_rate    = ELEC_PRICES.get(state)
    crop_price   = get_crop_price(state, crop) if crop else None
    prod_cost    = CROP_COSTS.get(crop) if crop else None
    climate_data = get_climate(state, county)
    slope_val    = SLOPE.get(f"{state}|{county}")
    county_yield = get_county_yield(state, county, crop) if crop else None

    price_source = (
        f"USDA NASS 5-yr avg · {state}"
        if state in STATE_CROP_PRICES and crop in STATE_CROP_PRICES.get(state, {})
        else "USDA NASS national average"
    )

    return jsonify({
        "elec_rate":    elec_rate,
        "crop_price":   crop_price,
        "prod_cost":    prod_cost,
        "county_yield": county_yield,
        "climate":      {
            "et":     climate_data['et'],
            "precip": climate_data['precip'],
            "temp":   climate_data['temp'],
            "aws":    climate_data['aws'],
        } if climate_data else None,
        "slope":        slope_val,
        "data_sources": {
            "elec":      f"EIA March 2026 · {state}",
            "price":     price_source,
            "prod_cost": f"USDA ERS 2025 · {crop}" if crop else "Select crop first",
            "climate":   "Excel Climate_data sheet" if climate_data else "Not available",
            "yield":     "County irrigated yield" if county_yield else "Enter manually",
        }
    })


@app.route("/calculate", methods=["POST"])
def calculate():
    inputs = request.get_json()

    required = ["state", "county", "current_system", "crop",
                "acres", "yield_per_acre", "water_scenario"]
    for f in required:
        if not inputs.get(f):
            return jsonify({"error": f"Missing required field: {f}"}), 400

    try:
        results, baseline = calculate_wise(inputs)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    rec       = next((r for r in results if r.get('recommended')), results[0])
    rationale = build_rationale(rec, inputs, baseline)

    horizon = int(inputs.get('horizon', 25))
    if rec['npv'] <= 0:
        priority = "Low"
    elif rec['payback'] and rec['payback'] <= horizon / 2:
        priority = "High"
    else:
        priority = "Medium"

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
