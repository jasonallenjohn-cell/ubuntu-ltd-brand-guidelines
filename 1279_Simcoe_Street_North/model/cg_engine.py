#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1279 Simcoe Street North, Oshawa — Common Ground Initiative monthly engine.

Implements the September 2026 financial-modelling handoff under the Ubuntu Land
Trust Common Ground structure (July 2026 capital stack):

  CMHC MLI Select 50% of TDC (50-yr amortization, construction-to-perm)
  Build Canada Homes 40% of TDC, split ~54.5% conditionally-repayable
      contribution (vests Yrs 1-10 vs deep affordability) + ~45.5% mezzanine
      (BoC + 100 bps, no interest or payments until Year 10, swept Yrs 11-20)
  City 5% of TDC
  Landowner retained land equity 5% of TDC (in-kind, converts to Year-10 buyout claim)

  Land is NOT purchased. The landowner receives a Day-1 draw = 60% of the
  appraisal at each phase's financial close (pro rata by homes, editable), keeps
  5% of TDC as retained equity, and donates the remainder of the appraisal to
  affordability.  Net cash flow after MLI debt service is split 60% landowner /
  40% non-profit for each phase's Years 1-10; at Year 10 the landowner is bought
  out (retained equity + MLI principal paid down) via a 50-yr supplemental charge
  and exits; BCH takes the 60% seat Years 11-20 to sweep the mezz.

Every number in the workbook is either an INPUT on the Inputs tab (with source,
date, units, confidence, owner) or a FORMULA.  This engine reproduces the same
arithmetic so that scenarios, residual solves and density cases (which need
goal-seek / data-table behaviour openpyxl cannot author) can be computed and
written as clearly-labelled engine values.  The base case is reconciled to the
live workbook formulas by LibreOffice recalculation (see build_workbook.py).
"""
from __future__ import annotations
import copy, math, json
from dataclasses import dataclass, field

M2_TO_SF = 10.7639104

# --------------------------------------------------------------------------
# INPUT REGISTER
# Each input: key -> dict(value, unit, label, source, date, conf, owner, flag)
#   conf: H / M / L ;  flag: "" | "PLACEHOLDER" | "UNVERIFIED" | "PENDING"
# --------------------------------------------------------------------------
def I(value, unit, label, source, date="2026-09-15", conf="M", owner="Analyst", flag="", group="General"):
    return dict(value=value, unit=unit, label=label, source=source, date=date,
                conf=conf, owner=owner, flag=flag, group=group)

def default_inputs():
    d = {}
    g = "Site & acquisition"
    d["site_net_m2"]      = I(31211, "m²", "Net site (R01 Block 1)", "Draft plan S-O-2022-05, Attachment 2 to ED-24-20 (3.1211 ha)", "2024-01", "H", "Survey", "", g)
    d["site_gross_ac"]    = I(7.85, "acres", "Gross site per acquisition summary", "Supplied acquisition summary; reconcile to survey", "2026-09", "L", "Vendor", "UNVERIFIED", g)
    d["land_appraisal"]   = I(17_000_000, "$", "Land value used for Common Ground draw (seller ask as placeholder)", "Seller ask $17.0M; NOT an appraisal. Common Ground requires an AACI appraisal", "2026-09", "L", "Vendor", "PLACEHOLDER", g)
    d["draw_pct"]         = I(0.60, "%", "Day-1 landowner draw as % of appraisal", "Common Ground July 2026 structure (locked)", "2026-07", "H", "Ubuntu", "", g)
    d["land_eq_pct"]      = I(0.05, "% of TDC", "Landowner retained equity", "Common Ground July 2026 structure (locked)", "2026-07", "H", "Ubuntu", "", g)
    d["draw_timing"]      = I(1, "1 = pro rata by phase homes; 0 = 100% at Phase 1 close", "Land draw timing", "Analyst assumption; editable", "2026-09", "M", "Analyst", "", g)
    d["existing_noi"]     = I(500_000, "$/yr", "Existing NOI (seller-supplied)", "Vendor summary; rent roll, leases and expenses NOT provided", "2026-09", "L", "Vendor", "UNVERIFIED", g)
    d["existing_noi_keep"]= I(0.0, "share", "Share of existing NOI retained during entitlement (0 = zero-retained-income case)", "Handoff §2: use zero-retained case unless leases support retention", "2026-09", "M", "Analyst", "", g)
    d["legal_title"]      = I(150_000, "$", "Legal, title, lease registration, due diligence", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["ltt_applies"]      = I(0, "1/0", "Land transfers to SPV (Ontario LTT applies)?", "Common Ground = land lease, no transfer; set 1 if structure changes", "2026-09", "M", "Ubuntu", "", g)
    d["carry_taxes"]      = I(120_000, "$/yr", "Property tax + insurance on existing site during entitlement", "Analyst placeholder pending tax bill", "2026-09", "L", "Analyst", "PLACEHOLDER", g)

    g = "Schedule"
    d["t_start"]               = I("2027-01-01", "date", "Month 1 (financial close / lease commencement)", "Analyst assumption; editable", "2026-09", "M", "Analyst", "", g)
    d["entitle_months"]   = I(18, "months", "Entitlement (OPA/ZBA + SPA) before Phase 1 construction", "Analyst assumption; planner input required", "2026-09", "L", "Planner", "PLACEHOLDER", g)
    d["entitle_delay"]    = I(0, "months", "Approval delay stress (added to entitlement)", "Scenario lever", "2026-09", "H", "Analyst", "", g)
    d["constr_m_8"]       = I(26, "months", "Construction duration, 8-storey building", "Analyst assumption; no schedule established (handoff §3)", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["constr_m_12"]      = I(30, "months", "Construction duration, 12-storey building", "Analyst assumption; no schedule established (handoff §3)", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["seq_mode"]         = I(1, "1 = sequential; 0 = overlap by interval", "Phase sequencing", "Handoff §6: sequential reference case", "2026-09", "H", "Analyst", "", g)
    d["overlap_interval"] = I(12, "months", "Start interval between phases when overlapping", "Scenario lever", "2026-09", "H", "Analyst", "", g)
    d["later_defer"]      = I(0, "months", "Defer Phases 2-4 (stress)", "Scenario lever", "2026-09", "H", "Analyst", "", g)
    d["lease_cap"]        = I(25, "homes/month", "Sitewide leasing capacity (all phases compete)", "Analyst assumption; no absorption evidence supplied", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["lease_cap_mult"]   = I(1.0, "×", "Leasing pace multiplier (stress)", "Scenario lever", "2026-09", "H", "Analyst", "", g)
    d["retail_delay"]     = I(6, "months", "Retail lease-up delay after phase completion", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["hold_after_stab"]  = I(120, "months", "Operating hold after final stabilization", "Handoff §1", "2026-09", "H", "Client", "", g)

    g = "Rents (market)"
    d["rent_0"] = I(1495, "$/mo", "Studio market rent", "NO COMP SUPPLIED. Analyst placeholder from 1-BR $/sf scaled; verify new-build achieved rents", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["rent_1"] = I(1795, "$/mo", "1-bedroom market rent", "Arterial leasing comps, 1 comp (650 sf resale condo, Centennial)", "2026-09", "L", "Arterial", "UNVERIFIED", g)
    d["rent_2"] = I(2173, "$/mo", "2-bedroom market rent", "Arterial leasing comps, 2 comps (avg 1,025 sf resale condo/townhouse)", "2026-09", "L", "Arterial", "UNVERIFIED", g)
    d["rent_3"] = I(2412, "$/mo", "3-bedroom market rent", "Arterial leasing comps, 4 comps (avg 1,300 sf, e.g. 53 Taunton Rd E $2,400-2,450)", "2026-09", "L", "Arterial", "UNVERIFIED", g)
    d["newbuild_prem"]    = I(0.0, "%", "New-build rent premium over comps", "Handoff §1: do not assume a premium", "2026-09", "H", "Analyst", "", g)
    d["rent_shock"]       = I(0.0, "%", "Starting-rent stress (±)", "Scenario lever", "2026-09", "H", "Analyst", "", g)
    d["comp_sf_1"] = I(650, "sf", "Comp reference size 1-BR", "Arterial comps", "2026-09", "M", "Arterial", "", g)
    d["comp_sf_2"] = I(1025, "sf", "Comp reference size 2-BR", "Arterial comps", "2026-09", "M", "Arterial", "", g)
    d["comp_sf_3"] = I(1300, "sf", "Comp reference size 3-BR", "Arterial comps", "2026-09", "M", "Arterial", "", g)

    g = "Rents (affordable / CMHC)"
    d["mmr_0"] = I(1250, "$/mo", "CMHC median market rent, bachelor (Oshawa CMA)", "PENDING — CMHC Rental Market Survey Oct 2025 (HMIP portal not reachable from this session); Arterial CMHC zone 'Oshawa (North)' returned no value", "2025-10", "L", "CMHC", "PENDING", g)
    d["mmr_1"] = I(1500, "$/mo", "CMHC median market rent, 1-bedroom", "PENDING — as above", "2025-10", "L", "CMHC", "PENDING", g)
    d["mmr_2"] = I(1750, "$/mo", "CMHC median market rent, 2-bedroom", "PENDING — as above", "2025-10", "L", "CMHC", "PENDING", g)
    d["mmr_3"] = I(1950, "$/mo", "CMHC median market rent, 3-bedroom", "PENDING — as above", "2025-10", "L", "CMHC", "PENDING", g)
    d["aff_share"]        = I(0.30, "% of homes", "Affordable share of homes (per type)", "Common Ground locked default", "2026-07", "H", "Ubuntu", "", g)
    d["aff_factor"]       = I(0.80, "% of MMR", "Affordable rent as % of CMHC MMR", "Common Ground locked default", "2026-07", "H", "Ubuntu", "", g)

    g = "Operating"
    d["rent_growth"]      = I(0.03, "%/yr", "Market rent growth", "Common Ground locked default", "2026-07", "M", "Ubuntu", "", g)
    d["opex_growth"]      = I(0.025, "%/yr", "Operating expense growth", "Common Ground locked default", "2026-07", "M", "Ubuntu", "", g)
    d["stab_occ"]         = I(0.97, "%", "Stabilized physical occupancy (vacancy 3%)", "Common Ground locked default (3% vacancy)", "2026-07", "M", "Ubuntu", "", g)
    d["bad_debt"]         = I(0.005, "% of rent", "Bad debt / concessions", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["opex_util"]        = I(1200, "$/occupied home/yr", "Utilities (variable)", "Common Ground generator default", "2026-07", "M", "Ubuntu", "", g)
    d["opex_rm"]          = I(900, "$/occupied home/yr", "Repairs, maintenance, turnover (variable)", "Common Ground generator default", "2026-07", "M", "Ubuntu", "", g)
    d["opex_ins"]         = I(450, "$/home/yr", "Insurance (fixed on delivered homes)", "Common Ground generator default", "2026-07", "M", "Ubuntu", "", g)
    d["opex_staff"]       = I(400, "$/home/yr", "On-site staffing (fixed)", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["opex_ga"]          = I(300, "$/home/yr", "Administration (fixed)", "Common Ground generator default", "2026-07", "M", "Ubuntu", "", g)
    d["mgmt_pct"]         = I(0.04, "% of EGI", "Property management", "Common Ground generator default", "2026-07", "M", "Ubuntu", "", g)
    d["reserve"]          = I(500, "$/home/yr", "Replacement reserve (below valuation NOI; deducted for lender DSCR)", "Common Ground generator default; treatment per handoff §4", "2026-07", "M", "Ubuntu", "", g)
    d["tax_rate"]         = I(0.009823, "% of assessed", "Multi-residential tax rate, Oshawa 2025", "Arterial (Oshawa 2025). UNVERIFIED — confirm class & current rate", "2025", "L", "Arterial", "UNVERIFIED", g)
    d["assess_ratio"]     = I(0.60, "× stabilized value", "Assessed value as share of stabilized value (new PBR, MPAC lag)", "Analyst placeholder; not cost, not price", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["cap_rate"]         = I(0.0425, "%", "Valuation cap rate", "Common Ground locked default", "2026-07", "M", "Ubuntu", "", g)
    d["cap_shock"]        = I(0.0, "bps as decimal", "Cap-rate stress (+0.01 = +100 bps)", "Scenario lever", "2026-09", "H", "Analyst", "", g)

    g = "Parking & retail"
    d["pk_m2_space"]      = I(32, "m²/space", "Below-grade gross area per space", "R01 volume allowance (unvalidated layout)", "2026-09", "L", "Design", "UNVERIFIED", g)
    d["pk_ratio"]         = I(0.65, "spaces/home", "Parking ratio (aggregate)", "R01; scenario lever", "2026-09", "M", "Design", "", g)
    d["pk_visitor"]       = I(0.10, "% of spaces", "Visitor share (not leased)", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["pk_takeup"]        = I(0.60, "% of resident spaces", "Resident take-up", "Analyst placeholder; no evidence supplied", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["pk_rent"]          = I(125, "$/space/mo", "Parking rent", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["ret_eff"]          = I(0.90, "%", "Retail rentable efficiency", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["ret_rent"]         = I(28, "$/sf/yr net", "Retail net rent", "Analyst placeholder; no retail comps supplied", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["ret_vac"]          = I(0.10, "%", "Retail vacancy", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["ret_ti"]           = I(60, "$/sf", "Retail TI + inducements", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)

    g = "Hard costs"
    d["hard_psf"]         = I(340, "$/sf above-ground gross", "Above-ground construction, up to 12 storeys (midpoint)", "Arterial / Altus benchmark $290-390/sf, 'Condos, up to 12 storeys'. Scope, basis and price date NOT inspected", "2026-09", "L", "Arterial", "UNVERIFIED", g)
    d["hard_shock"]       = I(0.0, "%", "Hard-cost stress (±)", "Scenario lever", "2026-09", "H", "Analyst", "", g)
    d["pk_psf"]           = I(200, "$/sf below-grade gross", "Underground parking incl. ramps/systems", "Analyst placeholder; assumes benchmark EXCLUDES parking (set 0 if included)", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["demo_m2"]          = I(75, "$/m² existing footprint", "Demolition + hazmat", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["exist_footprint"]  = I(6070, "m²", "Existing building footprint (5 buildings)", "Arterial site characteristics (2026-07)", "2026-07", "M", "Arterial", "", g)
    d["tenant_term"]      = I(500_000, "$", "Tenant relocation / lease termination", "Analyst placeholder; leases not provided", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["remed_offsite"]    = I(2_500_000, "$", "Remediation, earthworks, off-site & utility upgrades", "Analyst placeholder; no geotech/ESA/servicing study", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["site_m2"]          = I(150, "$/m² net site", "Streets, landscaping, public realm, shared infrastructure", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["infra_p1_share"]   = I(0.60, "%", "Share of site works + remediation paid in Phase 1 (front-loaded)", "Handoff §3: Phase 1 carries early shared infrastructure", "2026-09", "M", "Analyst", "", g)
    d["escalation"]       = I(0.03, "%/yr", "Escalation from benchmark date (Sept 2026) to phase mid-construction", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["cont_constr"]      = I(0.075, "% of hard", "Construction contingency", "Analyst placeholder", "2026-09", "M", "Analyst", "", g)
    d["cont_design"]      = I(0.025, "% of hard", "Design contingency", "Analyst placeholder", "2026-09", "M", "Analyst", "", g)

    g = "Soft costs & charges"
    d["design_pct"]       = I(0.08, "% of hard", "Design & consultants", "Analyst placeholder", "2026-09", "M", "Analyst", "", g)
    d["permit_pct"]       = I(0.01, "% of hard", "Approvals, permits, planning fees", "Analyst placeholder", "2026-09", "M", "Analyst", "", g)
    d["dc_per_home"]      = I(102_653, "$/home", "Development charges incl. Region + education (placeholder)", "Supplied placeholder $102,653/home; NOT a verified payable rate. City/Region 2025 schedules not reachable from this session", "2026-09", "L", "Client", "UNVERIFIED", g)
    d["dc_aff_exempt"]    = I(0, "1/0", "Affordable units DC-exempt (DC Act s.4.1)?", "Default 0 (conservative). Ontario exempts qualifying affordable rental units since June 2024; eligibility to be confirmed with City/Region", "2026-09", "M", "Analyst", "", g)
    d["dc_credits"]       = I(0, "$", "DC credits (existing units / prior payments)", "None evidenced", "2026-09", "L", "Analyst", "PENDING", g)
    d["lease_mkt"]        = I(1500, "$/home", "Leasing, marketing, pre-opening", "Analyst placeholder", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["dev_fee"]          = I(0.03, "% of hard+soft", "Development management fee (Ubuntu / NFP)", "Analyst placeholder", "2026-09", "M", "Analyst", "", g)
    d["builders_risk"]    = I(0.005, "% of hard", "Builder's risk & wrap-up insurance", "Analyst placeholder", "2026-09", "M", "Analyst", "", g)
    d["hst_net"]          = I(0.0, "% of hard+soft", "Net non-recoverable HST", "PENDING tax advice — enhanced GST rental rebate (2023) and Ontario rebate eligibility not confirmed", "2026-09", "L", "Tax", "PENDING", g)

    g = "Capital stack (Common Ground)"
    d["cmhc_share"]       = I(0.50, "% of TDC", "CMHC MLI Select insured loan", "Common Ground July 2026 (locked)", "2026-07", "H", "Ubuntu", "", g)
    d["bch_share"]        = I(0.40, "% of TDC", "Build Canada Homes tranche", "Common Ground July 2026 (locked)", "2026-07", "H", "Ubuntu", "", g)
    d["city_share"]       = I(0.05, "% of TDC", "City contribution", "Common Ground July 2026 (locked)", "2026-07", "H", "Ubuntu", "", g)
    d["grant_share"]      = I(0.545, "% of BCH", "BCH conditionally-repayable contribution share", "Common Ground July 2026 (~54.5%); raise if dscr10 < 1.10", "2026-07", "M", "Ubuntu", "", g)
    d["mezz_rate"]        = I(0.045, "%", "BCH mezz rate (BoC 10-yr + 100 bps), accrues from Year 10", "Common Ground July 2026; BoC bond yield UNVERIFIED", "2026-07", "L", "Ubuntu", "UNVERIFIED", g)
    d["mort_rate"]        = I(0.0425, "%", "MLI Select rate (construction + permanent)", "Common Ground locked default; lender term sheet NOT obtained", "2026-07", "L", "Ubuntu", "UNVERIFIED", g)
    d["rate_shock"]       = I(0.0, "decimal", "Borrowing-rate stress (+0.02 = +200 bps)", "Scenario lever", "2026-09", "H", "Analyst", "", g)
    d["amort_yrs"]        = I(50, "years", "MLI Select amortization", "Common Ground locked default", "2026-07", "H", "Ubuntu", "", g)
    d["cmhc_premium"]     = I(0.025, "% of loan", "CMHC insurance premium (MLI Select 100-point tier)", "Analyst placeholder; confirm premium schedule", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["cmhc_app_fee"]     = I(200, "$/home", "CMHC application fee", "CMHC schedule (analyst recollection)", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["leaseup_reserve_m"]= I(6, "months of MLI debt service", "Lease-up operating / interest reserve capitalized in TDC", "Analyst placeholder; lender completion-reserve requirement not obtained", "2026-09", "L", "Analyst", "PLACEHOLDER", g)
    d["takeout_ltv"]      = I(0.75, "%", "Year-20 takeout refinance LTV (must clear MLI + supplemental + mezz)", "Common Ground generator", "2026-07", "M", "Ubuntu", "", g)
    d["min_dscr"]         = I(1.10, "×", "Minimum DSCR (MLI Select affordability tier)", "CMHC MLI Select", "2026-07", "M", "CMHC", "", g)
    d["split_lo"]         = I(0.60, "%", "Landowner share of net cash flow, Yrs 1-10 (per phase)", "Common Ground July 2026 (locked)", "2026-07", "H", "Ubuntu", "", g)
    d["buyout_prem"]      = I(0.01, "%", "Supplemental charge sizing over buyout (fees)", "Common Ground generator", "2026-07", "M", "Ubuntu", "", g)
    d["sell_cost"]        = I(0.02, "% of price", "Selling costs (stabilize-and-sell test)", "Analyst placeholder", "2026-09", "M", "Analyst", "", g)
    d["target_yoc"]       = I(0.055, "%", "Hurdle: unlevered yield on cost (conventional cross-check)", "Handoff §7: labelled hurdle, not a house standard", "2026-09", "M", "Client", "", g)
    d["target_margin"]    = I(0.15, "%", "Hurdle: development margin on cost (conventional cross-check)", "Handoff §7: labelled hurdle", "2026-09", "M", "Client", "", g)
    d["disc_rate"]        = I(0.07, "%", "Discount rate for dated NPV", "Handoff §7: labelled hurdle", "2026-09", "M", "Client", "", g)
    return d

def vals(inputs): return {k: v["value"] for k, v in inputs.items()}

# --------------------------------------------------------------------------
# PROGRAMME  (R01 fixed quantities + density alternatives)
# --------------------------------------------------------------------------
R01_PHASES = [
    dict(ph=1, bldgs="A + B", storeys=8,  homes=250, above=21696, resi=20696, other=1000, retail=400, spaces=175, below=5600),
    dict(ph=2, bldgs="C",     storeys=12, homes=200, above=17568, resi=16768, other=800,  retail=300, spaces=125, below=4000),
    dict(ph=3, bldgs="D",     storeys=12, homes=200, above=17568, resi=16768, other=800,  retail=300, spaces=125, below=4000),
    dict(ph=4, bldgs="E + F", storeys=12, homes=350, above=29376, resi=29176, other=200,  retail=0,   spaces=225, below=7200),
]
SUITE_TYPES = [dict(t=0, label="Studio", homes=100, sf=450),
               dict(t=1, label="1 bedroom", homes=500, sf=600),
               dict(t=2, label="2 bedroom", homes=320, sf=800),
               dict(t=3, label="3 bedroom", homes=80, sf=1000)]

def allocate_suites(phases, suite_types):
    """Whole suites per phase, proportional with largest-remainder; disclose rounding."""
    total = sum(p["homes"] for p in phases)
    alloc = {p["ph"]: {} for p in phases}
    notes = []
    for st in suite_types:
        raw = [(p, st["homes"] * p["homes"] / total) for p in phases]
        fl = [(p, math.floor(r)) for p, r in raw]
        rem = st["homes"] - sum(f for _, f in fl)
        order = sorted(range(len(raw)), key=lambda i: -(raw[i][1] - fl[i][1]))
        counts = [f for _, f in fl]
        for i in order[:rem]: counts[i] += 1
        for (p, _), c in zip(raw, counts):
            alloc[p["ph"]][st["t"]] = c
        for (p, r), c in zip(raw, counts):
            if abs(r - c) > 1e-9:
                notes.append(f"Phase {p['ph']} {st['label']}: proportional {r:.2f} → {c}")
    # phase homes must equal sum of allocated suites
    for p in phases:
        s = sum(alloc[p["ph"]].values())
        assert s == p["homes"], (p, s)
    return alloc, notes

def density_programme(case: str):
    """Return (phases, suite_types, meta) for a density case.
    '1000' = R01 as modelled.  600/800/400 are analyst re-builds, NOT designs."""
    if case == "1000":
        return copy.deepcopy(R01_PHASES), copy.deepcopy(SUITE_TYPES), dict(
            label="R01 1,000 homes (modelled)", hard_psf_adj=1.0, note="Blender masterplan R01 — capacity concept")
    n = int(case)
    share = n / 1000
    # per-home ratios from R01 (disclosed), fewer storeys, fewer phases
    resi_per_home = 83408 / 1000
    if n == 800:
        cfg = [(1, "A + B", 8, 200), (2, "C", 10, 200), (3, "D", 10, 200), (4, "E", 10, 200)]
        adj, note = 0.97, "800 homes: 4 phases, 8-10 storeys; hard $/sf −3% (analyst)"
    elif n == 600:
        cfg = [(1, "A + B", 6, 200), (2, "C", 8, 200), (3, "D", 8, 200)]
        adj, note = 0.93, "600 homes: 3 phases, 6-8 storeys; hard $/sf −7% (analyst, wood/hybrid)"
    else:  # 400 downside
        cfg = [(1, "A + B", 6, 200), (2, "C", 6, 200)]
        adj, note = 0.90, "400 homes: 2 phases, 6 storeys; sensitivity only — NOT a confirmed as-of-right yield"
    phases = []
    for ph, b, st, h in cfg:
        resi = round(resi_per_home * h)
        other = round(2800 * h / n)
        retail = round(1000 * h / n) if ph < len(cfg) else max(0, 1000 - round(1000 / n) * (len(cfg) - 1) * 0)  # simple: pro rata
        retail = round(1000 * h / n)
        spaces = round(0.65 * h)
        phases.append(dict(ph=ph, bldgs=b, storeys=st, homes=h, above=resi + other, resi=resi, other=other,
                           retail=retail, spaces=spaces, below=spaces * 32))
    suites = [dict(st, homes=round(st["homes"] * share)) for st in SUITE_TYPES]
    diff = n - sum(s["homes"] for s in suites); suites[1]["homes"] += diff
    return phases, suites, dict(label=f"{n} homes (analyst re-build)", hard_psf_adj=adj, note=note)

# --------------------------------------------------------------------------
# ENGINE
# --------------------------------------------------------------------------
def pmt(rate_m, n, pv):
    return pv * rate_m / (1 - (1 + rate_m) ** -n) if rate_m else pv / n

def ltt_ontario(price):
    b = [(55000, .005), (250000, .01), (400000, .015), (2000000, .02), (float("inf"), .025)]
    t, lo = 0.0, 0.0
    for hi, r in b:
        if price > lo: t += (min(price, hi) - lo) * r
        lo = hi
    return t

@dataclass
class PhaseResult:
    ph: int
    tdc: float = 0.0
    hard: float = 0.0
    soft: float = 0.0
    dc: float = 0.0
    land_draw: float = 0.0
    land_eq: float = 0.0
    cap_int: float = 0.0
    fin_fees: float = 0.0
    mli: float = 0.0
    bch: float = 0.0
    bch_grant: float = 0.0
    bch_mezz: float = 0.0
    city: float = 0.0
    noi_stab: float = 0.0
    noi_lender: float = 0.0
    ads: float = 0.0
    dscr: float = 0.0
    dscr10: float = 0.0
    value: float = 0.0
    ltv: float = 0.0
    yoc: float = 0.0
    m_start: int = 0
    m_complete: int = 0
    m_stab: int = 0
    lo_cf10: float = 0.0
    np_cf: float = 0.0
    buyout10: float = 0.0
    sup: float = 0.0
    mezz_resid_end: float = 0.0
    unlev_irr: float = None
    months: dict = field(default_factory=dict)

def xirr(cfs, dates_m):
    """Annual IRR from monthly-dated cash flows (dates in months from t_start)."""
    def npv(r):
        return sum(cf / (1 + r) ** (m / 12.0) for cf, m in zip(cfs, dates_m))
    lo, hi = -0.99, 1.0
    if npv(lo) * npv(hi) > 0: return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if npv(mid) > 0: lo = mid
        else: hi = mid
    return (lo + hi) / 2

def run(inputs_vals: dict, phases=None, suite_types=None, meta=None, horizon=None, tax_seed=None):
    v = inputs_vals
    tax_seed = tax_seed or {}
    phases = copy.deepcopy(phases or R01_PHASES)
    suite_types = copy.deepcopy(suite_types or SUITE_TYPES)
    meta = meta or dict(hard_psf_adj=1.0)
    alloc, alloc_notes = allocate_suites(phases, suite_types)
    n_homes = sum(p["homes"] for p in phases)
    r_m = (v["mort_rate"] + v["rate_shock"]) / 12
    n_am = int(v["amort_yrs"] * 12)
    cap = v["cap_rate"] + v["cap_shock"]

    # ---------------- schedule
    ent = int(v["entitle_months"] + v["entitle_delay"])
    starts = []
    for i, p in enumerate(phases):
        dur = v["constr_m_8"] if p["storeys"] <= 8 else v["constr_m_12"]
        p["dur"] = int(dur)
        if i == 0:
            s = ent + 1
        else:
            prev = phases[i - 1]
            s = (prev["m_start"] + prev["dur"]) if v["seq_mode"] == 1 else (prev["m_start"] + int(v["overlap_interval"]))
            s += int(v["later_defer"]) if i == 1 else 0   # defer applies once, carried by chaining
        p["m_start"] = s
        p["m_complete"] = s + p["dur"] - 1    # last construction month; occupancy from m_complete+1
    # rough horizon: last completion + lease-up + hold
    last = max(p["m_complete"] for p in phases)
    lease_cap = max(1, v["lease_cap"] * v["lease_cap_mult"])
    H = horizon or int(last + math.ceil(n_homes / lease_cap) + 240 + 12)
    H = max(H, 300)

    # ---------------- per-phase static budget pieces
    sf = M2_TO_SF
    hard_psf = v["hard_psf"] * (1 + v["hard_shock"]) * meta.get("hard_psf_adj", 1.0)
    site_fixed = v["remed_offsite"] + v["site_m2"] * v["site_net_m2"]
    demo = v["demo_m2"] * v["exist_footprint"] + v["tenant_term"]
    nph = len(phases)
    total_draw_base = v["draw_pct"] * v["land_appraisal"]

    res = {}
    cons = {k: [0.0] * (H + 1) for k in [
        "delivered", "occupied", "rent_mkt", "rent_aff", "retail", "parking", "bad_debt", "egi",
        "opex_var", "opex_fixed", "mgmt", "tax", "noi", "reserve",
        "cost_land", "cost_hard", "cost_soft", "cost_dc", "cost_fin", "cost_total",
        "src_city", "src_grant", "src_mezz", "src_mli", "src_landeq",
        "mli_bal", "mli_ds", "mli_int", "mli_prin", "ncf", "lo_cf", "np_cf", "bch_sweep", "sup_ds", "buyout",
        "mezz_bal", "sup_bal", "existing_noi", "ret_ti", "reserve_draw", "reserve_bal"]}
    phase_out = []
    for p in phases:
        pr = PhaseResult(ph=p["ph"])
        pr.m_start, pr.m_complete = p["m_start"], p["m_complete"]
        h = p["homes"]
        mix = alloc[p["ph"]]
        # suites
        aff = {t: round(mix[t] * v["aff_share"]) for t in mix}
        mkt = {t: mix[t] - aff[t] for t in mix}
        rent = {t: v[f"rent_{t}"] * (1 + v["newbuild_prem"]) * (1 + v["rent_shock"]) for t in mix}
        arent = {t: round(v[f"mmr_{t}"] * v["aff_factor"]) for t in mix}
        gpr_mkt = sum(mkt[t] * rent[t] for t in mix) * 12
        gpr_aff = sum(aff[t] * arent[t] for t in mix) * 12
        gpr_m = (gpr_mkt + gpr_aff) / 12            # per month at 100% occupancy
        retail_sf = p["retail"] * sf * v["ret_eff"]
        retail_m = retail_sf * v["ret_rent"] / 12 * (1 - v["ret_vac"])
        spaces = p["spaces"]
        pk_m = spaces * (1 - v["pk_visitor"]) * v["pk_takeup"] * v["pk_rent"]

        # budget (static, escalated)
        above_sf = p["above"] * sf
        below_sf = p["below"] * sf
        mid = p["m_start"] + p["dur"] / 2
        esc = (1 + v["escalation"]) ** ((mid + 3) / 12.0)     # benchmark Sept 2026 -> t_start Jan 2027 (+3 mo) -> mid-build
        hard_ag = above_sf * hard_psf * esc
        hard_pk = below_sf * v["pk_psf"] * esc
        share_site = (v["infra_p1_share"] if p["ph"] == 1 else (1 - v["infra_p1_share"]) / (nph - 1)) if nph > 1 else 1.0
        hard_site = site_fixed * share_site
        hard_demo = demo if p["ph"] == 1 else 0.0
        hard_base = hard_ag + hard_pk + hard_site + hard_demo
        cont = hard_base * (v["cont_constr"] + v["cont_design"])
        hard = hard_base + cont
        dc_units = h - (sum(aff.values()) if v["dc_aff_exempt"] == 1 else 0)
        dc = dc_units * v["dc_per_home"] - (v["dc_credits"] * h / n_homes)
        soft_design = hard * v["design_pct"]
        soft_permit = hard * v["permit_pct"]
        soft_lease = h * v["lease_mkt"] + retail_sf * v["ret_ti"]
        soft_ins = hard * v["builders_risk"]
        soft_legal = v["legal_title"] if p["ph"] == 1 else 0.0
        carry = v["carry_taxes"] * ent / 12 if p["ph"] == 1 else 0.0
        soft_pre = soft_design + soft_permit + soft_lease + soft_ins + soft_legal + carry
        dev_fee = (hard + soft_pre + dc) * v["dev_fee"]
        hst = (hard + soft_pre) * v["hst_net"]
        # lease-up reserve: months of MLI debt service, sized on cost before reserve/financing (non-circular)
        dc_m = (r_m / (1 - (1 + r_m) ** -n_am)) if r_m else 1.0 / n_am
        reserve_lu = v["leaseup_reserve_m"] * (v["cmhc_share"] * (hard + dc + soft_pre + dev_fee + hst) / (1 - v["land_eq_pct"])) * dc_m
        soft = soft_pre + dev_fee + hst + reserve_lu
        # land recognition (Common Ground): draw + retained equity 5% of TDC
        draw = total_draw_base * (h / n_homes if v["draw_timing"] == 1 else (1.0 if p["ph"] == 1 else 0.0))
        ltt = ltt_ontario(draw) if v["ltt_applies"] == 1 else 0.0
        # Closed-form TDC (no circularity): MLI advances start once City + BCH sources are exhausted.
        # Cumulative spend is straight-line, so the MLI advance window = dur x (1 - (city+bch)/(1-eq)),
        # capitalized interest ~ MLI x r x window/2, CMHC premium = MLI x prem.
        base = hard + dc + soft + draw + ltt
        eq = v["land_eq_pct"]
        adv_months = p["dur"] * max(0.0, 1 - (v["city_share"] + v["bch_share"]) / (1 - eq))
        k_fin = v["cmhc_share"] * (r_m * adv_months / 2.0 + v["cmhc_premium"])
        tdc = (base + h * v["cmhc_app_fee"]) / ((1 - eq) - k_fin)
        cap_int = tdc * v["cmhc_share"] * r_m * adv_months / 2.0
        fees = tdc * v["cmhc_share"] * v["cmhc_premium"] + h * v["cmhc_app_fee"]
        land_eq = tdc * v["land_eq_pct"]
        mli = tdc * v["cmhc_share"]; bch = tdc * v["bch_share"]; city = tdc - mli - bch - land_eq
        grant = bch * v["grant_share"]; mezz = bch - grant
        pr.tdc, pr.hard, pr.soft, pr.dc, pr.land_draw, pr.land_eq = tdc, hard, soft, dc, draw + ltt, land_eq
        pr.cap_int, pr.fin_fees, pr.mli, pr.bch, pr.bch_grant, pr.bch_mezz, pr.city = cap_int, fees, mli, bch, grant, mezz, city
        # Static property-tax basis (closed form): NOI = NOI_pre / (1 + t*a/cap); tax = t*a*NOI/cap
        egi_s = (gpr_mkt + gpr_aff) * v["stab_occ"] * (1 - v["bad_debt"]) + retail_m * 12 + pk_m * 12 * v["stab_occ"]
        noi_pre = egi_s * (1 - v["mgmt_pct"]) - (v["opex_util"] + v["opex_rm"]) * h * v["stab_occ"] - (v["opex_ins"] + v["opex_staff"] + v["opex_ga"]) * h
        ta = v["tax_rate"] * v["assess_ratio"]
        noi_s = noi_pre / (1 + ta / cap) if cap else noi_pre
        tax_static = ta * noi_s / cap if cap else 0.0
        pr.tax_static = tax_static; pr.noi_static = noi_s; pr.egi_static = egi_s
        pr.adv_months = adv_months
        pr.budget = dict(hard_ag=hard_ag, hard_pk=hard_pk, hard_site=hard_site, hard_demo=hard_demo, cont=cont,
                         soft_design=soft_design, soft_permit=soft_permit, soft_lease=soft_lease, soft_ins=soft_ins,
                         soft_legal=soft_legal, carry=carry, dev_fee=dev_fee, hst=hst, ltt=ltt, esc=esc, reserve_lu=reserve_lu, dc_m=dc_m,
                         above_sf=above_sf, below_sf=below_sf, retail_sf=retail_sf, gpr_mkt=gpr_mkt, gpr_aff=gpr_aff,
                         mix=mix, aff=aff, mkt=mkt, rent=rent, arent=arent, spaces=spaces)
        ads = pmt(r_m, n_am, mli) * 12
        pr.ads = ads
        p["_pr"] = pr
        p["_ops"] = dict(gpr_m=gpr_m, retail_m=retail_m, pk_m=pk_m, aff_n=sum(aff.values()))
        phase_out.append(pr)

    # ---------------- monthly simulation (all phases share leasing capacity)
    occ = {p["ph"]: 0.0 for p in phases}
    mli_bal = {p["ph"]: 0.0 for p in phases}
    mezz_bal = {p["ph"]: 0.0 for p in phases}
    sup_bal = {p["ph"]: 0.0 for p in phases}
    res_bal = {p["ph"]: p["_pr"].budget["reserve_lu"] for p in phases}
    stab_m = {p["ph"]: None for p in phases}
    lo_done = {p["ph"]: False for p in phases}
    ph_series = {p["ph"]: {k: [0.0] * (H + 1) for k in cons} for p in phases}
    for m in range(1, H + 1):
        cap_left = lease_cap
        # existing income during entitlement (zero-retained default)
        if m <= ent:
            e = v["existing_noi"] * v["existing_noi_keep"] / 12
            ph_series[1]["existing_noi"][m] += e
        for p in phases:
            k = p["ph"]; pr = p["_pr"]; ops = p["_ops"]; s = ph_series[k]
            h = p["homes"]
            yrs_since_stab = None
            # ---- construction cash
            if p["m_start"] <= m <= p["m_complete"]:
                dur = p["dur"]; i = m - p["m_start"]
                s["cost_hard"][m] = pr.hard / dur
                s["cost_soft"][m] = pr.soft / dur
                if i == 0:
                    s["cost_dc"][m] = pr.dc
                    s["cost_land"][m] = pr.land_draw
                    s["cost_fin"][m] = pr.fin_fees
                # funding: city -> grant -> mezz -> MLI (cumulative)
            elif m == p["m_start"] - 1 and p["ph"] == 1:
                pass
            # ---- operations
            if m > p["m_complete"]:
                delivered = h
                s["delivered"][m] = delivered
                if occ[k] < h * v["stab_occ"]:
                    add = min(cap_left, h * v["stab_occ"] - occ[k])
                    occ[k] += add; cap_left -= add
                    if occ[k] >= h * v["stab_occ"] - 1e-9 and stab_m[k] is None:
                        stab_m[k] = m
                o = occ[k]
                s["occupied"][m] = o
                ym = (m - p["m_complete"] - 1) / 12.0
                g = (1 + v["rent_growth"]) ** math.floor(ym)
                gx = (1 + v["opex_growth"]) ** math.floor(ym)
                share_occ = o / h
                aff_n = ops["aff_n"]
                s["rent_mkt"][m] = pr.budget["gpr_mkt"] / 12 * share_occ * g
                s["rent_aff"][m] = pr.budget["gpr_aff"] / 12 * share_occ * g
                if m > p["m_complete"] + v["retail_delay"]:
                    s["retail"][m] = ops["retail_m"] * g
                s["parking"][m] = ops["pk_m"] * share_occ * g
                gross = s["rent_mkt"][m] + s["rent_aff"][m] + s["retail"][m] + s["parking"][m]
                s["bad_debt"][m] = -(s["rent_mkt"][m] + s["rent_aff"][m]) * v["bad_debt"]
                egi = gross + s["bad_debt"][m]
                s["egi"][m] = egi
                s["opex_var"][m] = -(v["opex_util"] + v["opex_rm"]) / 12 * o * gx
                s["opex_fixed"][m] = -(v["opex_ins"] + v["opex_staff"] + v["opex_ga"]) / 12 * h * gx
                s["mgmt"][m] = -egi * v["mgmt_pct"]
                # tax on assessed = ratio x stabilized value (phase), grown with opex
                s["tax"][m] = -pr.tax_static / 12 * gx
                s["noi"][m] = egi + s["opex_var"][m] + s["opex_fixed"][m] + s["mgmt"][m] + s["tax"][m]
                s["reserve"][m] = -v["reserve"] / 12 * h * gx
                if m == p["m_complete"] + 1:
                    # retail TI paid at completion
                    s["ret_ti"][m] = 0.0  # already in soft (leasing line)
        # second pass: debt & funding per phase (needs tax which needs noi_stab -> computed below on first stabilization)
        for p in phases:
            k = p["ph"]; pr = p["_pr"]; s = ph_series[k]
            cost = s["cost_hard"][m] + s["cost_soft"][m] + s["cost_dc"][m] + s["cost_land"][m] + s["cost_fin"][m]
            s["cost_total"][m] = cost
            if cost > 0:
                # cumulative funding waterfall
                cum_prev = sum(s["cost_total"][:m])
                cum_now = cum_prev + cost
                def take(lo, hi, cap_amt):
                    return max(0.0, min(hi, cap_amt) - min(lo, cap_amt))
                c1 = pr.city; c2 = c1 + pr.bch_grant; c3 = c2 + pr.bch_mezz
                s["src_city"][m] = take(cum_prev, cum_now, c1)
                s["src_grant"][m] = take(cum_prev, cum_now, c2) - s["src_city"][m]
                s["src_mezz"][m] = take(cum_prev, cum_now, c3) - s["src_city"][m] - s["src_grant"][m]
                s["src_mli"][m] = cost - s["src_city"][m] - s["src_grant"][m] - s["src_mezz"][m]
                mli_bal[k] += s["src_mli"][m]
                mezz_bal[k] += s["src_mezz"][m]
                # capitalized interest on MLI during construction
                if m <= p["m_complete"]:
                    ci = mli_bal[k] * r_m
                    s["mli_int"][m] = ci
                    mli_bal[k] += ci
                    s["src_mli"][m] += ci
                    s["cost_fin"][m] += ci
                    s["cost_total"][m] += ci
            if m > p["m_complete"] and mli_bal[k] > 0:
                # permanent amortizing loan from completion (loan = balance at completion)
                if m == p["m_complete"] + 1:
                    pr.mli_perm = mli_bal[k]
                    pr.pmt_m = pmt(r_m, n_am, pr.mli_perm)
                    pr.ads = pr.pmt_m * 12
                intr = mli_bal[k] * r_m
                prin = pr.pmt_m - intr
                s["mli_int"][m] = intr; s["mli_prin"][m] = prin; s["mli_ds"][m] = pr.pmt_m
                mli_bal[k] -= prin
                # supplemental charge service
                if sup_bal[k] > 0:
                    si = sup_bal[k] * r_m; sp = pr.sup_pmt - si
                    s["sup_ds"][m] = pr.sup_pmt; sup_bal[k] -= sp
                ncf = s["noi"][m] + s["reserve"][m] - s["mli_ds"][m] - s["sup_ds"][m]
                s["ncf"][m] = ncf
                # waterfall by phase-year since stabilization
                if stab_m[k] is not None and m > stab_m[k]:
                    ym = (m - stab_m[k] - 1) // 12 + 1
                else:
                    ym = 0
                if ym == 0:
                    # lease-up: deficits funded from the capitalized lease-up reserve; landowner clock starts at stabilization
                    draw = min(max(-ncf, 0.0), res_bal[k])
                    res_bal[k] -= draw
                    s["reserve_draw"][m] = draw
                    s["np_cf"][m] = ncf + draw
                elif ym <= 10:
                    s["lo_cf"][m] = ncf * v["split_lo"] if ncf > 0 else ncf * 0
                    s["np_cf"][m] = ncf - s["lo_cf"][m]
                    if ym == 10 and (m - stab_m[k]) % 12 == 0 and not lo_done[k]:
                        # Year-10 buyout: retained equity + MLI principal paid down, via 50-yr supplemental charge
                        paid = pr.mli_perm - mli_bal[k]
                        bo = pr.land_eq + paid
                        s["buyout"][m] = bo
                        pr.buyout10 = bo
                        pr.sup = bo * (1 + v["buyout_prem"])
                        sup_bal[k] = pr.sup
                        pr.sup_pmt = pmt(r_m, n_am, pr.sup)
                        lo_done[k] = True
                else:
                    # Yrs 11-20: BCH takes the 60% seat to sweep the mezz (accrues from Year 10)
                    if (m - stab_m[k]) % 12 == 1:
                        mezz_bal[k] *= (1 + v["mezz_rate"])
                    sweep = min(max(ncf * v["split_lo"], 0.0), mezz_bal[k]) if ym <= 20 else 0.0
                    s["bch_sweep"][m] = sweep; mezz_bal[k] -= sweep
                    s["np_cf"][m] = ncf - sweep
            s["mli_bal"][m] = mli_bal[k]; s["mezz_bal"][m] = mezz_bal[k]; s["sup_bal"][m] = sup_bal[k]; s["reserve_bal"][m] = res_bal[k]
            # stabilized NOI (first full 12 months after stabilization, annualised at month of stabilization)
        for p in phases:
            k = p["ph"]; pr = p["_pr"]; s = ph_series[k]
            if stab_m[k] is not None and m == stab_m[k] + 12 and not pr.noi_stab:
                # forward 12-month NOI at stabilization
                pr.noi_stab = sum(s["noi"][stab_m[k] + 1: stab_m[k] + 13])
                pr.noi_lender = pr.noi_stab + sum(s["reserve"][stab_m[k] + 1: stab_m[k] + 13])
                pr.m_stab = stab_m[k]
        for p in phases:
            for kk in cons:
                cons[kk][m] += ph_series[p["ph"]][kk][m]

    return dict(phases=phases, series=ph_series, cons=cons, H=H, alloc=alloc, alloc_notes=alloc_notes,
                stab_m=stab_m, ent=ent, lease_cap=lease_cap)

# --------------------------------------------------------------------------
# METRICS
# --------------------------------------------------------------------------
def metrics(r, v):
    cap = v["cap_rate"] + v["cap_shock"]
    out = dict(phases=[], H=r["H"])
    cons = r["cons"]
    for p in r["phases"]:
        pr = p["_pr"]; s = r["series"][p["ph"]]
        # recompute noi_stab in case seeding overwrote pass-2 value
        sm = r["stab_m"][p["ph"]]
        if sm is not None and sm + 12 <= r["H"]:
            pr.noi_stab = sum(s["noi"][sm + 1: sm + 13])
            pr.noi_lender = pr.noi_stab + sum(s["reserve"][sm + 1: sm + 13])
        pr.value = pr.noi_stab / cap if cap else 0
        pr.ltv = getattr(pr, "mli_perm", pr.mli) / pr.value if pr.value else None
        pr.dscr = pr.noi_lender / pr.ads if pr.ads else None
        pr.yoc = pr.noi_stab / pr.tdc if pr.tdc else None
        # year-10 DSCR incl supplemental
        if sm is not None:
            y10 = range(sm + 1 + 120, sm + 1 + 132)
            if y10[-1] <= r["H"]:
                noi10 = sum(s["noi"][m] + s["reserve"][m] for m in y10)
                ds10 = sum(s["mli_ds"][m] + s["sup_ds"][m] for m in y10)
                pr.dscr10 = noi10 / ds10 if ds10 else None
        pr.m_stab = sm if sm is not None else None
        # Year-20 takeout: refinance at takeout LTV on forward NOI must clear MLI + supplemental + mezz
        pr.takeout_ok = None; pr.refi20 = None; pr.debt20 = None; pr.value20 = None
        if sm is not None and sm + 252 <= r["H"]:
            m20 = sm + 240
            noi20 = sum(s["noi"][m20 + 1: m20 + 13])
            pr.value20 = noi20 / cap if cap else 0
            pr.refi20 = pr.value20 * v["takeout_ltv"]
            pr.debt20 = s["mli_bal"][m20] + s["sup_bal"][m20] + s["mezz_bal"][m20]
            pr.takeout_ok = pr.refi20 >= pr.debt20
            pr.mezz_bal20 = s["mezz_bal"][m20]
        # NFP annual net cash flow, phase-years 1..20 after stabilization (lease-up deficit is reserve-funded)
        pr.np_years = []
        if sm is not None:
            for y in range(1, 21):
                ms = range(sm + 1 + (y - 1) * 12, sm + 1 + y * 12)
                if ms[-1] <= r["H"]:
                    pr.np_years.append(sum(s["np_cf"][m] for m in ms))
        pr.np_min = min(pr.np_years) if pr.np_years else None
        pr.leaseup_deficit = -sum(min(0.0, s["ncf"][m]) for m in range(pr.m_complete + 1, (sm or r["H"]) + 1))
        pr.lo_cf10 = sum(s["lo_cf"])
        pr.np_cf = sum(s["np_cf"])
        pr.sweep = sum(s["bch_sweep"])
        pr.mezz_resid_end = s["mezz_bal"][r["H"]]
        pr.mli_end = s["mli_bal"][r["H"]]
        # unlevered project IRR on phase: -cost, +NOI, +terminal value at H (forward NOI / cap)
        cfs = [-(s["cost_total"][m] - s["src_landeq"][m]) + s["noi"][m] + s["reserve"][m] + s["existing_noi"][m] for m in range(1, r["H"] + 1)]
        term = (sum(s["noi"][r["H"] - 11: r["H"] + 1]) / cap) if cap else 0
        cfs[-1] += term
        pr.unlev_irr = xirr(cfs, list(range(1, r["H"] + 1)))
        out["phases"].append(pr)
    # consolidated
    H = r["H"]
    tdc = sum(pr.tdc for pr in out["phases"])
    noi_stab = sum(pr.noi_stab for pr in out["phases"])
    noi_lender = sum(pr.noi_lender for pr in out["phases"])
    ads = sum(pr.ads for pr in out["phases"])
    val = noi_stab / cap if cap else 0
    mli_perm = sum(getattr(pr, "mli_perm", pr.mli) for pr in out["phases"])
    cum_fund = 0.0; peak_public = 0.0; peak_mli = 0.0; cum_pub = 0.0
    for m in range(1, H + 1):
        cum_pub += cons["src_city"][m] + cons["src_grant"][m] + cons["src_mezz"][m]
        peak_public = max(peak_public, cum_pub)
        peak_mli = max(peak_mli, cons["mli_bal"][m])
    cfs = [-(cons["cost_total"][m]) + cons["noi"][m] + cons["reserve"][m] + cons["existing_noi"][m] for m in range(1, H + 1)]
    term = sum(cons["noi"][H - 11: H + 1]) / cap if cap else 0
    cfs[-1] += term
    irr = xirr(cfs, list(range(1, H + 1)))
    npv = sum(cf / (1 + v["disc_rate"]) ** (m / 12) for cf, m in zip(cfs, range(1, H + 1)))
    lo_draw = sum(pr.land_draw for pr in out["phases"])
    lo_eq = sum(pr.land_eq for pr in out["phases"])
    lo_cf = sum(pr.lo_cf10 for pr in out["phases"])
    lo_bo = sum(pr.buyout10 for pr in out["phases"])
    donated = max(v["land_appraisal"] - lo_draw - lo_eq, 0)
    # landowner IRR vs appraisal (what the land "earns" under Common Ground)
    lo_cfs = [cons["cost_land"][m] + cons["lo_cf"][m] + cons["buyout"][m] for m in range(1, H + 1)]
    lo_cfs[0] -= v["land_appraisal"]
    lo_irr = xirr(lo_cfs, list(range(1, H + 1)))
    stab_all = max([r["stab_m"][p["ph"]] or H for p in r["phases"]])
    out.update(dict(
        tdc=tdc, hard=sum(pr.hard for pr in out["phases"]), soft=sum(pr.soft for pr in out["phases"]),
        dc=sum(pr.dc for pr in out["phases"]), cap_int=sum(pr.cap_int for pr in out["phases"]),
        fin_fees=sum(pr.fin_fees for pr in out["phases"]),
        land_draw=lo_draw, land_eq=lo_eq, donated=donated,
        mli=sum(pr.mli for pr in out["phases"]), mli_perm=mli_perm, bch=sum(pr.bch for pr in out["phases"]),
        bch_grant=sum(pr.bch_grant for pr in out["phases"]), bch_mezz=sum(pr.bch_mezz for pr in out["phases"]),
        city=sum(pr.city for pr in out["phases"]),
        noi_stab=noi_stab, noi_lender=noi_lender, ads=ads, dscr=(noi_lender / ads if ads else None),
        value=val, ltv=(mli_perm / val if val else None), yoc=(noi_stab / tdc if tdc else None),
        margin=((val * (1 - v["sell_cost"]) - tdc) / tdc if tdc else None),
        peak_public=peak_public, peak_mli=peak_mli, unlev_irr=irr, npv=npv,
        lo_cf=lo_cf, lo_buyout=lo_bo, lo_total=lo_draw + lo_cf + lo_bo, lo_irr=lo_irr,
        np_cf=sum(pr.np_cf for pr in out["phases"]), bch_sweep=sum(pr.sweep for pr in out["phases"]),
        mezz_resid=sum(pr.mezz_resid_end for pr in out["phases"]), mli_end=sum(pr.mli_end for pr in out["phases"]),
        stab_all=stab_all, ent=r["ent"], homes=sum(p["homes"] for p in r["phases"]),
        aff_homes=sum(p["_ops"]["aff_n"] for p in r["phases"]),
        dscr10_min=min([pr.dscr10 for pr in out["phases"] if pr.dscr10] or [None]) if any(pr.dscr10 for pr in out["phases"]) else None,
        np_cf_min_year=None,
    ))
    # NFP minimum annual net cash flow (post-stabilization years)
    npy = []
    for y in range(1, H // 12 + 1):
        ms = range((y - 1) * 12 + 1, min(y * 12, H) + 1)
        npy.append(sum(cons["np_cf"][m] for m in ms))
    out["np_cf_annual"] = npy
    out["np_cf_min_year"] = min([pr.np_min for pr in out["phases"] if pr.np_min is not None] or [0.0])
    out["takeout_ok"] = all(pr.takeout_ok for pr in out["phases"] if pr.takeout_ok is not None)
    out["mezz_bal20"] = sum(getattr(pr, "mezz_bal20", 0.0) for pr in out["phases"])
    out["refi20"] = sum(pr.refi20 or 0 for pr in out["phases"]); out["debt20"] = sum(pr.debt20 or 0 for pr in out["phases"])
    out["leaseup_deficit"] = sum(pr.leaseup_deficit for pr in out["phases"])
    out["reserve_lu"] = sum(pr.budget["reserve_lu"] for pr in out["phases"])
    out["hold_end_m"] = stab_all + v["hold_after_stab"]
    # value - cost residual (conventional cross-check, market cap, no affordability)
    out["resid_simple"] = val * (1 - v["sell_cost"]) - (tdc - lo_draw - lo_eq) - v["target_margin"] * tdc
    return out

# --------------------------------------------------------------------------
# SOLVERS & SCENARIOS
# --------------------------------------------------------------------------
def evaluate(v, phases=None, suites=None, meta=None):
    r = run(v, phases, suites, meta)
    return r, metrics(r, v)

def passes(m, v):
    """Common Ground feasibility tests."""
    return (m["dscr"] is not None and m["dscr"] >= v["min_dscr"]
            and (m["dscr10_min"] is None or m["dscr10_min"] >= 1.10)
            and m["np_cf_min_year"] >= 0
            and m["takeout_ok"])

def test_detail(m, v):
    return dict(dscr=(m["dscr"], v["min_dscr"], m["dscr"] is not None and m["dscr"] >= v["min_dscr"]),
                dscr10=(m["dscr10_min"], 1.10, m["dscr10_min"] is None or m["dscr10_min"] >= 1.10),
                nfp=(m["np_cf_min_year"], 0.0, m["np_cf_min_year"] >= 0),
                takeout=(m["refi20"], m["debt20"], m["takeout_ok"]))

def solve_input(v, key, lo, hi, phases=None, suites=None, meta=None, increasing=True, iters=30):
    """Bisection on one input so that passes() is True. increasing=True: passing at lo, failing at hi (find max);
    increasing=False: failing at lo, passing at hi (find min)."""
    def ok(x):
        vv = dict(v); vv[key] = x
        _, m = evaluate(vv, phases, suites, meta)
        return passes(m, vv)
    if increasing:
        if not ok(lo): return None
        if ok(hi): return hi
        for _ in range(iters):
            mid = (lo + hi) / 2
            if ok(mid): lo = mid
            else: hi = mid
        return lo
    else:
        if ok(lo): return lo
        if not ok(hi): return None
        for _ in range(iters):
            mid = (lo + hi) / 2
            if ok(mid): hi = mid
            else: lo = mid
        return hi

def solve_land(v, phases=None, suites=None, meta=None, lo=0.0, hi=80e6):
    """Max land appraisal the Common Ground stack supports (all tests pass)."""
    base = dict(v)
    def ok(x):
        vv = dict(base); vv["land_appraisal"] = x
        _, m = evaluate(vv, phases, suites, meta)
        return passes(m, vv), m
    okl, ml = ok(lo)
    if not okl: return 0.0, ml
    okh, mh = ok(hi)
    if okh: return hi, mh
    for _ in range(28):
        mid = (lo + hi) / 2
        o, mm = ok(mid)
        if o: lo, ml = mid, mm
        else: hi = mid
    return lo, ml

def solve_land_conventional(v, phases=None, suites=None, meta=None):
    """Conventional cross-check: land price at which unlevered YoC = target and margin = target.
    Uses the same cost base but treats land as a purchase at the solved price (no CG stack)."""
    _, m = evaluate(dict(v, land_appraisal=0.0, land_eq_pct=0.0), phases, suites, meta)
    cost_ex_land = m["tdc"]
    noi = m["noi_stab"]; val = m["value"]
    land_yoc = noi / v["target_yoc"] - cost_ex_land
    land_margin = val * (1 - v["sell_cost"]) / (1 + v["target_margin"]) - cost_ex_land
    return dict(cost_ex_land=cost_ex_land, noi=noi, value=val, land_yoc=land_yoc, land_margin=land_margin,
                land_binding=min(land_yoc, land_margin))

def scenario_table(v):
    base_r, base_m = evaluate(v)
    rows = [("Reference case", {}, base_m)]
    shocks = [
        ("Rents −10%", dict(rent_shock=-0.10)),
        ("Rents +10%", dict(rent_shock=0.10)),
        ("Hard cost +15%", dict(hard_shock=0.15)),
        ("Hard cost −10%", dict(hard_shock=-0.10)),
        ("Cap rate +100 bps", dict(cap_shock=0.01)),
        ("Cap rate −50 bps", dict(cap_shock=-0.005)),
        ("Borrowing +200 bps", dict(rate_shock=0.02)),
        ("Borrowing −100 bps", dict(rate_shock=-0.01)),
        ("Approvals +12 months", dict(entitle_delay=12)),
        ("Approvals +24 months", dict(entitle_delay=24)),
        ("Leasing pace −50%", dict(lease_cap_mult=0.5)),
        ("Leasing pace +25%", dict(lease_cap_mult=1.25)),
        ("Parking 0.85/home", dict(pk_ratio=0.85)),
        ("Parking 1.00/home", dict(pk_ratio=1.00)),
        ("Parking 0.50/home", dict(pk_ratio=0.50)),
        ("Existing NOI retained during entitlement", dict(existing_noi_keep=1.0)),
        ("Phases 2-4 deferred 12 months", dict(later_defer=12)),
        ("Overlapping phases (12-mo interval)", dict(seq_mode=0)),
        ("Affordable units DC-exempt", dict(dc_aff_exempt=1)),
        ("Combined downside", dict(rent_shock=-0.10, hard_shock=0.15, cap_shock=0.01, entitle_delay=12, lease_cap_mult=0.5)),
    ]
    for label, ov in shocks:
        vv = dict(v); vv.update(ov)
        ph = None
        if "pk_ratio" in ov:
            ph = copy.deepcopy(R01_PHASES)
            for p in ph:
                p["spaces"] = round(p["homes"] * ov["pk_ratio"]); p["below"] = p["spaces"] * v["pk_m2_space"]
        _, m = evaluate(vv, ph)
        rows.append((label, ov, m))
    return rows

def two_way(v, xa, xs, ya, ys, fn):
    grid = []
    for yv in ys:
        row = []
        for xv in xs:
            vv = dict(v); vv[xa] = xv; vv[ya] = yv
            _, m = evaluate(vv)
            row.append(fn(m, vv))
        grid.append(row)
    return grid

def density_table(v):
    out = []
    for case in ["1000", "800", "600", "400"]:
        ph, su, meta = density_programme(case)
        _, m = evaluate(v, ph, su, meta)
        resid, _ = solve_land(v, ph, su, meta)
        conv = solve_land_conventional(v, ph, su, meta)
        out.append(dict(case=case, meta=meta, m=m, resid_cg=resid, conv=conv, phases=ph, suites=su))
    return out

if __name__ == "__main__":
    inp = default_inputs(); v = vals(inp)
    r, m = evaluate(v)
    print(json.dumps({k: (round(x) if isinstance(x, float) and abs(x) > 10 else x) for k, x in m.items()
                      if k not in ("phases", "np_cf_annual")}, indent=1, default=str))
    for pr in m["phases"]:
        print(pr.ph, "TDC", round(pr.tdc), "NOI", round(pr.noi_stab), "DSCR", round(pr.dscr or 0, 3), "DSCR10", pr.dscr10 and round(pr.dscr10, 3),
              "start", pr.m_start, "complete", pr.m_complete, "stab", pr.m_stab, "value", round(pr.value), "YoC", pr.yoc and round(pr.yoc, 4),
              "buyout", round(pr.buyout10), "mezz resid", round(pr.mezz_resid_end))
