#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Builds 1279_Simcoe_CommonGround_Model.xlsx — formula-driven, 14 tabs — from the
input register in cg_engine.py, then reconciles the live formulas (recalculated
headless in LibreOffice) against the engine's base case.

Usage:  python3 build_workbook.py [--out DIR]
"""
import os, sys, json, argparse, subprocess, shutil, datetime as dt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
import cg_engine as E

ESPRESSO, GOLD, CREAM, CREAM2, INK = "1F1612", "C89344", "F3EAD6", "F7EFDB", "161210"
INPUT_FILL = PatternFill("solid", fgColor="FFF4D6")     # editable input
FLAG_FILL  = PatternFill("solid", fgColor="F8D7DA")     # placeholder / unverified / pending
ENGINE_FILL= PatternFill("solid", fgColor="E8E4F3")     # engine-computed value (not a live formula)
HEAD_FILL  = PatternFill("solid", fgColor=ESPRESSO)
SUB_FILL   = PatternFill("solid", fgColor=CREAM2)
TOT_FILL   = PatternFill("solid", fgColor="EFE1C1")
thin = Side(style="thin", color="E0D5BD")
BORDER = Border(bottom=thin)
FMT_MONEY = '#,##0;[Red](#,##0);"-"'
FMT_MONEY2 = '#,##0.00;[Red](#,##0.00);"-"'
FMT_PCT = '0.00%'
FMT_X = '0.00"x"'
FMT_INT = '#,##0'
FMT_DATE = 'mmm-yy'

ADDR = "1279 Simcoe Street North, Oshawa, ON L1G 4X1"
PROJECT = "1279 Simcoe Street North"

def banner(ws, title, sub, width=10):
    ws["A1"] = "UBUNTU LAND TRUST & DEVELOPMENTS · Common Ground Initiative"
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=13); ws["A1"].fill = HEAD_FILL
    ws["A2"] = f"{PROJECT} — {title}"; ws["A2"].font = Font(bold=True, color=GOLD, size=11); ws["A2"].fill = HEAD_FILL
    ws["A3"] = sub; ws["A3"].font = Font(italic=True, color="FFFFFF", size=9); ws["A3"].fill = HEAD_FILL
    for r in (1, 2, 3):
        for c in range(2, width + 1):
            ws.cell(row=r, column=c).fill = HEAD_FILL
    ws.sheet_view.showGridLines = False

def hdr(ws, row, labels, start_col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=t)
        c.font = Font(bold=True, color=CREAM, size=10); c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center" if i else "left", vertical="center", wrap_text=True)

def put(ws, row, col, val, fmt=None, bold=False, fill=None, italic=False, color=None, align=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(bold=bold, italic=italic, size=10, color=color or INK)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if align: c.alignment = Alignment(horizontal=align)
    c.border = BORDER
    return c

# --------------------------------------------------------------------------
def build(out_dir):
    inputs = E.default_inputs(); v = E.vals(inputs)
    r, m = E.evaluate(v)
    H = r["H"]
    phases = r["phases"]; nph = len(phases)
    alloc, alloc_notes = r["alloc"], r["alloc_notes"]
    wb = Workbook()
    names = {}
    def dn(name, sheet, cell):
        wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!{cell}")
        names[name] = f"'{sheet}'!{cell}"

    # ======================================================================
    # READ ME
    ws = wb.active; ws.title = "Read Me"
    banner(ws, "Read Me", "Formula-driven Common Ground model — September 2026 handoff", 8)
    lines = [
        ("What this is", "An editable, formula-driven model of the 1,000-home R01 concept at 1279 Simcoe Street North, Oshawa, structured under the Ubuntu Land Trust Common Ground Initiative (land lease, CMHC MLI Select 50% / Build Canada Homes 40% split tranche / City 5% / landowner retained equity 5%; 30% of homes at 80% of CMHC median market rent). It answers the September 2026 handoff: test the $17M ask, solve supportable land value, compare density and timing alternatives, identify peak funding exposure."),
        ("Status", "PRELIMINARY SCENARIO BRIEF. The 1,000-home scheme is an unapproved capacity concept. Nothing here is a valuation, appraisal, offer or lender commitment."),
        ("Colour code", "Yellow cells = editable inputs (Inputs tab). Red cells = PLACEHOLDER / UNVERIFIED / PENDING inputs that must be replaced before reliance. Lilac cells = values computed by the companion engine (model/cg_engine.py) because they need goal-seek or data-table behaviour; re-run build_workbook.py after changing inputs to refresh them. Everything else is a live formula."),
        ("Under Common Ground the land is not bought", "The landowner receives a Day-1 draw (60% of appraisal) at each phase's financial close, keeps retained equity equal to 5% of that phase's total development cost, takes 60% of net cash flow after MLI Select debt service for the phase's Years 1-10, and is bought out at Year 10 (retained equity + MLI principal paid down) through a 50-year supplemental charge. Appraisal above draw + equity is value donated to affordability. BCH takes the 60% seat in Years 11-20 to sweep its mezzanine. The 'purchase price' in the handoff is therefore tested as the APPRAISAL that drives the draw, and 'supportable land value' is the largest appraisal at which every Common Ground test passes."),
        ("Common Ground feasibility tests", "1) Consolidated lender DSCR at stabilization ≥ min_dscr (NOI after replacement reserve ÷ MLI debt service). 2) Each phase's Year-10 DSCR incl. the supplemental charge ≥ 1.10. 3) Non-profit annual net cash flow ≥ 0 in every phase-year after stabilization (lease-up deficits are funded from the capitalized reserve). 4) Year-20 takeout: refinance at takeout_ltv × value clears MLI + supplemental + mezz balances."),
        ("Conventions", "Monthly periods; month 1 = t_start (Inputs). Nominal dollars; growth rates on Inputs. Rents step annually from each phase's completion. Stabilized NOI = the 12 months after the month occupancy first reaches stab_occ. Valuation NOI excludes the replacement reserve; lender DSCR NOI deducts it. Phase 1 pays front-loaded site/remediation works (infra_p1_share). Sources = uses each month: City → BCH contribution → BCH mezz → MLI Select, in that order; MLI construction interest is capitalized. Before income tax; land transfer tax off by default (land lease); HST net cost is a PENDING input."),
        ("Hurdles", "The handoff asks the investment team to select target IRR, minimum margin and hold period. Until supplied, the Exit & Residual tab shows several explicitly labelled hurdles (target_yoc, target_margin, disc_rate on Inputs) rather than a house standard."),
        ("Tabs", "Sources · Inputs · Programme · Schedule · Development Budget · Revenue & Opex · Debt · Monthly Cash Flow · Phase Returns · Exit & Residual · Sensitivities · Dashboard · Checks"),
        ("Companion files", "model/cg_engine.py (engine + input register), model/build_workbook.py (this builder), 1279_Simcoe_FINDINGS_MEMO.html, 1279_Simcoe_SOURCE_REGISTER.md"),
    ]
    rr = 5
    for k, t in lines:
        put(ws, rr, 1, k, bold=True); c = put(ws, rr, 2, t); c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[rr].height = max(30, 15 * (len(t) // 110 + 1)); rr += 1
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 120

    # ======================================================================
    # INPUTS  (register with defined names)
    wsI = wb.create_sheet("Inputs")
    banner(wsI, "Inputs & assumption register", "Yellow = editable · Red = placeholder/unverified/pending · every input carries source, date, units, confidence, owner", 9)
    hdr(wsI, 5, ["Group", "Key", "Value", "Unit", "Label", "Source / basis", "Date", "Conf.", "Owner", "Flag"])
    row = 6; input_row = {}
    cur_group = None
    for key, d in inputs.items():
        if d["group"] != cur_group:
            cur_group = d["group"]
            c = put(wsI, row, 1, cur_group, bold=True, fill=SUB_FILL)
            for cc in range(2, 11): wsI.cell(row=row, column=cc).fill = SUB_FILL
            row += 1
        put(wsI, row, 1, d["group"], color="8A7C6A")
        put(wsI, row, 2, key)
        val = d["value"]
        if key == "t_start":
            val = dt.datetime.strptime(val, "%Y-%m-%d")
        c = put(wsI, row, 3, val, fill=(FLAG_FILL if d["flag"] else INPUT_FILL))
        if isinstance(val, float) and (d["unit"].startswith("%") or "share" in d["unit"] or d["unit"] in ("decimal", "bps as decimal")):
            c.number_format = '0.00%'
        elif isinstance(val, (int, float)) and abs(val) >= 1000:
            c.number_format = FMT_INT
        elif key == "t_start":
            c.number_format = 'yyyy-mm-dd'
        put(wsI, row, 4, d["unit"]); put(wsI, row, 5, d["label"]); put(wsI, row, 6, d["source"])
        put(wsI, row, 7, d["date"]); put(wsI, row, 8, d["conf"], align="center"); put(wsI, row, 9, d["owner"])
        put(wsI, row, 10, d["flag"], bold=bool(d["flag"]), color=("A23B2C" if d["flag"] else None))
        dn(key, "Inputs", f"$C${row}"); input_row[key] = row
        row += 1
    # derived helper cells (formulas) — still on Inputs, below register
    row += 1; put(wsI, row, 1, "Derived (formulas)", bold=True, fill=SUB_FILL); row += 1
    derived = [
        ("r_m", "=(mort_rate+rate_shock)/12", "Monthly MLI rate"),
        ("n_am", "=amort_yrs*12", "Amortization periods"),
        ("cap_eff", "=cap_rate+cap_shock", "Effective cap rate"),
        ("ent_total", "=entitle_months+entitle_delay", "Entitlement months incl. delay"),
        ("cap_m", "=lease_cap*lease_cap_mult", "Effective sitewide leasing capacity, homes/month"),
        ("dc_const", "=r_m/(1-(1+r_m)^-n_am)", "Monthly debt constant (MLI)"),
        ("adv_share", "=MAX(0,1-(city_share+bch_share)/(1-land_eq_pct))", "Share of construction window funded by MLI advances"),
        ("k_fin", "=cmhc_share*(r_m*0/2+cmhc_premium)", "(placeholder, phase-specific k_fin on Budget)"),
        ("sf_per_m2", "=10.7639104", "1 m² = 10.7639104 sf"),
        ("horizon", f"={H}", "Model horizon, months (covers Year 20 of the last phase)"),
    ]
    for key, f, lab in derived:
        put(wsI, row, 2, key); c = put(wsI, row, 3, f); put(wsI, row, 5, lab)
        dn(key, "Inputs", f"$C${row}"); row += 1
    for col, w in zip("ABCDEFGHIJ", [22, 18, 14, 22, 46, 70, 10, 6, 10, 13]):
        wsI.column_dimensions[col].width = w

    # ======================================================================
    # PROGRAMME
    wsP = wb.create_sheet("Programme")
    banner(wsP, "Programme (R01 fixed quantities + suite allocation)", "Measured model quantities from Blender masterplan R01 — NOT certified zoning GFA, rentable area or QS areas", 8)
    hdr(wsP, 5, ["Quantity", "Unit"] + [f"Phase {p['ph']}" for p in phases] + ["Total", "Check / note"])
    PC = {p["ph"]: 3 + i for i, p in enumerate(phases)}   # phase columns C..F
    TC = 3 + nph
    prow = {}
    def prog_row(key, label, unit, vals_, fmt=FMT_INT, total=True, note="", is_input=True, formula_vals=None):
        nonlocal row
        put(wsP, row, 1, label); put(wsP, row, 2, unit)
        for p in phases:
            val = (formula_vals or vals_)[p["ph"]] if formula_vals else vals_[p["ph"]]
            c = put(wsP, row, PC[p["ph"]], val, fmt=fmt, fill=(INPUT_FILL if is_input and not formula_vals else None))
            dn(f"p{p['ph']}_{key}", "Programme", f"${L(PC[p['ph']])}${row}")
        if total:
            put(wsP, row, TC, f"=SUM({L(3)}{row}:{L(TC-1)}{row})", fmt=fmt, bold=True, fill=TOT_FILL)
            dn(f"tot_{key}", "Programme", f"${L(TC)}${row}")
        put(wsP, row, TC + 1, note, italic=True, color="8A7C6A")
        prow[key] = row; row += 1
    row = 6
    put(wsP, row, 1, "Buildings"); put(wsP, row, 2, "—")
    for p in phases: put(wsP, row, PC[p["ph"]], p["bldgs"], align="center")
    row += 1
    prog_row("storeys", "Storeys", "st", {p["ph"]: p["storeys"] for p in phases}, total=False, note="R01")
    prog_row("homes", "Homes", "homes", {p["ph"]: p["homes"] for p in phases}, note="R01 — must sum to 1,000 (Checks)")
    prog_row("above", "Above-ground area", "m²", {p["ph"]: p["above"] for p in phases}, note="R01 modeled floorplate area — 86,208 m²")
    prog_row("resi", "Residential gross area", "m²", {p["ph"]: p["resi"] for p in phases}, note="R01 — 83,408 m² (includes circulation/core)")
    prog_row("other", "Retail / amenity allowance", "m²", {p["ph"]: p["other"] for p in phases}, note="R01 — 2,800 m² (1,000 retail + 1,800 amenity, phase split unverified)")
    prog_row("retail", "  of which retail", "m²", {p["ph"]: p["retail"] for p in phases}, note="ANALYST ALLOCATION of the 1,000 m² retail — verify against design")
    prog_row("spaces_r01", "Parking spaces (R01)", "spaces", {p["ph"]: p["spaces"] for p in phases}, note="R01 aggregate 0.65/home; phase ratios differ")
    # spaces used = R01 unless pk_ratio changed
    put(wsP, row, 1, "Parking spaces (modelled)"); put(wsP, row, 2, "spaces")
    for p in phases:
        c = PC[p["ph"]]
        put(wsP, row, c, f"=IF(ROUND(pk_ratio,4)=0.65,{L(c)}{prow['spaces_r01']},ROUND({L(c)}{prow['homes']}*pk_ratio,0))", fmt=FMT_INT)
        dn(f"p{p['ph']}_spaces", "Programme", f"${L(c)}${row}")
    put(wsP, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=FMT_INT, bold=True, fill=TOT_FILL); dn("tot_spaces", "Programme", f"${L(TC)}${row}")
    put(wsP, row, TC + 1, "= R01 at pk_ratio 0.65; otherwise ROUND(homes × pk_ratio). An area allowance does not demonstrate physical fit.", italic=True, color="8A7C6A")
    prow["spaces"] = row; row += 1
    put(wsP, row, 1, "Below-grade area"); put(wsP, row, 2, "m²")
    for p in phases:
        c = PC[p["ph"]]; put(wsP, row, c, f"={L(c)}{prow['spaces']}*pk_m2_space", fmt=FMT_INT); dn(f"p{p['ph']}_below", "Programme", f"${L(c)}${row}")
    put(wsP, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=FMT_INT, bold=True, fill=TOT_FILL); dn("tot_below", "Programme", f"${L(TC)}${row}")
    put(wsP, row, TC + 1, "spaces × 32 m² — reconciles to R01 20,800 m² (Checks)", italic=True, color="8A7C6A"); prow["below"] = row; row += 1
    put(wsP, row, 1, "Concept FSI (above-ground ÷ net site)"); put(wsP, row, 2, "×")
    put(wsP, row, TC, f"={L(TC)}{prow['above']}/site_net_m2", fmt='0.000', bold=True, fill=TOT_FILL); row += 2

    # suite allocation
    put(wsP, row, 1, "Suite allocation (whole suites, proportional, largest-remainder)", bold=True, fill=SUB_FILL)
    for cc in range(2, TC + 2): wsP.cell(row=row, column=cc).fill = SUB_FILL
    row += 1
    hdr(wsP, row, ["Suite type", "Target sf"] + [f"Phase {p['ph']}" for p in phases] + ["Total", "R01 target / rounding"]); row += 1
    srow = {}
    for st in E.SUITE_TYPES:
        t = st["t"]
        put(wsP, row, 1, st["label"]); c = put(wsP, row, 2, st["sf"], fmt=FMT_INT, fill=INPUT_FILL); dn(f"sf_{t}", "Programme", f"$B${row}")
        for p in phases:
            cc = PC[p["ph"]]; put(wsP, row, cc, alloc[p["ph"]][t], fmt=FMT_INT, fill=INPUT_FILL); dn(f"p{p['ph']}_n{t}", "Programme", f"${L(cc)}${row}")
        put(wsP, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=FMT_INT, bold=True, fill=TOT_FILL); dn(f"tot_n{t}", "Programme", f"${L(TC)}${row}")
        put(wsP, row, TC + 1, f"target {st['homes']}", italic=True, color="8A7C6A")
        srow[t] = row; row += 1
    put(wsP, row, 1, "Total homes", bold=True)
    for p in phases:
        cc = PC[p["ph"]]; put(wsP, row, cc, f"=SUM({L(cc)}{srow[0]}:{L(cc)}{srow[3]})", fmt=FMT_INT, bold=True)
    put(wsP, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=FMT_INT, bold=True, fill=TOT_FILL); row += 1
    put(wsP, row, 1, "Suite rentable area"); put(wsP, row, 2, "sf")
    for p in phases:
        cc = PC[p["ph"]]
        put(wsP, row, cc, "=" + "+".join(f"{L(cc)}{srow[t]}*sf_{t}" for t in range(4)), fmt=FMT_INT); dn(f"p{p['ph']}_nra", "Programme", f"${L(cc)}${row}")
    put(wsP, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=FMT_INT, bold=True, fill=TOT_FILL); dn("tot_nra", "Programme", f"${L(TC)}${row}")
    put(wsP, row, TC + 1, "R01 681,000 sf (Checks). Suite area / residential gross ≈ 75.85%", italic=True, color="8A7C6A"); row += 1
    put(wsP, row, 1, "Affordable homes (per type, 30%)", bold=True); row += 1
    arow = {}
    for st in E.SUITE_TYPES:
        t = st["t"]; put(wsP, row, 1, f"  {st['label']} affordable")
        for p in phases:
            cc = PC[p["ph"]]; put(wsP, row, cc, f"=ROUND({L(cc)}{srow[t]}*aff_share,0)", fmt=FMT_INT); dn(f"p{p['ph']}_a{t}", "Programme", f"${L(cc)}${row}")
        put(wsP, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=FMT_INT, bold=True, fill=TOT_FILL); arow[t] = row; row += 1
    put(wsP, row, 1, "Affordable homes total")
    for p in phases:
        cc = PC[p["ph"]]; put(wsP, row, cc, f"=SUM({L(cc)}{arow[0]}:{L(cc)}{arow[3]})", fmt=FMT_INT); dn(f"p{p['ph']}_aff", "Programme", f"${L(cc)}${row}")
    put(wsP, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=FMT_INT, bold=True, fill=TOT_FILL); dn("tot_aff", "Programme", f"${L(TC)}${row}"); row += 2
    put(wsP, row, 1, "Rounding disclosure:", bold=True); row += 1
    for n in (alloc_notes or ["No rounding — proportional allocation produced whole suites."]):
        put(wsP, row, 1, n, italic=True, color="8A7C6A"); row += 1
    for col, w in zip("ABCDEFGH", [40, 10, 13, 13, 13, 13, 14, 70]): wsP.column_dimensions[col].width = w

    # ======================================================================
    # SCHEDULE
    wsS = wb.create_sheet("Schedule")
    banner(wsS, "Schedule", "No construction dates, durations or lease-up rates have been established (handoff §3) — all schedule inputs are analyst placeholders", 8)
    hdr(wsS, 5, ["Milestone", "Unit"] + [f"Phase {p['ph']}" for p in phases] + ["Note"])
    row = 6
    def srow_(key, label, unit, fn, fmt=FMT_INT, note=""):
        nonlocal row
        put(wsS, row, 1, label); put(wsS, row, 2, unit)
        for i, p in enumerate(phases):
            cc = PC[p["ph"]]; put(wsS, row, cc, fn(i, p, cc), fmt=fmt); dn(f"p{p['ph']}_{key}", "Schedule", f"${L(cc)}${row}")
        put(wsS, row, TC, note, italic=True, color="8A7C6A"); row += 1
    srow_("dur", "Construction duration", "months", lambda i, p, c: f"=IF(p{p['ph']}_storeys<=8,constr_m_8,constr_m_12)", note="8-storey vs 12-storey placeholder durations")
    srow_("start", "Construction start (month #)", "month", lambda i, p, c: ("=ent_total+1" if i == 0 else f"=IF(seq_mode=1,p{phases[i-1]['ph']}_start+p{phases[i-1]['ph']}_dur,p{phases[i-1]['ph']}_start+overlap_interval)+IF({i}=1,later_defer,0)"), note="Sequential: next phase starts when previous completes; overlap: fixed interval; later_defer chains through Phases 2-4")
    srow_("complete", "Construction complete (month #)", "month", lambda i, p, c: f"=p{p['ph']}_start+p{p['ph']}_dur-1", note="Occupancy begins the following month")
    srow_("first_occ", "First occupancy (month #)", "month", lambda i, p, c: f"=p{p['ph']}_complete+1")
    srow_("start_d", "Construction start", "date", lambda i, p, c: f"=EDATE(t_start,p{p['ph']}_start-1)", fmt=FMT_DATE)
    srow_("first_occ_d", "First occupancy", "date", lambda i, p, c: f"=EDATE(t_start,p{p['ph']}_complete)", fmt=FMT_DATE)
    srow_("esc", "Escalation factor to mid-construction", "×", lambda i, p, c: f"=(1+escalation)^((p{p['ph']}_start+p{p['ph']}_dur/2+3)/12)", fmt='0.0000', note="Benchmark priced Sept 2026; t_start is 3 months later")
    srow_("stab", "Stabilization (month #)", "month", lambda i, p, c: f"=IFERROR(MATCH(1,'Monthly Cash Flow'!$D${100+0}:$D${100+0},0),0)", note="(overwritten below once Monthly Cash Flow rows are known)")
    stab_row = row - 1
    srow_("stab_d", "Stabilization", "date", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,EDATE(t_start,p{p['ph']}_stab-1),\"n/a\")", fmt=FMT_DATE)
    srow_("y10_m", "Landowner buyout (month #)", "month", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,p{p['ph']}_stab+120,0)", note="Year 10 after stabilization")
    srow_("y20_m", "Takeout refinance (month #)", "month", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,p{p['ph']}_stab+240,0)", note="Year 20: mezz must be cleared")
    row += 1
    put(wsS, row, 1, "Entitlement (months incl. delay)"); put(wsS, row, 3, "=ent_total", fmt=FMT_INT); row += 1
    put(wsS, row, 1, "All phases stabilized (month #)"); put(wsS, row, 3, "=MAX(" + ",".join(f"p{p['ph']}_stab" for p in phases) + ")", fmt=FMT_INT); dn("stab_all", "Schedule", f"$C${row}"); row += 1
    put(wsS, row, 1, "Hold end = final stabilization + hold (month #)"); put(wsS, row, 3, "=stab_all+hold_after_stab", fmt=FMT_INT); dn("hold_end", "Schedule", f"$C${row}"); row += 1
    put(wsS, row, 1, "Hold end"); put(wsS, row, 3, "=EDATE(t_start,hold_end-1)", fmt=FMT_DATE); row += 1
    put(wsS, row, 1, "Model horizon (months)"); put(wsS, row, 3, "=horizon", fmt=FMT_INT); row += 1
    for col, w in zip("ABCDEFGH", [40, 10, 13, 13, 13, 13, 70, 10]): wsS.column_dimensions[col].width = w

    # ======================================================================
    # DEVELOPMENT BUDGET
    wsB = wb.create_sheet("Development Budget")
    banner(wsB, "Development budget (nominal, escalated to each phase's mid-construction)", "Benchmark $/sf scope not inspected; parking, site works and contingency added separately (double-count risk noted)", 8)
    hdr(wsB, 5, ["Line", "Basis"] + [f"Phase {p['ph']}" for p in phases] + ["Total", "Source / note"])
    row = 6; brow = {}
    def brow_(key, label, basis, fn, note="", total=True, fmt=FMT_MONEY, bold=False, fill=None, name=True):
        nonlocal row
        put(wsB, row, 1, label, bold=bold, fill=fill); put(wsB, row, 2, basis, fill=fill)
        for i, p in enumerate(phases):
            cc = PC[p["ph"]]; put(wsB, row, cc, fn(i, p, cc), fmt=fmt, bold=bold, fill=fill)
            if name: dn(f"p{p['ph']}_{key}", "Development Budget", f"${L(cc)}${row}")
        if total:
            put(wsB, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=fmt, bold=True, fill=(fill or TOT_FILL))
            if name: dn(f"tot_{key}", "Development Budget", f"${L(TC)}${row}")
        put(wsB, row, TC + 1, note, italic=True, color="8A7C6A"); brow[key] = row; row += 1
    def sec(t):
        nonlocal row
        put(wsB, row, 1, t, bold=True, fill=SUB_FILL)
        for cc in range(2, TC + 2): wsB.cell(row=row, column=cc).fill = SUB_FILL
        row += 1
    sec("1. Land & acquisition (Common Ground: lease, not purchase)")
    brow_("land_draw", "Landowner Day-1 draw", "draw_pct × appraisal, timed by draw_timing", lambda i, p, c: f"=draw_pct*land_appraisal*IF(draw_timing=1,p{p['ph']}_homes/tot_homes,IF({i}=0,1,0))", note="60% of appraisal (Common Ground). Appraisal = seller ask PLACEHOLDER")
    brow_("ltt", "Land transfer tax", "Ontario LTT on draw if ltt_applies", lambda i, p, c: f"=IF(ltt_applies=1,IF(p{p['ph']}_land_draw>0,0.005*MIN(p{p['ph']}_land_draw,55000)+IF(p{p['ph']}_land_draw>55000,0.01*(MIN(p{p['ph']}_land_draw,250000)-55000),0)+IF(p{p['ph']}_land_draw>250000,0.015*(MIN(p{p['ph']}_land_draw,400000)-250000),0)+IF(p{p['ph']}_land_draw>400000,0.02*(MIN(p{p['ph']}_land_draw,2000000)-400000),0)+IF(p{p['ph']}_land_draw>2000000,0.025*(p{p['ph']}_land_draw-2000000),0),0),0)", note="Off by default — land lease, no transfer")
    brow_("legal", "Legal, title, lease registration, due diligence", "lump sum, Phase 1", lambda i, p, c: f"=IF({i}=0,legal_title,0)", note="PLACEHOLDER")
    brow_("carry", "Carrying costs during entitlement", "carry_taxes × ent_total/12, Phase 1", lambda i, p, c: f"=IF({i}=0,carry_taxes*ent_total/12,0)", note="PLACEHOLDER pending tax bill")
    sec("2-3. Demolition, remediation, site")
    brow_("demo", "Demolition, hazmat, tenant relocation / lease termination", "demo_m2 × footprint + tenant_term, Phase 1", lambda i, p, c: f"=IF({i}=0,demo_m2*exist_footprint+tenant_term,0)", note="PLACEHOLDER; leases and hazmat survey not provided")
    brow_("remed", "Remediation, earthworks, off-site & utility upgrades", "remed_offsite, front-loaded by infra_p1_share", lambda i, p, c: f"=remed_offsite*IF({i}=0,infra_p1_share,(1-infra_p1_share)/{nph-1})", note="PLACEHOLDER; no geotech / ESA / servicing study")
    brow_("site", "Streets, landscaping, public realm, shared infrastructure", "site_m2 × net site, front-loaded", lambda i, p, c: f"=site_m2*site_net_m2*IF({i}=0,infra_p1_share,(1-infra_p1_share)/{nph-1})", note="Phase 1 carries early shared infrastructure (handoff §3)")
    sec("4-5. Building construction")
    brow_("above_sf", "Above-ground gross area", "m² × 10.7639", lambda i, p, c: f"=p{p['ph']}_above*sf_per_m2", note="R01 floorplate area, not QS gross", fmt=FMT_INT)
    brow_("below_sf", "Below-grade gross area", "m² × 10.7639", lambda i, p, c: f"=p{p['ph']}_below*sf_per_m2", fmt=FMT_INT)
    brow_("hard_ag", "Above-ground construction (residential, retail shell, amenity)", "above_sf × hard_psf × (1+hard_shock) × escalation", lambda i, p, c: f"=p{p['ph']}_above_sf*hard_psf*(1+hard_shock)*p{p['ph']}_esc", note="Arterial/Altus $290-390 midpoint — scope UNVERIFIED")
    brow_("hard_pk", "Underground parking, ramps, systems", "below_sf × pk_psf × escalation", lambda i, p, c: f"=p{p['ph']}_below_sf*pk_psf*(1+hard_shock)*p{p['ph']}_esc", note="PLACEHOLDER; set pk_psf = 0 if benchmark includes parking")
    brow_("hard_base", "Hard cost before contingency", "sum", lambda i, p, c: f"=p{p['ph']}_demo+p{p['ph']}_remed+p{p['ph']}_site+p{p['ph']}_hard_ag+p{p['ph']}_hard_pk", bold=True)
    brow_("cont", "Construction + design contingency", "(cont_constr + cont_design) × hard base", lambda i, p, c: f"=p{p['ph']}_hard_base*(cont_constr+cont_design)")
    brow_("hard", "TOTAL HARD COST", "", lambda i, p, c: f"=p{p['ph']}_hard_base+p{p['ph']}_cont", bold=True, fill=TOT_FILL)
    brow_("hard_psf_ag", "Hard cost per above-ground gross sf", "hard ÷ above_sf", lambda i, p, c: f"=p{p['ph']}_hard/p{p['ph']}_above_sf", fmt=FMT_MONEY2, total=False)
    sec("7. Development charges & municipal obligations")
    brow_("dc_units", "DC-paying homes", "homes less exempt affordable if dc_aff_exempt", lambda i, p, c: f"=p{p['ph']}_homes-IF(dc_aff_exempt=1,p{p['ph']}_aff,0)", fmt=FMT_INT)
    brow_("dc", "Development + education charges (net of credits)", "dc_units × dc_per_home − credits (pro rata)", lambda i, p, c: f"=p{p['ph']}_dc_units*dc_per_home-dc_credits*p{p['ph']}_homes/tot_homes", note="$102,653/home PLACEHOLDER — not a verified payable rate; payment at construction start")
    sec("7, 9, 10. Soft costs")
    brow_("design", "Design & consultants", "design_pct × hard", lambda i, p, c: f"=p{p['ph']}_hard*design_pct")
    brow_("permit", "Approvals, permits, planning fees", "permit_pct × hard", lambda i, p, c: f"=p{p['ph']}_hard*permit_pct")
    brow_("retail_sf", "Retail rentable area", "retail m² × 10.7639 × ret_eff", lambda i, p, c: f"=p{p['ph']}_retail*sf_per_m2*ret_eff", fmt=FMT_INT)
    brow_("leasing", "Leasing, marketing, pre-opening + retail TI", "lease_mkt × homes + ret_ti × retail sf", lambda i, p, c: f"=lease_mkt*p{p['ph']}_homes+ret_ti*p{p['ph']}_retail_sf")
    brow_("brisk", "Builder's risk & wrap-up insurance", "builders_risk × hard", lambda i, p, c: f"=p{p['ph']}_hard*builders_risk")
    brow_("soft_pre", "Soft costs before fee, HST, reserve", "sum incl. legal + carry", lambda i, p, c: f"=p{p['ph']}_design+p{p['ph']}_permit+p{p['ph']}_leasing+p{p['ph']}_brisk+p{p['ph']}_legal+p{p['ph']}_carry", bold=True)
    brow_("devfee", "Development management fee", "dev_fee × (hard + soft_pre + DC)", lambda i, p, c: f"=(p{p['ph']}_hard+p{p['ph']}_soft_pre+p{p['ph']}_dc)*dev_fee")
    brow_("hst", "Net non-recoverable HST", "hst_net × (hard + soft_pre)", lambda i, p, c: f"=(p{p['ph']}_hard+p{p['ph']}_soft_pre)*hst_net", note="PENDING tax advice — 0 until confirmed")
    brow_("reserve_lu", "Lease-up operating / interest reserve (capitalized)", "leaseup_reserve_m × MLI debt service on cost before financing", lambda i, p, c: f"=leaseup_reserve_m*cmhc_share*(p{p['ph']}_hard+p{p['ph']}_dc+p{p['ph']}_soft_pre+p{p['ph']}_devfee+p{p['ph']}_hst)/(1-land_eq_pct)*dc_const", note="Completion-reserve requirement not obtained from lender")
    brow_("soft", "TOTAL SOFT COST", "", lambda i, p, c: f"=p{p['ph']}_soft_pre+p{p['ph']}_devfee+p{p['ph']}_hst+p{p['ph']}_reserve_lu", bold=True, fill=TOT_FILL)
    sec("11. Financing (closed-form; monthly actuals on Monthly Cash Flow)")
    brow_("land", "Land recognized in cash uses (draw + LTT)", "", lambda i, p, c: f"=p{p['ph']}_land_draw+p{p['ph']}_ltt")
    brow_("base", "Cash uses before financing", "hard + DC + soft + land", lambda i, p, c: f"=p{p['ph']}_hard+p{p['ph']}_dc+p{p['ph']}_soft+p{p['ph']}_land", bold=True)
    brow_("adv_m", "MLI advance window", "dur × adv_share (months)", lambda i, p, c: f"=p{p['ph']}_dur*adv_share", fmt='0.0', total=False, note="MLI advances start once City + BCH are exhausted")
    brow_("kfin", "Financing factor k", "cmhc_share × (r_m × window/2 + premium)", lambda i, p, c: f"=cmhc_share*(r_m*p{p['ph']}_adv_m/2+cmhc_premium)", fmt='0.0000', total=False)
    brow_("tdc", "TOTAL DEVELOPMENT COST (incl. retained land equity)", "(base + app fees) ÷ ((1 − land_eq_pct) − k)", lambda i, p, c: f"=(p{p['ph']}_base+p{p['ph']}_homes*cmhc_app_fee)/((1-land_eq_pct)-p{p['ph']}_kfin)", bold=True, fill=TOT_FILL, note="Closed form: TDC = cash uses + capitalized interest + CMHC fees + in-kind land equity")
    brow_("capint", "Capitalized MLI construction interest (plan)", "TDC × cmhc_share × r_m × window/2", lambda i, p, c: f"=p{p['ph']}_tdc*cmhc_share*r_m*p{p['ph']}_adv_m/2", note="Actual monthly interest on Monthly Cash Flow; reconciliation on Checks")
    brow_("fees", "CMHC premium + application fees", "MLI × premium + homes × fee", lambda i, p, c: f"=p{p['ph']}_tdc*cmhc_share*cmhc_premium+p{p['ph']}_homes*cmhc_app_fee", note="Premium schedule PLACEHOLDER")
    brow_("landeq", "Landowner retained land equity (in-kind)", "land_eq_pct × TDC", lambda i, p, c: f"=p{p['ph']}_tdc*land_eq_pct", note="Converts to the Year-10 buyout claim")
    brow_("tdc_home", "TDC per home", "", lambda i, p, c: f"=p{p['ph']}_tdc/p{p['ph']}_homes", total=False, fmt=FMT_MONEY)
    brow_("tdc_sf", "TDC per gross sf (above + below)", "", lambda i, p, c: f"=p{p['ph']}_tdc/(p{p['ph']}_above_sf+p{p['ph']}_below_sf)", total=False, fmt=FMT_MONEY2)
    sec("Capital stack (Common Ground July 2026)")
    brow_("mli", "CMHC MLI Select (plan)", "cmhc_share × TDC", lambda i, p, c: f"=p{p['ph']}_tdc*cmhc_share")
    brow_("bch", "Build Canada Homes tranche", "bch_share × TDC", lambda i, p, c: f"=p{p['ph']}_tdc*bch_share")
    brow_("grant", "  BCH conditionally-repayable contribution", "grant_share × BCH", lambda i, p, c: f"=p{p['ph']}_bch*grant_share", note="Vests Yrs 1-10 vs deep affordability; clawback on early sale")
    brow_("mezz", "  BCH mezzanine (BoC+100, payment-free to Yr 10)", "BCH − contribution", lambda i, p, c: f"=p{p['ph']}_bch-p{p['ph']}_grant")
    brow_("city", "City contribution", "TDC − MLI − BCH − land equity", lambda i, p, c: f"=p{p['ph']}_tdc-p{p['ph']}_mli-p{p['ph']}_bch-p{p['ph']}_landeq")
    brow_("stack_chk", "Stack − TDC (must be 0)", "", lambda i, p, c: f"=p{p['ph']}_mli+p{p['ph']}_bch+p{p['ph']}_city+p{p['ph']}_landeq-p{p['ph']}_tdc", fmt=FMT_MONEY)
    brow_("donated", "Appraised value donated to affordability", "appraisal − Σdraw − Σequity (project level)", lambda i, p, c: f"=IF({i}=0,MAX(land_appraisal-tot_land_draw-tot_landeq,0),0)", note="Always shown (Common Ground)")
    for col, w in zip("ABCDEFGH", [52, 44, 15, 15, 15, 15, 16, 60]): wsB.column_dimensions[col].width = w

    # ======================================================================
    # REVENUE & OPEX (static, year-1 dollars)
    wsR = wb.create_sheet("Revenue & Opex")
    banner(wsR, "Revenue & operating assumptions (year-1 dollars at stabilization)", "Rents are 2024-26 resale listing references (Arterial), not new-build PBR comps; studio rent has NO comp", 8)
    hdr(wsR, 5, ["Item", "Basis"] + [f"Phase {p['ph']}" for p in phases] + ["Total", "Note"])
    row = 6; rrow = {}
    def rrow_(key, label, basis, fn, note="", fmt=FMT_MONEY, total=True, bold=False, fill=None):
        nonlocal row
        put(wsR, row, 1, label, bold=bold, fill=fill); put(wsR, row, 2, basis, fill=fill)
        for i, p in enumerate(phases):
            cc = PC[p["ph"]]; put(wsR, row, cc, fn(i, p, cc), fmt=fmt, bold=bold, fill=fill); dn(f"p{p['ph']}_{key}", "Revenue & Opex", f"${L(cc)}${row}")
        if total:
            put(wsR, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=fmt, bold=True, fill=(fill or TOT_FILL)); dn(f"tot_{key}", "Revenue & Opex", f"${L(TC)}${row}")
        put(wsR, row, TC + 1, note, italic=True, color="8A7C6A"); rrow[key] = row; row += 1
    put(wsR, row, 1, "Rent schedule by suite type", bold=True, fill=SUB_FILL); row += 1
    hdr(wsR, row, ["Suite", "Target sf", "Market rent $/mo", "$/sf/mo", "Comp $/sf/mo", "Affordable rent $/mo (80% MMR)", "Affordable $/sf", "Basis"]); row += 1
    for st in E.SUITE_TYPES:
        t = st["t"]
        put(wsR, row, 1, st["label"]); put(wsR, row, 2, f"=sf_{t}", fmt=FMT_INT)
        put(wsR, row, 3, f"=rent_{t}*(1+newbuild_prem)*(1+rent_shock)", fmt=FMT_MONEY); dn(f"mrent_{t}", "Revenue & Opex", f"$C${row}")
        put(wsR, row, 4, f"=C{row}/B{row}", fmt=FMT_MONEY2)
        put(wsR, row, 5, (f"=rent_{t}/comp_sf_{t}" if t else "n/a"), fmt=FMT_MONEY2)
        put(wsR, row, 6, f"=ROUND(mmr_{t}*aff_factor,0)", fmt=FMT_MONEY); dn(f"arent_{t}", "Revenue & Opex", f"$F${row}")
        put(wsR, row, 7, f"=F{row}/B{row}", fmt=FMT_MONEY2)
        put(wsR, row, 8, ("NO COMP — placeholder" if t == 0 else f"Arterial comp at {v[f'comp_sf_{t}']} sf; concept suite is smaller so $/sf rises"), italic=True, color="8A7C6A")
        row += 1
    row += 1
    put(wsR, row, 1, "Phase revenue (100% occupancy, year-1 $)", bold=True, fill=SUB_FILL); row += 1
    rrow_("gpr_mkt", "Market-rent GPR (annual)", "Σ market homes × rent × 12", lambda i, p, c: "=12*(" + "+".join(f"(p{p['ph']}_n{t}-p{p['ph']}_a{t})*mrent_{t}" for t in range(4)) + ")")
    rrow_("gpr_aff", "Affordable-rent GPR (annual)", "Σ affordable homes × 80% MMR × 12", lambda i, p, c: "=12*(" + "+".join(f"p{p['ph']}_a{t}*arent_{t}" for t in range(4)) + ")", note="CMHC MMR values PENDING")
    rrow_("retail_m", "Retail net rent (monthly, after vacancy)", "retail sf × ret_rent/12 × (1−ret_vac)", lambda i, p, c: f"=p{p['ph']}_retail_sf*ret_rent/12*(1-ret_vac)", note="PLACEHOLDER rent; no retail comps")
    rrow_("pk_m", "Parking income (monthly, 100% occupancy)", "spaces × (1−visitor) × take-up × rent", lambda i, p, c: f"=p{p['ph']}_spaces*(1-pk_visitor)*pk_takeup*pk_rent", note="PLACEHOLDER take-up and rent; no EV income assumed")
    rrow_("egi_s", "Stabilized EGI (annual)", "(GPR × occ × (1−bad debt)) + retail + parking × occ", lambda i, p, c: f"=(p{p['ph']}_gpr_mkt+p{p['ph']}_gpr_aff)*stab_occ*(1-bad_debt)+p{p['ph']}_retail_m*12+p{p['ph']}_pk_m*12*stab_occ", bold=True)
    rrow_("opex_var_s", "Variable opex (utilities, R&M/turnover)", "(util + rm) × homes × occ", lambda i, p, c: f"=-(opex_util+opex_rm)*p{p['ph']}_homes*stab_occ")
    rrow_("opex_fix_s", "Fixed opex (insurance, staffing, admin)", "(ins + staff + ga) × homes", lambda i, p, c: f"=-(opex_ins+opex_staff+opex_ga)*p{p['ph']}_homes")
    rrow_("mgmt_s", "Management", "mgmt_pct × EGI", lambda i, p, c: f"=-p{p['ph']}_egi_s*mgmt_pct")
    rrow_("noi_pre", "NOI before property tax", "", lambda i, p, c: f"=p{p['ph']}_egi_s+p{p['ph']}_opex_var_s+p{p['ph']}_opex_fix_s+p{p['ph']}_mgmt_s", bold=True)
    rrow_("noi_s", "NOI after property tax (closed form)", "NOI_pre ÷ (1 + tax_rate × assess_ratio ÷ cap)", lambda i, p, c: f"=p{p['ph']}_noi_pre/(1+tax_rate*assess_ratio/cap_eff)", bold=True, note="Assessed value = assess_ratio × stabilized value — NOT cost or price (handoff §4)")
    rrow_("tax", "Property tax (annual, year-1 $)", "tax_rate × assess_ratio × NOI ÷ cap", lambda i, p, c: f"=tax_rate*assess_ratio*p{p['ph']}_noi_s/cap_eff", note="0.9823% multi-residential rate UNVERIFIED")
    rrow_("reserve_s", "Replacement reserve (annual)", "reserve × homes — below valuation NOI; deducted for lender DSCR", lambda i, p, c: f"=-reserve*p{p['ph']}_homes")
    rrow_("opex_ratio", "Opex ratio (incl. tax) to EGI", "", lambda i, p, c: f"=-(p{p['ph']}_opex_var_s+p{p['ph']}_opex_fix_s+p{p['ph']}_mgmt_s-p{p['ph']}_tax)/p{p['ph']}_egi_s", fmt=FMT_PCT, total=False)
    rrow_("noi_home", "NOI per home (year-1 $)", "", lambda i, p, c: f"=p{p['ph']}_noi_s/p{p['ph']}_homes", total=False)
    for col, w in zip("ABCDEFGH", [46, 48, 16, 16, 16, 18, 16, 60]): wsR.column_dimensions[col].width = w

    # ======================================================================
    # MONTHLY CASH FLOW
    wsM = wb.create_sheet("Monthly Cash Flow")
    banner(wsM, "Monthly cash flow — phases and consolidated", f"Month 1 = t_start; {H} months (covers Year 20 of the last phase). Sources = uses monthly; MLI interest capitalized in construction; waterfall per Common Ground", 12)
    C0 = 4  # first month column (D)
    def col(mm): return L(C0 + mm - 1)
    lastcol = col(H)
    put(wsM, 4, 1, "Month #", bold=True); put(wsM, 5, 1, "Date", bold=True); put(wsM, 6, 1, "Year #", bold=True)
    for mm in range(1, H + 1):
        c = wsM.cell(row=4, column=C0 + mm - 1, value=mm); c.font = Font(bold=True, size=9); c.fill = SUB_FILL
        d = wsM.cell(row=5, column=C0 + mm - 1, value=f"=EDATE(t_start,{col(mm)}4-1)"); d.number_format = FMT_DATE; d.font = Font(size=9)
        y = wsM.cell(row=6, column=C0 + mm - 1, value=f"=INT(({col(mm)}4-1)/12)+1"); y.font = Font(size=9)
        wsM.column_dimensions[col(mm)].width = 11
    wsM.column_dimensions["A"].width = 44; wsM.column_dimensions["B"].width = 9; wsM.column_dimensions["C"].width = 14
    wsM.freeze_panes = "D7"
    MROW = {}   # (phase, key) -> row
    row = 8
    PENDING = []   # specs awaiting write: (ph, key, label, fn, fmt, bold, fill, total_col)
    def mrow(ph, key, label, fn, fmt=FMT_MONEY, bold=False, fill=None, total_col=True):
        """Register a monthly row; rows are numbered on registration so formulas may reference rows
        registered later in the same block (flush() writes them)."""
        nonlocal row
        MROW[(ph, key)] = row; PENDING.append((row, ph, key, label, fn, fmt, bold, fill, total_col)); row += 1
        return row - 1
    def flush():
        while PENDING:
            rr_, ph, key, label, fn, fmt, bold, fill, total_col = PENDING.pop(0)
            put(wsM, rr_, 1, label, bold=bold, fill=fill); put(wsM, rr_, 2, f"P{ph}" if ph else "ALL", fill=fill)
            for mm in range(1, H + 1):
                c = wsM.cell(row=rr_, column=C0 + mm - 1, value=fn(mm)); c.number_format = fmt; c.font = Font(size=9, bold=bold)
                if fill: c.fill = fill
            if total_col:
                t = wsM.cell(row=rr_, column=3, value=f"=SUM(D{rr_}:{lastcol}{rr_})"); t.number_format = fmt; t.font = Font(size=9, bold=True); t.fill = TOT_FILL
    def R(ph, key, mm): return f"{col(mm)}{MROW[(ph, key)]}"
    def RP(ph, key, mm): return f"{col(mm-1)}{MROW[(ph, key)]}" if mm > 1 else "0"
    def Mc(mm): return f"{col(mm)}$4"
    for p in phases:
        k = p["ph"]; P = f"p{k}_"
        put(wsM, row, 1, f"PHASE {k} — {p['bldgs']} ({p['homes']} homes)", bold=True, fill=HEAD_FILL, color="FFFFFF")
        for cc in range(2, 8): wsM.cell(row=row, column=cc).fill = HEAD_FILL
        row += 1
        mrow(k, "in_constr", "Construction flag", lambda mm: f"=IF(AND({Mc(mm)}>={P}start,{Mc(mm)}<={P}complete),1,0)", fmt=FMT_INT, total_col=False)
        mrow(k, "oper", "Operating flag", lambda mm: f"=IF({Mc(mm)}>{P}complete,1,0)", fmt=FMT_INT, total_col=False)
        mrow(k, "hard_m", "Hard cost", lambda mm: f"=-{R(k,'in_constr',mm)}*{P}hard/{P}dur")
        mrow(k, "soft_m", "Soft cost", lambda mm: f"=-{R(k,'in_constr',mm)}*{P}soft/{P}dur")
        mrow(k, "dc_m", "Development charges", lambda mm: f"=-IF({Mc(mm)}={P}start,{P}dc,0)")
        mrow(k, "land_m", "Landowner draw (+LTT)", lambda mm: f"=-IF({Mc(mm)}={P}start,{P}land,0)")
        mrow(k, "fees_m", "CMHC premium & fees", lambda mm: f"=-IF({Mc(mm)}={P}start,{P}fees,0)")
        mrow(k, "cost_pre", "Cash uses before capitalized interest", lambda mm: f"=-({R(k,'hard_m',mm)}+{R(k,'soft_m',mm)}+{R(k,'dc_m',mm)}+{R(k,'land_m',mm)}+{R(k,'fees_m',mm)})", bold=True)
        mrow(k, "cum_prev", "Cumulative uses, prior month", lambda mm: f"={RP(k,'cum',mm)}", total_col=False)
        mrow(k, "cum", "Cumulative uses", lambda mm: f"={R(k,'cum_prev',mm)}+{R(k,'cost_pre',mm)}", total_col=False)
        mrow(k, "src_city", "Source: City", lambda mm: f"=MAX(0,MIN({R(k,'cum',mm)},{P}city)-MIN({R(k,'cum_prev',mm)},{P}city))")
        mrow(k, "src_grant", "Source: BCH contribution", lambda mm: f"=MAX(0,MIN({R(k,'cum',mm)},{P}city+{P}grant)-MIN({R(k,'cum_prev',mm)},{P}city+{P}grant))-{R(k,'src_city',mm)}")
        mrow(k, "src_mezz", "Source: BCH mezz", lambda mm: f"=MAX(0,MIN({R(k,'cum',mm)},{P}city+{P}grant+{P}mezz)-MIN({R(k,'cum_prev',mm)},{P}city+{P}grant+{P}mezz))-{R(k,'src_city',mm)}-{R(k,'src_grant',mm)}")
        mrow(k, "src_mli_pre", "Source: MLI Select advance (ex interest)", lambda mm: f"={R(k,'cost_pre',mm)}-{R(k,'src_city',mm)}-{R(k,'src_grant',mm)}-{R(k,'src_mezz',mm)}")
        mrow(k, "mli_int_c", "MLI construction interest (capitalized)", lambda mm: f"=IF({R(k,'in_constr',mm)}=1,({RP(k,'cbal',mm)}+{R(k,'src_mli_pre',mm)})*r_m,0)")
        mrow(k, "cbal", "MLI construction balance (pre-conversion)", lambda mm: f"=IF({R(k,'in_constr',mm)}=1,{RP(k,'cbal',mm)}+{R(k,'src_mli_pre',mm)}+{R(k,'mli_int_c',mm)},0)", total_col=False)
        mrow(k, "cost_total", "TOTAL USES (incl. capitalized interest)", lambda mm: f"={R(k,'cost_pre',mm)}+{R(k,'mli_int_c',mm)}", bold=True)
        mrow(k, "src_total", "TOTAL SOURCES", lambda mm: f"={R(k,'src_city',mm)}+{R(k,'src_grant',mm)}+{R(k,'src_mezz',mm)}+{R(k,'src_mli_pre',mm)}+{R(k,'mli_int_c',mm)}", bold=True)
        mrow(k, "mli_int_p", "MLI permanent interest", lambda mm: f"=IF({R(k,'oper',mm)}=1,IF({Mc(mm)}={P}complete+1,{P}perm,{RP(k,'mli_bal',mm)})*r_m,0)")
        mrow(k, "mli_prin", "MLI principal", lambda mm: f"=IF({R(k,'oper',mm)}=1,{P}pmt-{R(k,'mli_int_p',mm)},0)")
        mrow(k, "mli_ds", "MLI debt service", lambda mm: f"=IF({R(k,'oper',mm)}=1,{P}pmt,0)")
        mrow(k, "mli_bal", "MLI balance (end of month)", lambda mm: f"=IF({R(k,'in_constr',mm)}=1,{R(k,'cbal',mm)},IF({R(k,'oper',mm)}=1,IF({Mc(mm)}={P}complete+1,{P}perm,{RP(k,'mli_bal',mm)})-{R(k,'mli_prin',mm)},0))", total_col=False)
        mrow(k, "delivered", "Homes delivered", lambda mm: f"={R(k,'oper',mm)}*{P}homes", fmt=FMT_INT, total_col=False)
        earlier = [q["ph"] for q in phases if q["ph"] < k]
        mrow(k, "occ", "Homes occupied (sitewide leasing capacity, phase priority)",
             lambda mm: f"=IF({R(k,'oper',mm)}=1,MIN({P}homes*stab_occ,{RP(k,'occ',mm)}+MAX(0,cap_m" + "".join(f"-{R(j,'absorb',mm)}" for j in earlier) + ")),0)", fmt='#,##0.0', total_col=False)
        mrow(k, "absorb", "Net absorption", lambda mm: f"={R(k,'occ',mm)}-{RP(k,'occ',mm)}", fmt='#,##0.0')
        mrow(k, "stab_flag", "Stabilized flag", lambda mm: f"=IF(AND({R(k,'oper',mm)}=1,{R(k,'occ',mm)}>={P}homes*stab_occ-0.001),1,0)", fmt=FMT_INT, total_col=False)
        mrow(k, "ym", "Phase-year since stabilization", lambda mm: f"=IF(AND({P}stab>0,{Mc(mm)}>{P}stab),INT(({Mc(mm)}-{P}stab-1)/12)+1,0)", fmt=FMT_INT, total_col=False)
        mrow(k, "g_rent", "Rent growth factor", lambda mm: f"=IF({R(k,'oper',mm)}=1,(1+rent_growth)^INT(({Mc(mm)}-{P}complete-1)/12),0)", fmt='0.0000', total_col=False)
        mrow(k, "g_opex", "Opex growth factor", lambda mm: f"=IF({R(k,'oper',mm)}=1,(1+opex_growth)^INT(({Mc(mm)}-{P}complete-1)/12),0)", fmt='0.0000', total_col=False)
        mrow(k, "rent_mkt", "Market rent", lambda mm: f"={P}gpr_mkt/12*{R(k,'occ',mm)}/{P}homes*{R(k,'g_rent',mm)}")
        mrow(k, "rent_aff", "Affordable rent", lambda mm: f"={P}gpr_aff/12*{R(k,'occ',mm)}/{P}homes*{R(k,'g_rent',mm)}")
        mrow(k, "retail", "Retail net rent", lambda mm: f"=IF({Mc(mm)}>{P}complete+retail_delay,{P}retail_m*{R(k,'g_rent',mm)},0)")
        mrow(k, "parking", "Parking", lambda mm: f"={P}pk_m*{R(k,'occ',mm)}/{P}homes*{R(k,'g_rent',mm)}")
        mrow(k, "bad_debt", "Bad debt / concessions", lambda mm: f"=-({R(k,'rent_mkt',mm)}+{R(k,'rent_aff',mm)})*bad_debt")
        mrow(k, "egi", "Effective gross income", lambda mm: f"={R(k,'rent_mkt',mm)}+{R(k,'rent_aff',mm)}+{R(k,'retail',mm)}+{R(k,'parking',mm)}+{R(k,'bad_debt',mm)}", bold=True)
        mrow(k, "opex_var", "Variable opex (occupied homes)", lambda mm: f"=-(opex_util+opex_rm)/12*{R(k,'occ',mm)}*{R(k,'g_opex',mm)}")
        mrow(k, "opex_fixed", "Fixed opex (delivered homes)", lambda mm: f"=-(opex_ins+opex_staff+opex_ga)/12*{P}homes*{R(k,'g_opex',mm)}")
        mrow(k, "mgmt", "Management", lambda mm: f"=-{R(k,'egi',mm)}*mgmt_pct")
        mrow(k, "tax", "Property tax", lambda mm: f"=-{P}tax/12*{R(k,'g_opex',mm)}")
        mrow(k, "noi", "NET OPERATING INCOME", lambda mm: f"={R(k,'egi',mm)}+{R(k,'opex_var',mm)}+{R(k,'opex_fixed',mm)}+{R(k,'mgmt',mm)}+{R(k,'tax',mm)}", bold=True)
        mrow(k, "reserve", "Replacement reserve", lambda mm: f"=-reserve/12*{P}homes*{R(k,'g_opex',mm)}")
        mrow(k, "sup_ds", "Supplemental charge debt service (Yr 10+)", lambda mm: f"=IF(AND({P}stab>0,{Mc(mm)}>{P}stab+120),{P}sup_pmt,0)")
        mrow(k, "sup_bal", "Supplemental charge balance", lambda mm: f"=IF(AND({P}stab>0,{Mc(mm)}={P}stab+120),{P}sup,IF(AND({P}stab>0,{Mc(mm)}>{P}stab+120),{RP(k,'sup_bal',mm)}-({P}sup_pmt-{RP(k,'sup_bal',mm)}*r_m),0))", total_col=False)
        mrow(k, "ncf", "Net cash flow after debt service", lambda mm: f"={R(k,'noi',mm)}+{R(k,'reserve',mm)}-{R(k,'mli_ds',mm)}-{R(k,'sup_ds',mm)}", bold=True)
        mrow(k, "res_draw", "Lease-up reserve draw", lambda mm: f"=IF(AND({R(k,'oper',mm)}=1,{R(k,'ym',mm)}=0,{R(k,'ncf',mm)}<0),MIN(-{R(k,'ncf',mm)},{RP(k,'res_bal',mm)}),0)")
        mrow(k, "res_bal", "Lease-up reserve balance", lambda mm: f"=IF({Mc(mm)}={P}complete,{P}reserve_lu,{RP(k,'res_bal',mm)}-{R(k,'res_draw',mm)})", total_col=False)
        mrow(k, "buyout", "Landowner Year-10 buyout (equity + MLI paydown)", lambda mm: f"=IF(AND({P}stab>0,{Mc(mm)}={P}stab+120),{P}landeq+({P}perm-{R(k,'mli_bal',mm)}),0)")
        mrow(k, "lo_cf", "Landowner cash-flow share (Yrs 1-10)", lambda mm: f"=IF(AND({R(k,'ym',mm)}>=1,{R(k,'ym',mm)}<=10,{R(k,'ncf',mm)}>0),{R(k,'ncf',mm)}*split_lo,0)")
        mrow(k, "mezz_acc", "BCH mezz accrual flag", lambda mm: f"=IF(AND({R(k,'ym',mm)}>=11,MOD({Mc(mm)}-{P}stab,12)=1),1,0)", fmt=FMT_INT, total_col=False)
        mrow(k, "mezz_pre", "BCH mezz balance before sweep", lambda mm: f"={RP(k,'mezz_bal',mm)}*(1+{R(k,'mezz_acc',mm)}*mezz_rate)+{R(k,'src_mezz',mm)}", total_col=False)
        mrow(k, "sweep", "BCH mezz sweep (Yrs 11-20, 60% seat)", lambda mm: f"=IF(AND({R(k,'ym',mm)}>=11,{R(k,'ym',mm)}<=20),MIN(MAX({R(k,'ncf',mm)}*split_lo,0),{R(k,'mezz_pre',mm)}),0)")
        mrow(k, "mezz_bal", "BCH mezz balance", lambda mm: f"={R(k,'mezz_pre',mm)}-{R(k,'sweep',mm)}", total_col=False)
        mrow(k, "np_cf", "Non-profit net cash flow", lambda mm: f"=IF({R(k,'ym',mm)}=0,{R(k,'ncf',mm)}+{R(k,'res_draw',mm)},IF({R(k,'ym',mm)}<=10,{R(k,'ncf',mm)}-{R(k,'lo_cf',mm)},{R(k,'ncf',mm)}-{R(k,'sweep',mm)}))", bold=True)
        if k == 1:
            mrow(k, "existing", "Existing site NOI retained (entitlement only)", lambda mm: f"=IF({Mc(mm)}<=ent_total,existing_noi*existing_noi_keep/12,0)")
        else:
            mrow(k, "existing", "Existing site NOI retained", lambda mm: "=0")
        mrow(k, "unlev", "Unlevered project cash flow", lambda mm: f"=-{R(k,'cost_total',mm)}+{R(k,'noi',mm)}+{R(k,'reserve',mm)}+{R(k,'existing',mm)}", bold=True)
        mrow(k, "chk_su", "Check: uses − sources", lambda mm: f"={R(k,'cost_total',mm)}-{R(k,'src_total',mm)}")
        flush(); row += 1
    # consolidated
    put(wsM, row, 1, "CONSOLIDATED (interphase transfers: none — each phase funds itself; common works paid once in Phase 1)", bold=True, fill=HEAD_FILL, color="FFFFFF")
    for cc in range(2, 8): wsM.cell(row=row, column=cc).fill = HEAD_FILL
    row += 1
    def csum(key): return lambda mm: "=" + "+".join(R(p["ph"], key, mm) for p in phases)
    for key, label, fmt, bold in [
        ("cost_total", "Total uses", FMT_MONEY, True), ("src_city", "City", FMT_MONEY, False), ("src_grant", "BCH contribution", FMT_MONEY, False),
        ("src_mezz", "BCH mezz", FMT_MONEY, False), ("src_mli_pre", "MLI advances (ex interest)", FMT_MONEY, False), ("mli_int_c", "MLI capitalized interest", FMT_MONEY, False),
        ("src_total", "Total sources", FMT_MONEY, True), ("delivered", "Homes delivered", FMT_INT, False), ("occ", "Homes occupied", '#,##0.0', False),
        ("egi", "Effective gross income", FMT_MONEY, False), ("noi", "NET OPERATING INCOME", FMT_MONEY, True), ("reserve", "Replacement reserve", FMT_MONEY, False),
        ("mli_ds", "MLI debt service", FMT_MONEY, False), ("sup_ds", "Supplemental debt service", FMT_MONEY, False), ("ncf", "Net cash flow after debt service", FMT_MONEY, True),
        ("lo_cf", "Landowner share", FMT_MONEY, False), ("buyout", "Landowner buyout", FMT_MONEY, False), ("sweep", "BCH mezz sweep", FMT_MONEY, False),
        ("np_cf", "Non-profit net cash flow", FMT_MONEY, True), ("mli_bal", "MLI balance", FMT_MONEY, False), ("mezz_bal", "BCH mezz balance", FMT_MONEY, False),
        ("sup_bal", "Supplemental balance", FMT_MONEY, False), ("existing", "Existing NOI retained", FMT_MONEY, False), ("unlev", "Unlevered project cash flow", FMT_MONEY, True),
        ("chk_su", "Check: uses − sources (must be 0)", FMT_MONEY, False)]:
        mrow(0, key, label, csum(key), fmt=fmt, bold=bold, total_col=(key not in ("mli_bal", "mezz_bal", "sup_bal", "occ", "delivered")))
    mrow(0, "land_cash", "Landowner cash received (draw + share + buyout)", lambda mm: "=" + "+".join(f"-{R(p['ph'],'land_m',mm)}+{R(p['ph'],'lo_cf',mm)}+{R(p['ph'],'buyout',mm)}" for p in phases))
    mrow(0, "cum_public", "Cumulative public non-repayable + mezz funded", lambda mm: f"={RP(0,'cum_public',mm)}+{R(0,'src_city',mm)}+{R(0,'src_grant',mm)}+{R(0,'src_mezz',mm)}", total_col=False)
    mrow(0, "unlev_term", "Unlevered CF incl. terminal value at horizon", lambda mm: f"={R(0,'unlev',mm)}+IF({Mc(mm)}=horizon,SUM({col(H-11)}{MROW[(0,'noi')]}:{col(H)}{MROW[(0,'noi')]})/cap_eff,0)", bold=True)
    flush()
    m_next = row   # next free row on the Monthly sheet (row counter is shared across tabs)
    # names for month/date rows and key rows
    dn("m_row", "Monthly Cash Flow", f"$D$4:${lastcol}$4"); dn("d_row", "Monthly Cash Flow", f"$D$5:${lastcol}$5")
    for (ph, key), rr_ in MROW.items():
        dn(f"{'c' if ph == 0 else 'p'+str(ph)}_r_{key}", "Monthly Cash Flow", f"$D${rr_}:${lastcol}${rr_}")
    # fix Schedule stab formulas now that rows exist
    for p in phases:
        cc = PC[p["ph"]]; rr_ = MROW[(p["ph"], "stab_flag")]
        wsS.cell(row=stab_row, column=cc, value=f"=IFERROR(MATCH(1,'Monthly Cash Flow'!$D${rr_}:${lastcol}${rr_},0),0)")
    wsS.cell(row=stab_row, column=TC, value="First month occupancy ≥ homes × stab_occ (sitewide leasing capacity, phase priority)")

    # ======================================================================
    # DEBT
    wsD = wb.create_sheet("Debt")
    banner(wsD, "Debt & funding (Common Ground)", "MLI Select construction-to-perm sized on TDC; permanent loan = balance at completion; supplemental charge funds the Year-10 buyout", 8)
    hdr(wsD, 5, ["Item", "Basis"] + [f"Phase {p['ph']}" for p in phases] + ["Total", "Note"])
    row = 6
    def drow_(key, label, basis, fn, note="", fmt=FMT_MONEY, total=True, bold=False, name=True):
        nonlocal row
        put(wsD, row, 1, label, bold=bold); put(wsD, row, 2, basis)
        for i, p in enumerate(phases):
            cc = PC[p["ph"]]; put(wsD, row, cc, fn(i, p, cc), fmt=fmt, bold=bold)
            if name: dn(f"p{p['ph']}_{key}", "Debt", f"${L(cc)}${row}")
        if total:
            put(wsD, row, TC, f"=SUM(C{row}:{L(TC-1)}{row})", fmt=fmt, bold=True, fill=TOT_FILL)
            if name: dn(f"tot_{key}", "Debt", f"${L(TC)}${row}")
        put(wsD, row, TC + 1, note, italic=True, color="8A7C6A"); row += 1
    drow_("mli_plan", "MLI Select (plan, from Budget)", "cmhc_share × TDC", lambda i, p, c: f"=p{p['ph']}_mli")
    drow_("perm", "MLI permanent loan = construction balance at completion", "INDEX(construction balance row, complete)", lambda i, p, c: f"=INDEX(p{p['ph']}_r_cbal,1,p{p['ph']}_complete)", note="Includes actual capitalized interest; construction balance row has no dependency on the permanent payment (no circularity)")
    drow_("pmt", "Monthly payment", "PMT(r_m, n_am, perm)", lambda i, p, c: f"=-PMT(r_m,n_am,p{p['ph']}_perm)", fmt=FMT_MONEY2)
    drow_("ads", "Annual debt service", "", lambda i, p, c: f"=p{p['ph']}_pmt*12")
    drow_("perm_vs_plan", "Permanent − plan (interest approximation)", "", lambda i, p, c: f"=p{p['ph']}_perm-p{p['ph']}_mli", note="Small; reconciled on Checks (tolerance 2%)")
    drow_("sup", "Supplemental charge (Year-10 buyout × (1+prem))", "Σ buyout × (1 + buyout_prem)", lambda i, p, c: f"=SUM(p{p['ph']}_r_buyout)*(1+buyout_prem)")
    drow_("sup_pmt", "Supplemental monthly payment", "PMT(r_m, n_am, sup)", lambda i, p, c: f"=-PMT(r_m,n_am,p{p['ph']}_sup)", fmt=FMT_MONEY2)
    drow_("city_f", "City funded", "Σ", lambda i, p, c: f"=SUM(p{p['ph']}_r_src_city)")
    drow_("grant_f", "BCH contribution funded", "Σ", lambda i, p, c: f"=SUM(p{p['ph']}_r_src_grant)")
    drow_("mezz_f", "BCH mezz funded", "Σ", lambda i, p, c: f"=SUM(p{p['ph']}_r_src_mezz)")
    drow_("mli_f", "MLI advances incl. capitalized interest", "Σ", lambda i, p, c: f"=SUM(p{p['ph']}_r_src_mli_pre)+SUM(p{p['ph']}_r_mli_int_c)")
    drow_("capint_act", "Capitalized interest (actual)", "Σ", lambda i, p, c: f"=SUM(p{p['ph']}_r_mli_int_c)")
    drow_("int_total", "Total MLI interest over horizon", "Σ construction + permanent", lambda i, p, c: f"=SUM(p{p['ph']}_r_mli_int_c)+SUM(p{p['ph']}_r_mli_int_p)")
    drow_("mli_peak", "Peak MLI balance", "MAX", lambda i, p, c: f"=MAX(p{p['ph']}_r_mli_bal)")
    drow_("sweep_t", "BCH mezz swept (Yrs 11-20)", "Σ", lambda i, p, c: f"=SUM(p{p['ph']}_r_sweep)")
    drow_("mezz_y20", "BCH mezz balance at Year 20", "INDEX(balance, stab+240)", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,INDEX(p{p['ph']}_r_mezz_bal,1,p{p['ph']}_y20_m),0)")
    row += 1
    put(wsD, row, 1, "Funding order each month: City → BCH contribution → BCH mezz → MLI Select advances (construction interest capitalized). Equity-first vs pari-passu is not a lever under Common Ground because there is no private equity; the landowner's 5% is in-kind.", italic=True, color="8A7C6A"); row += 1
    put(wsD, row, 1, "Permanent debt is NOT sized to DSCR/LTV — Common Ground fixes MLI at 50% of TDC. The DSCR, LTV and debt-yield that result are reported on Phase Returns and tested on Checks; a failing DSCR means the stack does not close at these inputs.", italic=True, color="8A7C6A"); row += 1
    for col_, w in zip("ABCDEFGH", [48, 40, 16, 16, 16, 16, 17, 60]): wsD.column_dimensions[col_].width = w

    # ======================================================================
    # PHASE RETURNS
    wsX = wb.create_sheet("Phase Returns")
    banner(wsX, "Phase & consolidated returns", "Stabilized NOI = 12 months after stabilization; DSCR on lender NOI (after reserve); IRR/NPV dated (XIRR/XNPV)", 8)
    hdr(wsX, 5, ["Metric", "Basis"] + [f"Phase {p['ph']}" for p in phases] + ["Consolidated", "Note"])
    row = 6
    def xrow_(key, label, basis, fn, cons=None, note="", fmt=FMT_MONEY, bold=False, name=True):
        nonlocal row
        put(wsX, row, 1, label, bold=bold); put(wsX, row, 2, basis)
        for i, p in enumerate(phases):
            cc = PC[p["ph"]]; put(wsX, row, cc, fn(i, p, cc), fmt=fmt, bold=bold)
            if name: dn(f"p{p['ph']}_{key}", "Phase Returns", f"${L(cc)}${row}")
        if cons is not None:
            put(wsX, row, TC, (f"=SUM(C{row}:{L(TC-1)}{row})" if cons == "sum" else cons), fmt=fmt, bold=True, fill=TOT_FILL)
            if name: dn(f"tot_{key}", "Phase Returns", f"${L(TC)}${row}")
        put(wsX, row, TC + 1, note, italic=True, color="8A7C6A"); row += 1
    def sec_x(t):
        nonlocal row
        put(wsX, row, 1, t, bold=True, fill=SUB_FILL)
        for cc in range(2, TC + 2): wsX.cell(row=row, column=cc).fill = SUB_FILL
        row += 1
    sec_x("Programme")
    xrow_("homes_x", "Homes", "", lambda i, p, c: f"=p{p['ph']}_homes", cons="sum", fmt=FMT_INT)
    xrow_("aff_x", "Affordable homes (80% MMR)", "", lambda i, p, c: f"=p{p['ph']}_aff", cons="sum", fmt=FMT_INT)
    xrow_("nra_x", "Suite rentable area (sf)", "", lambda i, p, c: f"=p{p['ph']}_nra", cons="sum", fmt=FMT_INT)
    xrow_("ag_x", "Above-ground gross (sf)", "", lambda i, p, c: f"=p{p['ph']}_above_sf", cons="sum", fmt=FMT_INT)
    xrow_("bg_x", "Below-ground gross (sf)", "", lambda i, p, c: f"=p{p['ph']}_below_sf", cons="sum", fmt=FMT_INT)
    xrow_("pk_x", "Parking spaces", "", lambda i, p, c: f"=p{p['ph']}_spaces", cons="sum", fmt=FMT_INT)
    sec_x("Cost")
    xrow_("tdc_x", "Total development cost", "", lambda i, p, c: f"=p{p['ph']}_tdc", cons="sum", bold=True)
    xrow_("tdc_home_x", "Cost per home", "", lambda i, p, c: f"=p{p['ph']}_tdc/p{p['ph']}_homes", cons=f"={L(TC)}{row-1}/tot_homes")
    xrow_("tdc_sf_x", "Cost per gross sf (above + below)", "", lambda i, p, c: f"=p{p['ph']}_tdc/(p{p['ph']}_above_sf+p{p['ph']}_below_sf)", cons=f"={L(TC)}{row-2}/(tot_above_sf+tot_below_sf)", fmt=FMT_MONEY2)
    xrow_("hard_sf_x", "Hard cost per above-ground gross sf", "", lambda i, p, c: f"=p{p['ph']}_hard/p{p['ph']}_above_sf", cons="=tot_hard/tot_above_sf", fmt=FMT_MONEY2)
    sec_x("Operations at stabilization")
    xrow_("stab_x", "Stabilization month #", "", lambda i, p, c: f"=p{p['ph']}_stab", cons="=stab_all", fmt=FMT_INT)
    xrow_("noi", "Stabilized NOI (12 months after stabilization)", "SUMIFS on Monthly Cash Flow", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,SUMIFS(p{p['ph']}_r_noi,m_row,\">\"&p{p['ph']}_stab,m_row,\"<=\"&p{p['ph']}_stab+12),0)", cons="sum", bold=True)
    xrow_("res12", "Replacement reserve (same 12 months)", "", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,SUMIFS(p{p['ph']}_r_reserve,m_row,\">\"&p{p['ph']}_stab,m_row,\"<=\"&p{p['ph']}_stab+12),0)", cons="sum")
    xrow_("noi_l", "Lender NOI (after reserve)", "", lambda i, p, c: f"=p{p['ph']}_noi+p{p['ph']}_res12", cons="sum")
    xrow_("occ_x", "Stabilized occupancy", "", lambda i, p, c: "=stab_occ", cons="=stab_occ", fmt=FMT_PCT)
    xrow_("avg_rent", "Average market rent / home / month (year-1 $)", "", lambda i, p, c: f"=p{p['ph']}_gpr_mkt/12/(p{p['ph']}_homes-p{p['ph']}_aff)", cons="=tot_gpr_mkt/12/(tot_homes-tot_aff)")
    xrow_("avg_arent", "Average affordable rent / home / month", "", lambda i, p, c: f"=p{p['ph']}_gpr_aff/12/p{p['ph']}_aff", cons="=tot_gpr_aff/12/tot_aff")
    xrow_("yoc", "Unlevered yield on cost (NOI ÷ TDC incl. financing & land equity)", "", lambda i, p, c: f"=p{p['ph']}_noi/p{p['ph']}_tdc", cons="=tot_noi/tot_tdc_x", fmt=FMT_PCT, bold=True)
    xrow_("yoc_exfin", "Yield on cost ex financing & land equity", "NOI ÷ (TDC − cap. interest − fees − land equity)", lambda i, p, c: f"=p{p['ph']}_noi/(p{p['ph']}_tdc-p{p['ph']}_capint-p{p['ph']}_fees-p{p['ph']}_landeq)", cons="=tot_noi/(tot_tdc_x-tot_capint-tot_fees-tot_landeq)", fmt=FMT_PCT)
    xrow_("cap_x", "Cap rate", "", lambda i, p, c: "=cap_eff", cons="=cap_eff", fmt=FMT_PCT)
    xrow_("value", "Stabilized value (NOI ÷ cap)", "", lambda i, p, c: f"=p{p['ph']}_noi/cap_eff", cons="sum", bold=True)
    xrow_("margin", "Development margin on cost ((value × (1−sell) − TDC) ÷ TDC)", "", lambda i, p, c: f"=(p{p['ph']}_value*(1-sell_cost)-p{p['ph']}_tdc)/p{p['ph']}_tdc", cons="=(tot_value*(1-sell_cost)-tot_tdc_x)/tot_tdc_x", fmt=FMT_PCT)
    sec_x("Coverage (Common Ground tests)")
    xrow_("dscr", "DSCR at stabilization (lender NOI ÷ MLI debt service)", "", lambda i, p, c: f"=IF(p{p['ph']}_ads>0,p{p['ph']}_noi_l/p{p['ph']}_ads,0)", cons="=IF(tot_ads>0,tot_noi_l/tot_ads,0)", fmt=FMT_X, bold=True, note="Test 1: ≥ min_dscr")
    xrow_("ltv", "LTV (permanent MLI ÷ value)", "", lambda i, p, c: f"=p{p['ph']}_perm/p{p['ph']}_value", cons="=tot_perm/tot_value", fmt=FMT_PCT)
    xrow_("dy", "Debt yield (lender NOI ÷ permanent MLI)", "", lambda i, p, c: f"=p{p['ph']}_noi_l/p{p['ph']}_perm", cons="=tot_noi_l/tot_perm", fmt=FMT_PCT)
    xrow_("noi10", "Year-10 lender NOI (months stab+121..132)", "", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,SUMIFS(p{p['ph']}_r_noi,m_row,\">\"&(p{p['ph']}_stab+120),m_row,\"<=\"&(p{p['ph']}_stab+132))+SUMIFS(p{p['ph']}_r_reserve,m_row,\">\"&(p{p['ph']}_stab+120),m_row,\"<=\"&(p{p['ph']}_stab+132)),0)", cons="sum")
    xrow_("ds10", "Year-10 debt service (MLI + supplemental)", "", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,SUMIFS(p{p['ph']}_r_mli_ds,m_row,\">\"&(p{p['ph']}_stab+120),m_row,\"<=\"&(p{p['ph']}_stab+132))+SUMIFS(p{p['ph']}_r_sup_ds,m_row,\">\"&(p{p['ph']}_stab+120),m_row,\"<=\"&(p{p['ph']}_stab+132)),0)", cons="sum")
    xrow_("dscr10", "Year-10 DSCR incl. supplemental", "", lambda i, p, c: f"=IF(p{p['ph']}_ds10>0,p{p['ph']}_noi10/p{p['ph']}_ds10,0)", cons="=IF(tot_ds10>0,tot_noi10/tot_ds10,0)", fmt=FMT_X, bold=True, note="Test 2: ≥ 1.10 each phase")
    xrow_("np_min", "Minimum non-profit annual cash flow, phase-years 1-20", "MIN of annual SUMIFS by phase-year", lambda i, p, c: f"=MIN(" + ",".join(f"SUMIFS(p{p['ph']}_r_np_cf,p{p['ph']}_r_ym,{y})" for y in range(1, 21)) + ")", cons=f"=MIN(C{row}:{L(TC-1)}{row})", bold=True, note="Test 3: ≥ 0")
    xrow_("value20", "Year-20 value (forward NOI ÷ cap)", "", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,SUMIFS(p{p['ph']}_r_noi,m_row,\">\"&p{p['ph']}_y20_m,m_row,\"<=\"&p{p['ph']}_y20_m+12)/cap_eff,0)", cons="sum")
    xrow_("refi20", "Year-20 takeout at takeout_ltv", "", lambda i, p, c: f"=p{p['ph']}_value20*takeout_ltv", cons="sum")
    xrow_("debt20", "Year-20 debt to clear (MLI + supplemental + mezz)", "", lambda i, p, c: f"=IF(p{p['ph']}_stab>0,INDEX(p{p['ph']}_r_mli_bal,1,p{p['ph']}_y20_m)+INDEX(p{p['ph']}_r_sup_bal,1,p{p['ph']}_y20_m)+INDEX(p{p['ph']}_r_mezz_bal,1,p{p['ph']}_y20_m),0)", cons="sum")
    xrow_("takeout_ok", "Takeout clears all debt?", "", lambda i, p, c: f"=IF(p{p['ph']}_refi20>=p{p['ph']}_debt20,\"PASS\",\"FAIL\")", cons=f"=IF(COUNTIF(C{row}:{L(TC-1)}{row},\"FAIL\")=0,\"PASS\",\"FAIL\")", fmt="@", bold=True, note="Test 4")
    sec_x("Who gets what (nominal, over horizon)")
    xrow_("lo_draw", "Landowner Day-1 draw", "", lambda i, p, c: f"=p{p['ph']}_land_draw", cons="sum")
    xrow_("lo_share", "Landowner cash-flow share, Yrs 1-10", "", lambda i, p, c: f"=SUM(p{p['ph']}_r_lo_cf)", cons="sum")
    xrow_("lo_bo", "Landowner Year-10 buyout", "", lambda i, p, c: f"=SUM(p{p['ph']}_r_buyout)", cons="sum")
    xrow_("lo_total", "Landowner all-in by Year 10", "", lambda i, p, c: f"=p{p['ph']}_lo_draw+p{p['ph']}_lo_share+p{p['ph']}_lo_bo", cons="sum", bold=True, note="Compare with $17M ask (Dashboard)")
    xrow_("lo_eq", "Landowner retained equity (in-kind)", "", lambda i, p, c: f"=p{p['ph']}_landeq", cons="sum")
    xrow_("np_total", "Non-profit net cash flow (horizon)", "", lambda i, p, c: f"=SUM(p{p['ph']}_r_np_cf)", cons="sum")
    xrow_("bch_sw", "BCH mezz recovered (Yrs 11-20)", "", lambda i, p, c: f"=p{p['ph']}_sweep_t", cons="sum")
    xrow_("bch_mz20", "BCH mezz outstanding at Year 20 (before takeout)", "", lambda i, p, c: f"=p{p['ph']}_mezz_y20", cons="sum")
    xrow_("grant_x", "BCH contribution (non-repayable if affordability vests)", "", lambda i, p, c: f"=p{p['ph']}_grant", cons="sum")
    xrow_("city_x", "City contribution", "", lambda i, p, c: f"=p{p['ph']}_city", cons="sum")
    sec_x("Exposure & returns")
    xrow_("peak_pub", "Peak cumulative public funding (City + BCH)", "", lambda i, p, c: f"=p{p['ph']}_city_f+p{p['ph']}_grant_f+p{p['ph']}_mezz_f", cons="=MAX(c_r_cum_public)", bold=True, note="Consolidated = peak of monthly cumulative; phases never overlap-fund under sequential")
    xrow_("peak_mli", "Peak MLI balance", "", lambda i, p, c: f"=p{p['ph']}_mli_peak", cons="=MAX(c_r_mli_bal)")
    xrow_("int_x", "Total MLI interest over horizon", "", lambda i, p, c: f"=p{p['ph']}_int_total", cons="sum")
    xrow_("irr", "Unlevered project IRR (dated, terminal value at horizon)", "XIRR", lambda i, p, c: f"=IFERROR(XIRR(p{p['ph']}_r_unlev_t,d_row),\"n/a\")", cons="=IFERROR(XIRR(c_r_unlev_term,d_row),\"n/a\")", fmt=FMT_PCT, bold=True)
    xrow_("npv", "Dated NPV at disc_rate", "XNPV", lambda i, p, c: f"=XNPV(disc_rate,p{p['ph']}_r_unlev_t,d_row)", cons="=XNPV(disc_rate,c_r_unlev_term,d_row)")
    xrow_("first_occ_x", "First occupancy", "", lambda i, p, c: f"=p{p['ph']}_first_occ_d", cons="=MIN(C{0}:{1}{0})".format(row, L(TC-1)), fmt=FMT_DATE)
    xrow_("stab_d_x", "Stabilization", "", lambda i, p, c: f"=p{p['ph']}_stab_d", cons=f"=EDATE(t_start,stab_all-1)", fmt=FMT_DATE)
    xrow_("hold_end_x", "Hold end (final stabilization + hold)", "", lambda i, p, c: "=EDATE(t_start,hold_end-1)", cons="=EDATE(t_start,hold_end-1)", fmt=FMT_DATE)
    # per-phase unlevered with terminal value rows on Monthly sheet (needed for XIRR)
    row_save = row; row = m_next
    for p in phases:
        k = p["ph"]
        mrow(k, "unlev_t", f"P{k} unlevered CF incl. terminal at horizon", lambda mm, k=k: f"={R(k,'unlev',mm)}+IF({Mc(mm)}=horizon,SUM({col(H-11)}{MROW[(k,'noi')]}:{col(H)}{MROW[(k,'noi')]})/cap_eff,0)")
        dn(f"p{k}_r_unlev_t", "Monthly Cash Flow", f"$D${MROW[(k,'unlev_t')]}:${lastcol}${MROW[(k,'unlev_t')]}")
    flush(); m_next = row; row = row_save
    for col_, w in zip("ABCDEFGH", [58, 34, 16, 16, 16, 16, 18, 50]): wsX.column_dimensions[col_].width = w

    # ======================================================================
    # EXIT & RESIDUAL
    wsE = wb.create_sheet("Exit & Residual")
    banner(wsE, "Exit tests & residual land value", "Common Ground residual = largest appraisal passing all four tests (engine-solved). Conventional residuals shown as labelled cross-checks", 8)
    row = 5
    put(wsE, row, 1, "A. Stabilize-and-sell test (each phase sold at its stabilization)", bold=True, fill=SUB_FILL); row += 1
    hdr(wsE, row, ["Item", "Basis"] + [f"Phase {p['ph']}" for p in phases] + ["Total", "Note"]); row += 1
    def erow_(key, label, basis, fn, cons="sum", note="", fmt=FMT_MONEY, bold=False):
        nonlocal row
        put(wsE, row, 1, label, bold=bold); put(wsE, row, 2, basis)
        for i, p in enumerate(phases):
            cc = PC[p["ph"]]; put(wsE, row, cc, fn(i, p, cc), fmt=fmt, bold=bold); dn(f"p{p['ph']}_{key}", "Exit & Residual", f"${L(cc)}${row}")
        if cons is not None:
            put(wsE, row, TC, (f"=SUM(C{row}:{L(TC-1)}{row})" if cons == "sum" else cons), fmt=fmt, bold=True, fill=TOT_FILL); dn(f"tot_{key}", "Exit & Residual", f"${L(TC)}${row}")
        put(wsE, row, TC + 1, note, italic=True, color="8A7C6A"); row += 1
    erow_("sale_v", "Sale price at stabilization (NOI ÷ cap)", "", lambda i, p, c: f"=p{p['ph']}_value")
    erow_("sale_c", "Selling costs", "sell_cost × price", lambda i, p, c: f"=-p{p['ph']}_sale_v*sell_cost")
    erow_("sale_mli", "MLI balance discharged at stabilization", "INDEX(balance, stab)", lambda i, p, c: f"=-IF(p{p['ph']}_stab>0,INDEX(p{p['ph']}_r_mli_bal,1,p{p['ph']}_stab),0)")
    erow_("sale_mezz", "BCH mezz discharged", "", lambda i, p, c: f"=-IF(p{p['ph']}_stab>0,INDEX(p{p['ph']}_r_mezz_bal,1,p{p['ph']}_stab),0)")
    erow_("sale_claw", "BCH contribution clawback (unvested at sale)", "contribution × (1 − years vested/10)", lambda i, p, c: f"=-p{p['ph']}_grant*(1-MIN(1,0/10))", note="Sale at stabilization = 0 years vested → 100% clawback")
    erow_("sale_lo", "Landowner buyout on sale (retained equity)", "", lambda i, p, c: f"=-p{p['ph']}_landeq")
    erow_("sale_net", "Net proceeds to non-profit", "", lambda i, p, c: f"=p{p['ph']}_sale_v+p{p['ph']}_sale_c+p{p['ph']}_sale_mli+p{p['ph']}_sale_mezz+p{p['ph']}_sale_claw+p{p['ph']}_sale_lo", bold=True)
    erow_("sale_profit", "Development profit vs TDC (value × (1−sell) − TDC)", "", lambda i, p, c: f"=p{p['ph']}_sale_v*(1-sell_cost)-p{p['ph']}_tdc", bold=True)
    erow_("sale_margin", "Margin on cost", "", lambda i, p, c: f"=p{p['ph']}_sale_profit/p{p['ph']}_tdc", cons="=tot_sale_profit/tot_tdc_x", fmt=FMT_PCT)
    row += 1
    put(wsE, row, 1, "B. Single sale after all phases stabilize (month stab_all)", bold=True, fill=SUB_FILL); row += 1
    single = [
        ("single_noi", "Forward NOI, 12 months after stab_all", "=SUMIFS(c_r_noi,m_row,\">\"&stab_all,m_row,\"<=\"&stab_all+12)"),
        ("single_v", "Sale price (NOI ÷ cap)", "=single_noi/cap_eff"),
        ("single_c", "Selling costs", "=-single_v*sell_cost"),
        ("single_debt", "Debt discharged (MLI + supplemental + mezz at stab_all)", "=-(INDEX(c_r_mli_bal,1,stab_all)+INDEX(c_r_sup_bal,1,stab_all)+INDEX(c_r_mezz_bal,1,stab_all))"),
        ("single_claw", "BCH contribution clawback (unvested; Phase 1 partly vested)", "=-(" + "+".join(f"p{p['ph']}_grant*(1-MIN(1,MAX(0,(stab_all-p{p['ph']}_stab)/120)))" for p in phases) + ")"),
        ("single_lo", "Landowner retained equity settled", "=-tot_landeq"),
        ("single_net", "Net proceeds to non-profit", "=single_v+single_c+single_debt+single_claw+single_lo"),
        ("single_profit", "Development profit vs TDC", "=single_v*(1-sell_cost)-tot_tdc_x"),
        ("single_margin", "Margin on cost", "=single_profit/tot_tdc_x"),
        ("single_vs_phased", "Single sale net proceeds − phased sale net proceeds", "=single_net-tot_sale_net"),
    ]
    for key, lab, f in single:
        put(wsE, row, 1, lab); c = put(wsE, row, 3, f, fmt=(FMT_PCT if "margin" in key else FMT_MONEY), bold=key in ("single_net", "single_profit")); dn(key, "Exit & Residual", f"$C${row}"); row += 1
    put(wsE, row, 1, "Common infrastructure is carried in Phase 1 TDC (infra_p1_share); per-phase profit therefore understates Phase 1 and overstates later phases — the consolidated line is the fair comparison.", italic=True, color="8A7C6A"); row += 2
    put(wsE, row, 1, "C. Hold (Common Ground base case) — position at hold end", bold=True, fill=SUB_FILL); row += 1
    hold = [
        ("hold_noi", "Forward NOI at hold end", "=SUMIFS(c_r_noi,m_row,\">\"&hold_end,m_row,\"<=\"&hold_end+12)"),
        ("hold_v", "Value at hold end", "=hold_noi/cap_eff"),
        ("hold_debt", "Debt at hold end (MLI + supplemental + mezz)", "=INDEX(c_r_mli_bal,1,hold_end)+INDEX(c_r_sup_bal,1,hold_end)+INDEX(c_r_mezz_bal,1,hold_end)"),
        ("hold_eq", "Non-profit equity at hold end", "=hold_v-hold_debt"),
        ("hold_np_cum", "Cumulative non-profit cash flow to hold end", "=SUMIFS(c_r_np_cf,m_row,\"<=\"&hold_end)"),
        ("hold_lo_cum", "Cumulative landowner cash to hold end", "=SUMIFS(c_r_land_cash,m_row,\"<=\"&hold_end)"),
    ]
    for key, lab, f in hold:
        put(wsE, row, 1, lab); put(wsE, row, 3, f, fmt=FMT_MONEY, bold=key == "hold_eq"); dn(key, "Exit & Residual", f"$C${row}"); row += 1
    row += 1
    put(wsE, row, 1, "D. Residual land value", bold=True, fill=SUB_FILL); row += 1
    hdr(wsE, row, ["Measure", "Hurdle / test", "Value", "", "", "", "", "Note"]); row += 1
    resid_cg, _ = E.solve_land(v)
    be_hard17 = E.solve_input(v, "hard_psf", 0, v["hard_psf"], increasing=True)
    be_hard0 = E.solve_input(dict(v, land_appraisal=0.0), "hard_psf", 0, v["hard_psf"], increasing=True)
    be_rent = E.solve_input(v, "rent_shock", 0, 3.0, increasing=False)
    be_dc = E.solve_input(v, "dc_per_home", 0, v["dc_per_home"], increasing=True)
    # appraisal at which landowner all-in = ask (engine)
    def lo_total_at(x):
        _, mm = E.evaluate(dict(v, land_appraisal=x)); return mm["lo_total"]
    engine_rows = [
        ("Common Ground supportable appraisal", "All four tests pass (DSCR ≥ min_dscr; DSCR10 ≥ 1.10; NFP ≥ 0; Year-20 takeout)", resid_cg, FMT_MONEY, "Engine-solved by bisection on land_appraisal. $0 = fails at any land value"),
        ("Surplus / (shortfall) vs $17.0M ask", "", resid_cg - v["land_appraisal"], FMT_MONEY, ""),
        ("Break-even above-ground hard cost at $17.0M appraisal", "$/sf at which all tests pass", be_hard17, FMT_MONEY2, f"vs benchmark ${v['hard_psf']}/sf"),
        ("Break-even above-ground hard cost at $0 appraisal", "$/sf", be_hard0, FMT_MONEY2, ""),
        ("Break-even starting-rent uplift at benchmark cost", "rent_shock at which all tests pass", be_rent, FMT_PCT, "Uniform uplift on market rents"),
        ("Break-even development charge per home (at $17.0M)", "$/home at which all tests pass", be_dc, FMT_MONEY, "0 or blank = DCs alone cannot fix it"),
        ("Landowner all-in by Year 10 at $17.0M appraisal (engine)", "draw + share + buyout", lo_total_at(v["land_appraisal"]), FMT_MONEY, "Reconciles to Phase Returns (live)"),
    ]
    for lab, hurdle, val, fmt, note in engine_rows:
        put(wsE, row, 1, lab, bold=True); put(wsE, row, 2, hurdle)
        c = put(wsE, row, 3, (val if val is not None else "n/a"), fmt=fmt, fill=ENGINE_FILL)
        put(wsE, row, 8, note, italic=True, color="8A7C6A"); row += 1
    row += 1
    put(wsE, row, 1, "Conventional cross-checks (live formulas, labelled hurdles — not a house standard)", bold=True); row += 1
    conv = [
        ("cost_ex_land", "Cost excluding land recognition (TDC − draw − LTT − land equity)", "=tot_tdc_x-tot_land-tot_landeq", FMT_MONEY),
        ("land_yoc", "Land supported at target_yoc (NOI ÷ target_yoc − cost ex land)", "=tot_noi/target_yoc-cost_ex_land", FMT_MONEY),
        ("land_margin", "Land supported at target_margin (value × (1−sell) ÷ (1+margin) − cost ex land)", "=tot_value*(1-sell_cost)/(1+target_margin)-cost_ex_land", FMT_MONEY),
        ("land_binding", "Binding conventional residual", "=MIN(land_yoc,land_margin)", FMT_MONEY),
        ("land_simple", "Simple value-minus-cost residual (value × (1−sell) − cost ex land − target_margin × TDC)", "=tot_value*(1-sell_cost)-cost_ex_land-target_margin*tot_tdc_x", FMT_MONEY),
        ("lo_vs_ask", "Landowner all-in by Year 10 − $17.0M ask (live)", "=tot_lo_total-land_appraisal", FMT_MONEY),
        ("donated_x", "Appraised value donated to affordability", "=tot_donated", FMT_MONEY),
    ]
    for key, lab, f, fmt in conv:
        put(wsE, row, 1, lab); put(wsE, row, 3, f, fmt=fmt); dn(key, "Exit & Residual", f"$C${row}"); row += 1
    row += 1
    put(wsE, row, 1, "E. Approved 172-townhouse plan — for-sale cross-check (what the $17M ask rests on)", bold=True, fill=SUB_FILL); row += 1
    put(wsE, row, 1, "Inputs below are placeholders: resale comps are 50-year-old stock at ~$410/sf (Arterial, 8 listings, Centennial, Jun-Jul 2026); Katanna's new-build pricing is VIP-only. Hard cost $250/sf (Arterial, 3-storey stacked townhouse $230-270). DC per townhouse and soft/finance ratios are analyst placeholders.", italic=True, color="8A7C6A"); row += 1
    th = [
        ("th_units", "Units (draft plan; marketing now says 165)", 172, FMT_INT, True),
        ("th_sf", "Average sellable sf per unit", 1300, FMT_INT, True),
        ("th_psf", "Sale price $/sf (new build)", 550, FMT_MONEY, True),
        ("th_gfa_sf", "Gross floor area, sf (supplied 313,966)", 313966, FMT_INT, True),
        ("th_hard", "Hard cost $/sf GFA (Arterial 3-storey stacked, midpoint)", 250, FMT_MONEY, True),
        ("th_dc", "DC + education per townhouse (placeholder)", 65000, FMT_MONEY, True),
        ("th_soft", "Soft costs % of hard", 0.18, FMT_PCT, True),
        ("th_fin", "Financing & carry % of hard+soft", 0.07, FMT_PCT, True),
        ("th_sell", "Selling costs % of revenue", 0.04, FMT_PCT, True),
        ("th_margin", "Developer margin % of revenue (hurdle)", 0.15, FMT_PCT, True),
        ("th_rev", "Revenue", "=th_units*th_sf*th_psf", FMT_MONEY, False),
        ("th_cost", "Costs ex land", "=th_gfa_sf*th_hard*(1+th_soft)*(1+th_fin)+th_units*th_dc", FMT_MONEY, False),
        ("th_resid", "Residual land value (revenue × (1 − sell − margin) − costs)", "=th_rev*(1-th_sell-th_margin)-th_cost", FMT_MONEY, False),
        ("th_resid_500", "  at $500/sf", "=th_units*th_sf*500*(1-th_sell-th_margin)-th_cost", FMT_MONEY, False),
        ("th_resid_450", "  at $450/sf", "=th_units*th_sf*450*(1-th_sell-th_margin)-th_cost", FMT_MONEY, False),
        ("th_resid_600", "  at $600/sf", "=th_units*th_sf*600*(1-th_sell-th_margin)-th_cost", FMT_MONEY, False),
    ]
    for key, lab, val, fmt, is_in in th:
        put(wsE, row, 1, lab); put(wsE, row, 3, val, fmt=fmt, fill=(FLAG_FILL if is_in else None), bold=key.startswith("th_resid")); dn(key, "Exit & Residual", f"$C${row}"); row += 1
    for col_, w in zip("ABCDEFGH", [70, 52, 18, 16, 16, 16, 18, 60]): wsE.column_dimensions[col_].width = w

    # ======================================================================
    # SENSITIVITIES (engine values)
    wsT = wb.create_sheet("Sensitivities")
    banner(wsT, "Sensitivities & stress cases", "ENGINE-COMPUTED VALUES (lilac). Analyst stress instructions per handoff §6, not forecasts. Re-run build_workbook.py after changing Inputs. Base row reconciles to live formulas (Checks)", 12)
    rows = E.scenario_table(v)
    cols = ["Scenario", "Homes", "TDC", "TDC/home", "Stab. NOI", "YoC", "DSCR", "Min DSCR10", "NFP min yr", "Takeout", "Landowner all-in", "BCH mezz @Y20", "Peak public", "Stab month", "Unlev IRR", "PASS?"]
    hdr(wsT, 5, cols); row = 6
    def scen_row(label, mm, vv, fill=ENGINE_FILL):
        nonlocal row
        vals_ = [label, mm["homes"], mm["tdc"], mm["tdc"] / mm["homes"], mm["noi_stab"], mm["yoc"], mm["dscr"], mm["dscr10_min"], mm["np_cf_min_year"],
                 ("PASS" if mm["takeout_ok"] else "FAIL"), mm["lo_total"], mm["mezz_bal20"], mm["peak_public"], mm["stab_all"], mm["unlev_irr"], ("PASS" if E.passes(mm, vv) else "FAIL")]
        fmts = [None, FMT_INT, FMT_MONEY, FMT_MONEY, FMT_MONEY, FMT_PCT, FMT_X, FMT_X, FMT_MONEY, "@", FMT_MONEY, FMT_MONEY, FMT_MONEY, FMT_INT, FMT_PCT, "@"]
        for i, (x, f) in enumerate(zip(vals_, fmts)):
            put(wsT, row, 1 + i, x, fmt=f, fill=fill, bold=(i == 0))
        row += 1
    put(wsT, row, 1, "Single-variable and combined stress (R01 1,000 homes)", bold=True, fill=SUB_FILL); row += 1
    for label, ov, mm in rows:
        vv = dict(v); vv.update(ov); scen_row(label, mm, vv)
    row += 1
    put(wsT, row, 1, "Density / residual comparison (600/800 are analyst re-builds, 400 is a sensitivity — not designs or verified zoning yields)", bold=True, fill=SUB_FILL); row += 1
    dens = E.density_table(v)
    hdr(wsT, row, ["Programme", "Homes", "TDC", "TDC/home", "Stab. NOI", "YoC", "DSCR", "CG residual appraisal", "Conv. residual (binding hurdle)", "Note"]); row += 1
    for d in dens:
        mm = d["m"]
        vals_ = [d["meta"]["label"], mm["homes"], mm["tdc"], mm["tdc"] / mm["homes"], mm["noi_stab"], mm["yoc"], mm["dscr"], d["resid_cg"], d["conv"]["land_binding"], d["meta"]["note"]]
        fmts = [None, FMT_INT, FMT_MONEY, FMT_MONEY, FMT_MONEY, FMT_PCT, FMT_X, FMT_MONEY, FMT_MONEY, None]
        for i, (x, f) in enumerate(zip(vals_, fmts)): put(wsT, row, 1 + i, x, fmt=f, fill=ENGINE_FILL, bold=(i == 0))
        row += 1
    row += 1
    # two-way tables
    def grid(title, xa, xs, xl, ya, ys, yl, fn, fmt):
        nonlocal row
        put(wsT, row, 1, title, bold=True, fill=SUB_FILL); row += 1
        put(wsT, row, 1, f"{yl} ↓ / {xl} →", bold=True)
        for j, x in enumerate(xs): put(wsT, row, 2 + j, x, fmt=(FMT_PCT if isinstance(x, float) and abs(x) < 1 else FMT_MONEY), bold=True, fill=SUB_FILL)
        row += 1
        g = E.two_way(v, xa, xs, ya, ys, fn)
        for i, y in enumerate(ys):
            put(wsT, row, 1, y, fmt=(FMT_PCT if isinstance(y, float) and abs(y) < 1 else FMT_INT), bold=True, fill=SUB_FILL)
            for j, x in enumerate(xs): put(wsT, row, 2 + j, g[i][j], fmt=fmt, fill=ENGINE_FILL)
            row += 1
        row += 1
    grid("Consolidated DSCR at stabilization — rent shock (rows) vs hard-cost shock (columns)", "hard_shock", [-0.10, 0.0, 0.15], "hard cost", "rent_shock", [-0.10, 0.0, 0.10], "rent", lambda mm, vv: mm["dscr"], FMT_X)
    grid("Consolidated stabilized value — rent shock (rows) vs cap-rate shock (columns)", "cap_shock", [-0.005, 0.0, 0.01], "cap shock", "rent_shock", [-0.10, 0.0, 0.10], "rent", lambda mm, vv: mm["value"], FMT_MONEY)
    grid("Landowner all-in by Year 10 — appraisal (rows) vs approval delay months (columns)", "entitle_delay", [0, 12, 24], "delay", "land_appraisal", [10e6, 17e6, 25e6], "appraisal", lambda mm, vv: mm["lo_total"], FMT_MONEY)
    grid("Consolidated DSCR — appraisal (rows) vs approval delay (columns)", "entitle_delay", [0, 12, 24], "delay", "land_appraisal", [10e6, 17e6, 25e6], "appraisal", lambda mm, vv: mm["dscr"], FMT_X)
    grid("Consolidated YoC — hard $/sf (columns: $290 / $340 / $390 benchmark range)", "hard_psf", [290, 340, 390], "hard $/sf", "rent_shock", [0.0], "rent", lambda mm, vv: mm["yoc"], FMT_PCT)
    for i, w in enumerate([46, 12, 16, 14, 16, 10, 10, 12, 14, 10, 18, 16, 16, 12, 12, 8]): wsT.column_dimensions[L(1 + i)].width = w

    # ======================================================================
    # DASHBOARD
    wsDash = wb.create_sheet("Dashboard")
    banner(wsDash, "Investment dashboard — Common Ground, R01 1,000 homes", "Live formulas; scenario summary on Sensitivities", 6)
    row = 5
    dash = [
        ("PROGRAMME", None, None),
        ("Homes / affordable homes", "=tot_homes&\" / \"&tot_aff", "@"),
        ("Suite rentable area (sf)", "=tot_nra", FMT_INT),
        ("Above / below-ground gross (sf)", "=TEXT(tot_above_sf,\"#,##0\")&\" / \"&TEXT(tot_below_sf,\"#,##0\")", "@"),
        ("Parking spaces", "=tot_spaces", FMT_INT),
        ("COST", None, None),
        ("Total development cost", "=tot_tdc_x", FMT_MONEY),
        ("Cost per home", "=tot_tdc_x/tot_homes", FMT_MONEY),
        ("Hard cost per above-ground gross sf", "=tot_hard/tot_above_sf", FMT_MONEY2),
        ("Development charges", "=tot_dc", FMT_MONEY),
        ("Land recognized (draw + retained equity) / donated", "=TEXT(tot_land+tot_landeq,\"$#,##0\")&\" / \"&TEXT(tot_donated,\"$#,##0\")", "@"),
        ("OPERATIONS", None, None),
        ("Stabilized NOI (consolidated)", "=tot_noi", FMT_MONEY),
        ("NOI per home", "=tot_noi/tot_homes", FMT_MONEY),
        ("Unlevered yield on cost", "=tot_yoc", FMT_PCT),
        ("Cap rate / stabilized value", "=TEXT(cap_eff,\"0.00%\")&\" / \"&TEXT(tot_value,\"$#,##0\")", "@"),
        ("Development margin on cost", "=tot_margin", FMT_PCT),
        ("COVERAGE — COMMON GROUND TESTS", None, None),
        ("DSCR at stabilization (test ≥ min_dscr)", "=tot_dscr", FMT_X),
        ("Minimum phase Year-10 DSCR (test ≥ 1.10)", "=MIN(" + ",".join(f"p{p['ph']}_dscr10" for p in phases) + ")", FMT_X),
        ("Minimum non-profit annual cash flow (test ≥ 0)", "=tot_np_min", FMT_MONEY),
        ("Year-20 takeout clears debt (test)", "=tot_takeout_ok", "@"),
        ("ALL TESTS PASS?", "=IF(AND(tot_dscr>=min_dscr,MIN(" + ",".join(f"p{p['ph']}_dscr10" for p in phases) + ")>=1.1,tot_np_min>=0,tot_takeout_ok=\"PASS\"),\"PASS\",\"FAIL\")", "@"),
        ("LTV / debt yield on permanent MLI", "=TEXT(tot_ltv,\"0.0%\")&\" / \"&TEXT(tot_dy,\"0.00%\")", "@"),
        ("LAND", None, None),
        ("Appraisal tested (seller ask, placeholder)", "=land_appraisal", FMT_MONEY),
        ("Common Ground supportable appraisal (engine)", resid_cg, FMT_MONEY),
        ("Surplus / (shortfall) vs ask (engine)", resid_cg - v["land_appraisal"], FMT_MONEY),
        ("Landowner all-in by Year 10 (draw + share + buyout)", "=tot_lo_total", FMT_MONEY),
        ("Landowner all-in − ask", "=lo_vs_ask", FMT_MONEY),
        ("Conventional residual at labelled hurdles (binding)", "=land_binding", FMT_MONEY),
        ("Break-even hard $/sf at ask (engine)", be_hard17, FMT_MONEY2),
        ("Break-even rent uplift at benchmark cost (engine)", be_rent, FMT_PCT),
        ("EXPOSURE", None, None),
        ("Peak cumulative public funding (City + BCH)", "=tot_peak_pub", FMT_MONEY),
        ("Peak MLI balance", "=tot_peak_mli", FMT_MONEY),
        ("Total MLI interest over horizon", "=tot_int_x", FMT_MONEY),
        ("BCH mezz outstanding at Year 20 (before takeout)", "=tot_bch_mz20", FMT_MONEY),
        ("Unlevered project IRR / dated NPV", "=IFERROR(TEXT(tot_irr,\"0.00%\"),\"n/a\")&\" / \"&TEXT(tot_npv,\"$#,##0\")", "@"),
        ("DATES", None, None),
        ("Financial close (t_start)", "=t_start", FMT_DATE),
        ("Phase 1 construction start", "=p1_start_d", FMT_DATE),
        ("First occupancy (Phase 1)", "=p1_first_occ_d", FMT_DATE),
        ("All phases stabilized", "=EDATE(t_start,stab_all-1)", FMT_DATE),
        ("Hold end", "=EDATE(t_start,hold_end-1)", FMT_DATE),
    ]
    for lab, f, fmt in dash:
        if f is None:
            put(wsDash, row, 1, lab, bold=True, fill=SUB_FILL); wsDash.cell(row=row, column=2).fill = SUB_FILL; row += 1; continue
        put(wsDash, row, 1, lab)
        is_engine = not (isinstance(f, str) and f.startswith("="))
        put(wsDash, row, 2, f, fmt=fmt, bold=True, fill=(ENGINE_FILL if is_engine else None)); row += 1
    row += 1
    put(wsDash, row, 1, "Phase summary (live)", bold=True, fill=SUB_FILL); row += 1
    hdr(wsDash, row, ["Phase", "Homes", "TDC", "TDC/home", "Stab. NOI", "YoC", "DSCR", "DSCR10", "Value", "Landowner all-in", "Stabilization"]); row += 1
    for p in phases:
        k = p["ph"]
        for i, (f, fmt) in enumerate([(f"=\"Phase {k}\"", "@"), (f"=p{k}_homes", FMT_INT), (f"=p{k}_tdc", FMT_MONEY), (f"=p{k}_tdc/p{k}_homes", FMT_MONEY), (f"=p{k}_noi", FMT_MONEY), (f"=p{k}_yoc", FMT_PCT), (f"=p{k}_dscr", FMT_X), (f"=p{k}_dscr10", FMT_X), (f"=p{k}_value", FMT_MONEY), (f"=p{k}_lo_total", FMT_MONEY), (f"=p{k}_stab_d", FMT_DATE)]):
            put(wsDash, row, 1 + i, f, fmt=fmt)
        row += 1
    wsDash.column_dimensions["A"].width = 56; wsDash.column_dimensions["B"].width = 26
    for i in range(3, 12): wsDash.column_dimensions[L(i)].width = 16

    # ======================================================================
    # CHECKS
    wsC = wb.create_sheet("Checks")
    banner(wsC, "Model checks", "All must read TRUE", 6)
    hdr(wsC, 5, ["Check", "Result", "Detail"]); row = 6
    checks = [
        ("Phase homes sum to 1,000", "=tot_homes=1000", "=tot_homes"),
        ("Suite types sum to 1,000", "=SUM(tot_n0,tot_n1,tot_n2,tot_n3)=1000", "=SUM(tot_n0,tot_n1,tot_n2,tot_n3)"),
        ("Suite area sums to 681,000 sf", "=ROUND(tot_nra,0)=681000", "=tot_nra"),
        ("Above-ground area sums to 86,208 m²", "=tot_above=86208", "=tot_above"),
        ("Residential gross sums to 83,408 m²", "=tot_resi=83408", "=tot_resi"),
        ("Other allowances sum to 2,800 m²", "=tot_other=2800", "=tot_other"),
        ("Parking sums to 650 spaces at R01 ratio", "=OR(ROUND(pk_ratio,4)<>0.65,tot_spaces=650)", "=tot_spaces"),
        ("Below-grade sums to 20,800 m² at R01 ratio", "=OR(ROUND(pk_ratio,4)<>0.65,tot_below=20800)", "=tot_below"),
        ("Sources = uses every month (all phases)", "=ABS(SUM(c_r_chk_su))<1", "=SUM(c_r_chk_su)"),
        ("Capital stack = TDC each phase", "=ABS(" + "+".join(f"p{p['ph']}_stack_chk" for p in phases) + ")<1", "=" + "+".join(f"p{p['ph']}_stack_chk" for p in phases)),
        ("Permanent MLI within 2% of plan (interest approximation)", "=AND(" + ",".join(f"ABS(p{p['ph']}_perm_vs_plan)/p{p['ph']}_mli<0.02" for p in phases) + ")", "=MAX(" + ",".join(f"ABS(p{p['ph']}_perm_vs_plan)/p{p['ph']}_mli" for p in phases) + ")"),
        ("MLI balances roll (no negative balance)", "=MIN(c_r_mli_bal)>=-1", "=MIN(c_r_mli_bal)"),
        ("Mezz balances roll (no negative balance)", "=MIN(c_r_mezz_bal)>=-1", "=MIN(c_r_mezz_bal)"),
        ("No rent before occupancy", "=SUMPRODUCT(--(c_r_occ=0),c_r_egi)=0", "=SUMPRODUCT(--(c_r_occ=0),c_r_egi)"),
        ("No existing income after entitlement", "=SUMIFS(p1_r_existing,m_row,\">\"&ent_total)=0", "=SUMIFS(p1_r_existing,m_row,\">\"&ent_total)"),
        ("Each phase stabilizes within horizon", "=MIN(" + ",".join(f"p{p['ph']}_stab" for p in phases) + ")>0", "=stab_all"),
        ("Year 20 of last phase within horizon", "=stab_all+252<=horizon", "=stab_all+252"),
        ("Common works counted once (Phase 1 shares sum to 1)", "=ABS(tot_site-site_m2*site_net_m2)<1", "=tot_site"),
        ("Non-profit never carries an unfunded lease-up deficit beyond reserve", "=MIN(" + ",".join(f"MIN(p{p['ph']}_r_res_bal)" for p in phases) + ")>=-1", "min reserve balance"),
        ("Landowner buyout occurs exactly once per phase", "=AND(" + ",".join(f"COUNTIF(p{p['ph']}_r_buyout,\">0\")=IF(p{p['ph']}_stab>0,1,0)" for p in phases) + ")", ""),
        ("Scenario levers at baseline (rent/hard/cap/rate shocks = 0, delay = 0, pace = 1)", "=AND(rent_shock=0,hard_shock=0,cap_shock=0,rate_shock=0,entitle_delay=0,lease_cap_mult=1,later_defer=0)", "Reset these to reproduce the engine base case"),
        ("Engine base case reconciles: TDC within 0.1%", f"=ABS(tot_tdc_x-{m['tdc']})/{m['tdc']}<0.001", f"engine TDC {m['tdc']:,.0f}"),
        ("Engine base case reconciles: stabilized NOI within 0.5%", f"=ABS(tot_noi-{m['noi_stab']})/{m['noi_stab']}<0.005", f"engine NOI {m['noi_stab']:,.0f}"),
        ("Engine base case reconciles: DSCR within 0.01x", f"=ABS(tot_dscr-{m['dscr']})<0.01", f"engine DSCR {m['dscr']:.3f}"),
        ("Engine base case reconciles: landowner all-in within 1%", f"=ABS(tot_lo_total-{m['lo_total']})/{m['lo_total']}<0.01", f"engine {m['lo_total']:,.0f}"),
        ("Common Ground tests: DSCR ≥ min_dscr", "=tot_dscr>=min_dscr", "=tot_dscr"),
        ("Common Ground tests: every phase DSCR10 ≥ 1.10", "=MIN(" + ",".join(f"p{p['ph']}_dscr10" for p in phases) + ")>=1.1", "=MIN(" + ",".join(f"p{p['ph']}_dscr10" for p in phases) + ")"),
        ("Common Ground tests: non-profit annual CF ≥ 0", "=tot_np_min>=0", "=tot_np_min"),
        ("Common Ground tests: Year-20 takeout clears debt", "=tot_takeout_ok=\"PASS\"", "=tot_refi20-tot_debt20"),
    ]
    for lab, f, det in checks:
        put(wsC, row, 1, lab); put(wsC, row, 2, f, bold=True, align="center")
        put(wsC, row, 3, det, fmt=FMT_MONEY if isinstance(det, str) and det.startswith("=") else None, italic=True, color="8A7C6A"); row += 1
    put(wsC, row + 1, 1, "ALL MODEL-INTEGRITY CHECKS PASS (rows 6-25)", bold=True); put(wsC, row + 1, 2, f"=AND(B6:B{row-5})", bold=True, align="center")
    put(wsC, row + 2, 1, "COMMON GROUND FEASIBILITY (last four rows) — expected FAIL at benchmark inputs", bold=True); put(wsC, row + 2, 2, f"=AND(B{row-4}:B{row-1})", bold=True, align="center")
    wsC.column_dimensions["A"].width = 70; wsC.column_dimensions["B"].width = 12; wsC.column_dimensions["C"].width = 40

    # ======================================================================
    # SOURCES
    wsSrc = wb.create_sheet("Sources", 1)
    banner(wsSrc, "Source register", "Where every external input came from, what was reachable, and what remains PENDING", 8)
    hdr(wsSrc, 5, ["#", "Source", "What it supplied", "Date", "Reached?", "Confidence", "Used for"]); row = 6
    sources = [
        ("City of Oshawa, Economic & Development Services — Item ED-24-20, Attachment 2, Proposed Draft Plan of Subdivision, 1279 Simcoe St N (Katanna Simcoe Ltd.), files S-O-2022-05 / Z-2022-12 / C-O-2022-08", "Block 1 = 3.1211 ha, 172 condo townhouses (traditional, back-to-back, live-work); Block 2 road widening 5.5 m × 87.4 m; assembly of PINs 16285-0006/-0009/-0153/-0151; Lot 27 (PIN 16285-0008) and Lot 24 excluded", "Jan 2024", "Yes (uploaded PDF)", "H", "site_net_m2; approved-plan cross-check"),
        ("Financial Modelling Handoff (internal), 15 Sep 2026", "R01 programme (86,208 m² above ground, 83,408 m² residential gross, 2,800 m² other, 650 spaces, 20,800 m² below grade), suite mix, $17M ask, 7.85 ac, ~$500k existing NOI, 170-unit / 313,966 sf townhouse benchmark, $102,653/home DC placeholder, 0.9823% tax rate, $290-390/sf benchmark", "2026-09-15", "Yes", "M (measured) / L (commercial)", "Programme; land_appraisal; dc_per_home; scenario framework"),
        ("Arterial — resolve_site", "Parcel 17,014 m² (ONE PIN of the assembly, 72.8 m × 276.6 m), 5 existing buildings 6,070 m² footprint, Regional Road 2 (Simcoe St N, 5 lanes)", "2026-09-15", "Yes", "M", "exist_footprint; site note"),
        ("Arterial — financial_context (constructionCost)", "Condos up to 12 storeys $290-390/sf (dataset 8fd5356d); 3-storey stacked townhouse $230-270/sf", "2026-09-15", "Yes", "L (scope not inspected)", "hard_psf; townhouse cross-check"),
        ("Arterial — financial_context (propertyTax)", "0.9823% multi-residential, Oshawa 2025", "2026-09-15", "Yes", "L (unverified)", "tax_rate"),
        ("Arterial — financial_context (developmentCharges, cmhcRents)", "NO Oshawa development-charge dataset; CMHC zone 'Oshawa (North)' returned no rent values", "2026-09-15", "Yes (empty)", "—", "DC and MMR remain placeholders"),
        ("Arterial — run_development_scenario (Optimal Townhouse, 172 units, rental)", "3-BR comps avg $2,412 (4 comps; 53 Taunton Rd E $2,400-2,450 at ~1,100-1,300 sf; 97 Nonquon Rd $2,399); 1-BR $1,795 (1 comp); 2-BR $2,173 (2 comps); residual land value NEGATIVE as rental", "2026-09-15", "Yes", "L", "rent_1/2/3; comp sizes"),
        ("Arterial — find_unit_comps (sales, condo townhouse, Centennial)", "8 active listings Jun-Jul 2026: 3-BR ~1,100 sf $430k-475k (~$411/sf), 2-BR ~950 sf $330k-399k (~$388/sf); 50-year-old stock", "2026-09-15", "Yes", "L (asking, not closed)", "Townhouse for-sale cross-check"),
        ("Arterial — run_hbu_development_scenario", "Job started three times; result never returned within the session ('computing') and the result URL is blocked from this environment", "2026-09-15", "No", "—", "Not used"),
        ("Arterial — analyze_zoning (via scenario policy context)", "'No policy text or zoning data covering this parcel was found' — Oshawa By-law 60-94 not in Arterial", "2026-09-15", "Yes (empty)", "—", "Entitlement treated as rezoning required"),
        ("CMHC Rental Market Survey, Oshawa CMA, Oct 2025 (HMIP portal)", "Median market rents by bedroom", "2025-10", "NO — domain blocked from this session", "—", "mmr_0..3 are PENDING placeholders"),
        ("City of Oshawa development charges page / By-law 104-2025; Region of Durham DC page; DRHBA updates", "Current per-unit DC schedule (City, Region, education)", "2025", "NO — domains blocked from this session", "—", "dc_per_home stays at supplied placeholder"),
        ("Durham Post, 'Oshawa council to consider 172-unit townhouse project'; Oshawa council information memo (escribe DocumentId 14949)", "Approval status of Z-2022-12 / S-O-2022-05", "2024", "NO — domains blocked; search snippet only", "—", "Approval status UNCONFIRMED"),
        ("Platinum Condo Deals / LoopNet / Zolo listings for 1279 Simcoe St N", "Marketing now shows 165 units in 20 blocks (Katanna Developments); commercial units D/F/G for lease at the existing plaza", "2026", "Search snippets only", "L", "Existing-use note; unit-count revision"),
        ("Ubuntu Common Ground skill (July 2026 structure) — generate_package.py", "Capital stack, splits, Day-1 draw, Year-10 buyout, Year-20 takeout, opex defaults, cap/mortgage/vacancy defaults", "2026-07", "Yes", "H (programme rules)", "All Common Ground mechanics"),
    ]
    for i, s in enumerate(sources, 1):
        put(wsSrc, row, 1, i);
        for j, t in enumerate(s): c = put(wsSrc, row, 2 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
        wsSrc.row_dimensions[row].height = 48; row += 1
    for col_, w in zip("ABCDEFG", [4, 60, 70, 12, 22, 14, 34]): wsSrc.column_dimensions[col_].width = w

    # sheet order
    order = ["Read Me", "Sources", "Inputs", "Programme", "Schedule", "Development Budget", "Revenue & Opex", "Debt", "Monthly Cash Flow", "Phase Returns", "Exit & Residual", "Sensitivities", "Dashboard", "Checks"]
    wb._sheets = [wb[n] for n in order]
    for wsx in wb.worksheets:
        wsx.sheet_properties.tabColor = GOLD if wsx.title in ("Dashboard", "Checks") else ESPRESSO
    path = os.path.join(out_dir, "1279_Simcoe_CommonGround_Model.xlsx")
    wb.save(path)
    return path, m, dict(resid_cg=resid_cg, be_hard17=be_hard17, be_hard0=be_hard0, be_rent=be_rent, be_dc=be_dc, scen=rows, dens=dens)

# --------------------------------------------------------------------------
def recalc_and_verify(path, m):
    """Recalculate in LibreOffice headless and read back key cells."""
    outdir = os.path.join(os.path.dirname(path), "_recalc")
    os.makedirs(outdir, exist_ok=True)
    cmd = ["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir, path]
    pr = subprocess.run(cmd, capture_output=True, timeout=900, text=True); print(pr.stdout[-400:], pr.stderr[-400:])
    rp = os.path.join(outdir, os.path.basename(path))
    wb = load_workbook(rp, data_only=True)
    wsC = wb["Checks"]
    results = []
    for rrow in wsC.iter_rows(min_row=6, max_row=wsC.max_row):
        a, b, c = rrow[0].value, rrow[1].value, rrow[2].value
        if a: results.append((a, b, c))
    wsD = wb["Dashboard"]
    dash = {r[0].value: r[1].value for r in wsD.iter_rows(min_row=5, max_row=60) if r[0].value}
    return results, dash, rp

if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    ap.add_argument("--no-verify", action="store_true")
    a = ap.parse_args()
    out = os.path.abspath(a.out)
    path, m, extra = build(out)
    print("wrote", path)
    json.dump({k: (v if not isinstance(v, list) else None) for k, v in extra.items()}, open(os.path.join(out, "model", "engine_summary.json"), "w"), indent=1, default=str)
    if not a.no_verify:
        results, dash, rp = recalc_and_verify(path, m)
        bad = [r for r in results if r[1] not in (True, "TRUE", 1)]
        print("checks:", len(results), "failing:", len(bad))
        for r in results: print("  ", "OK " if r[1] in (True, "TRUE", 1) else "XX ", r[0], "|", r[1], "|", r[2])
        print("dashboard:"); [print("  ", k, "=", v) for k, v in dash.items()]
