#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render the findings memo (HTML + Markdown), the source register (Markdown) and the
package index from the engine — no typed numbers. Run after build_workbook.py."""
import os, json, datetime as dt, html
import cg_engine as E

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
inp = E.default_inputs(); v = E.vals(inp)
r, m = E.evaluate(v)
summ = json.load(open(os.path.join(OUT, "model", "engine_summary.json")))
scen = E.scenario_table(v)
dens = E.density_table(v)
conv = E.solve_land_conventional(v)
detail = E.test_detail(m, v)
be_dc = summ.get("be_dc")

def M(x, d=1):
    if x is None: return "n/a"
    return f"(${-x/1e6:,.{d}f}M)" if x < 0 else f"${x/1e6:,.{d}f}M"
def money(x): return "n/a" if x is None else (f"(${-x:,.0f})" if x < 0 else f"${x:,.0f}")
def pct(x, d=1): return "n/a" if x is None else f"{x*100:.{d}f}%"
def xx(x): return "n/a" if x is None else f"{x:.2f}x"
def dte(mo):
    y = 2027 + (mo - 1) // 12; mm = (mo - 1) % 12 + 1
    return dt.date(y, mm, 1).strftime("%b %Y")

flagged = [(k, d) for k, d in inp.items() if d["flag"]]
n_inputs = len(inp); n_flag = len(flagged)
p1, p4 = m["phases"][0], m["phases"][-1]
per_home = m["tdc"] / m["homes"]
noi_home = m["noi_stab"] / m["homes"]
be_hard17 = summ["be_hard17"]; be_hard0 = summ["be_hard0"]; be_rent = summ["be_rent"]; resid_cg = summ["resid_cg"]
comb = [s for s in scen if s[0] == "Combined downside"][0][2]
overlap = [s for s in scen if s[0].startswith("Overlapping")][0][2]
dcex = [s for s in scen if s[0].startswith("Affordable units DC-exempt")][0][2]
rentup = [s for s in scen if s[0] == "Rents +10%"][0][2]
hardm = [s for s in scen if s[0] == "Hard cost −10%"][0][2]
d1000 = dens[0]; d600 = [d for d in dens if d["case"] == "600"][0]; d400 = [d for d in dens if d["case"] == "400"][0]

# ---- townhouse for-sale cross-check (same placeholders as workbook section E)
th_units, th_sf, th_gfa, th_hard, th_dc, th_soft, th_fin, th_sell, th_margin = 172, 1300, 313966, 250, 65000, 0.18, 0.07, 0.04, 0.15
th_cost = th_gfa * th_hard * (1 + th_soft) * (1 + th_fin) + th_units * th_dc
def th_resid(psf): return th_units * th_sf * psf * (1 - th_sell - th_margin) - th_cost
th = {psf: th_resid(psf) for psf in (450, 500, 550, 600)}
th_be = th_cost / (th_units * th_sf * (1 - th_sell - th_margin))   # sale $/sf at which land residual = 0

# ======================================================================
findings = [
    ("The 1,000-home concept does not close under Common Ground at the handoff's benchmark inputs, at any land value.",
     f"Total development cost is {M(m['tdc'])} ({money(per_home)} per home) against stabilized NOI of {M(m['noi_stab'],2)} ({money(noi_home)} per home). "
     f"Yield on cost is {pct(m['yoc'],2)} against a {pct(v['cap_rate'],2)} cap rate. With MLI Select at only {pct(v['cmhc_share'],0)} of cost the consolidated DSCR is {xx(m['dscr'])} "
     f"(test {xx(v['min_dscr'])}), the weakest phase Year-10 DSCR is {xx(m['dscr10_min'])}, the non-profit's worst post-stabilization year is {M(m['np_cf_min_year'],2)}, "
     f"and the Year-20 takeout is short by {M(m['refi20']-m['debt20'])}. Setting the appraisal to zero does not change the verdict: the supportable Common Ground appraisal is {money(resid_cg)}."),
    ("What has to be true for the stack to close at a $17.0M appraisal.",
     f"Above-ground hard cost of about {money(be_hard17)}/sf (benchmark {money(v['hard_psf'])}/sf, a {pct(1-be_hard17/v['hard_psf'],0)} reduction), or starting rents about {pct(be_rent,0)} above the resale comps, "
     f"or a combination. DC relief alone cannot fix it: with the affordable units exempt the DSCR only moves to {xx(dcex['dscr'])}"
     + (f"; the break-even DC per home is {money(be_dc)}" if be_dc else "; no DC level between $0 and the placeholder makes all four tests pass") + "."),
    ("The $17.0M ask can only be a for-sale townhouse number, and even that is not demonstrated.",
     f"On placeholder costs (Arterial $250/sf on the supplied 313,966 sf GFA, $65k DC per unit, 18% soft, 7% finance, 4% selling, 15% margin) the approved 172-unit condo townhouse plan needs new-build pricing of about {money(th_be)}/sf before it supports any land value: "
     f"{M(th[450])} at $450/sf, {M(th[500])} at $500/sf, {M(th[550])} at $550/sf, {M(th[600])} at $600/sf. Resale comps for 50-year-old stock near the site sit around $410/sf, and Katanna's new-build pricing is VIP-only. "
     f"As a rental, Arterial's own underwrite of the townhouse plan returned a negative residual. The ask therefore rests on the vendor's for-sale pro forma (lower cost basis or higher pricing than these placeholders), which has not been supplied; Common Ground would in any case be paying a for-sale price for rental land."),
    ("Density does not rescue it.",
     f"Cost per home barely moves with scale ({money(d1000['m']['tdc']/d1000['m']['homes'])} at 1,000 homes, {money(d600['m']['tdc']/d600['m']['homes'])} at 600, {money(d400['m']['tdc']/d400['m']['homes'])} at 400) because the benchmark is a per-square-foot rate and the suite mix is fixed. "
     f"Every density case fails all four tests and every conventional residual is negative (600 homes: {M(d600['conv']['land_binding'])}; 400 homes: {M(d400['conv']['land_binding'])}). Fewer homes lowers the absolute shortfall, not the verdict."),
    ("Timing helps at the margin; the combined downside is severe.",
     f"Overlapping phases at 12-month intervals cut TDC to {M(overlap['tdc'])} (less escalation) and lift DSCR to {xx(overlap['dscr'])}. The combined downside (rents −10%, hard +15%, cap +100 bps, approvals +12 months, leasing −50%) "
     f"takes TDC to {M(comb['tdc'])}, DSCR to {xx(comb['dscr'])} and BCH mezz outstanding at Year 20 to {M(comb['mezz_bal20'])}."),
    ("Exposure if it were built anyway.",
     f"Peak cumulative public funding (City + BCH contribution + BCH mezz) is {M(m['peak_public'])}; peak MLI Select balance {M(m['peak_mli'])}; BCH mezz still outstanding at Year 20 {M(m['mezz_bal20'])}. "
     f"The landowner's all-in by Year 10 would be {M(m['lo_total'])} on a $17.0M appraisal (draw {M(m['land_draw'])} + share {M(m['lo_cf'])} + buyout {M(m['lo_buyout'])}), which is why the landowner test is not the binding one: the building cannot carry the debt that funds it."),
]

next_steps = [
    "Obtain a QS Class D estimate on R01 (above-ground, parking, site works, contingency) to replace the $340/sf benchmark; the verdict turns on this line.",
    "Obtain CMHC MMR (Oct 2025) and City/Region DC schedules; both are placeholders. Confirm the affordable-unit DC exemption and the enhanced GST rental rebate with tax counsel.",
    "Confirm the approval status of Z-2022-12 / S-O-2022-05 and whether the 165-unit revision is what is on file. If the seller is selling entitlement, price it as for-sale townhouse land.",
    "Ask the investment team for target IRR, minimum margin and hold period; until then the conventional residuals carry labelled hurdles only.",
    "If Common Ground is still wanted here, the credible re-scope is a smaller wood-frame program on the Simcoe frontage with a much lower cost per home, re-run through this model with QS costs.",
]

# ======================================================================
def memo_md():
    L = []
    L.append(f"# 1279 Simcoe Street North — Common Ground findings memo\n")
    L.append(f"**Prepared for:** Internal investment and development team · **Date:** 15 September 2026 · **Status:** Preliminary. R01 (1,000 homes) is an unapproved capacity concept.\n")
    L.append(f"**Bottom line:** Under the Common Ground Initiative the 1,000-home concept fails all four feasibility tests at the handoff's benchmark inputs, and it fails at a $0 land value as well. The supportable Common Ground appraisal is {money(resid_cg)}. The $17.0M ask is only explainable as for-sale townhouse land under the approved 172-unit plan.\n")
    L.append("## Headline numbers (R01, sequential phasing, benchmark inputs)\n")
    L.append("| Metric | Value |\n|---|---:|")
    rows = [("Homes / affordable", f"{m['homes']:,} / {m['aff_homes']:,}"), ("Total development cost", M(m['tdc'])), ("Cost per home", money(per_home)),
            ("Stabilized NOI", M(m['noi_stab'], 2)), ("NOI per home", money(noi_home)), ("Yield on cost", pct(m['yoc'], 2)), ("Stabilized value at cap", M(m['value'])),
            ("DSCR at stabilization (test ≥ 1.10)", xx(m['dscr'])), ("Weakest phase Year-10 DSCR (test ≥ 1.10)", xx(m['dscr10_min'])),
            ("Non-profit worst year after stabilization (test ≥ 0)", M(m['np_cf_min_year'], 2)), ("Year-20 takeout surplus / (shortfall)", M(m['refi20'] - m['debt20'])),
            ("Common Ground supportable appraisal", money(resid_cg)), ("Break-even hard $/sf at $17.0M", money(be_hard17)), ("Break-even rent uplift at $340/sf", pct(be_rent, 0)),
            ("Landowner all-in by Year 10 at $17.0M", M(m['lo_total'])), ("Peak public funding (City + BCH)", M(m['peak_public'])), ("Peak MLI Select balance", M(m['peak_mli'])),
            ("Phase 1 start / all phases stabilized", f"{dte(p1.m_start)} / {dte(m['stab_all'])}")]
    for a, b in rows: L.append(f"| {a} | {b} |")
    L.append("\n## Findings\n")
    for i, (h, t) in enumerate(findings, 1): L.append(f"{i}. **{h}** {t}\n")
    L.append("## Density and residual comparison (engine)\n")
    L.append("| Programme | Homes | TDC | TDC/home | Stab. NOI | YoC | DSCR | CG appraisal | Conventional residual |\n|---|---:|---:|---:|---:|---:|---:|---:|---:|")
    for d in dens:
        mm = d["m"]; L.append(f"| {d['meta']['label']} | {mm['homes']:,} | {M(mm['tdc'])} | {money(mm['tdc']/mm['homes'])} | {M(mm['noi_stab'],2)} | {pct(mm['yoc'],2)} | {xx(mm['dscr'])} | {money(d['resid_cg'])} | {M(d['conv']['land_binding'])} |")
    L.append("\n## Stress cases (engine)\n")
    L.append("| Scenario | TDC | Stab. NOI | YoC | DSCR | Min DSCR10 | NFP worst yr | Landowner all-in | Mezz @Y20 | Pass |\n|---|---:|---:|---:|---:|---:|---:|---:|---:|:--:|")
    for label, ov, mm in scen:
        vv = dict(v); vv.update(ov)
        L.append(f"| {label} | {M(mm['tdc'])} | {M(mm['noi_stab'],2)} | {pct(mm['yoc'],2)} | {xx(mm['dscr'])} | {xx(mm['dscr10_min'])} | {M(mm['np_cf_min_year'],2)} | {M(mm['lo_total'])} | {M(mm['mezz_bal20'])} | {'PASS' if E.passes(mm, vv) else 'FAIL'} |")
    L.append("\n## Input quality\n")
    L.append(f"{n_inputs} inputs; {n_flag} are PLACEHOLDER, UNVERIFIED or PENDING and are shown in red on the Inputs tab. The ones that move the verdict: hard_psf (Arterial/Altus benchmark, scope not inspected), dc_per_home (supplied placeholder), rent_0..3 (resale listings, not PBR comps; no studio comp), mmr_0..3 (CMHC portal unreachable), land_appraisal (seller ask, not an appraisal), mort_rate / cmhc_premium (no term sheet).\n")
    L.append("## Next steps\n")
    for s in next_steps: L.append(f"- {s}")
    L.append("\n## Files\n- `1279_Simcoe_CommonGround_Model.xlsx` — 14-tab formula-driven model (Read Me · Sources · Inputs · Programme · Schedule · Development Budget · Revenue & Opex · Debt · Monthly Cash Flow · Phase Returns · Exit & Residual · Sensitivities · Dashboard · Checks)\n- `1279_Simcoe_SOURCE_REGISTER.md` — every external input with source, date, units, confidence, owner, flag\n- `model/cg_engine.py`, `model/build_workbook.py`, `model/write_memo.py` — engine, builder (recalculates and reconciles in LibreOffice), memo renderer\n")
    L.append("\n*Illustrative underwriting. Not a valuation, appraisal, offer or financial advice. Subject to a QS budget, verified rents, CMHC MLI Select, Build Canada Homes and City of Oshawa approvals.*\n")
    return "\n".join(L)

CSS = """
:root{--espresso:#1f1612;--cream:#f3ead6;--cream2:#f7efdb;--gold:#c89344;--gold2:#b07f33;--ink:#161210;--muted:#6b5d4f;--line:#e0d5bd;--flag:#a23b2c}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Newsreader',Georgia,serif;color:var(--ink);background:var(--cream);line-height:1.55;-webkit-font-smoothing:antialiased}
.mono{font-family:'JetBrains Mono',ui-monospace,monospace}
.wrap{max-width:980px;margin:0 auto;padding:0 26px}
header.top{border-bottom:1px solid var(--line);padding:16px 0}
header.top .wrap{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.brand{font-size:20px;font-weight:600}.brand span{font-weight:300}
.nav{font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:var(--muted)}
.hero{background:var(--espresso);color:var(--cream);padding:54px 0 44px}
.kick{font-family:'JetBrains Mono',monospace;font-size:12px;letter-spacing:.24em;text-transform:uppercase;color:var(--gold)}
.hero h1{font-weight:300;font-size:50px;line-height:1.05;margin:18px 0 14px;letter-spacing:-.01em}.hero h1 em{color:var(--gold);font-style:italic}
.hero p{max-width:760px;color:#d8cdb5;font-size:18px;font-weight:300}.hero p b{color:var(--gold);font-weight:500}
.stripe{display:flex;flex-wrap:wrap;border:1px solid rgba(200,147,68,.32);margin:30px 0 0}
.stripe .cell{flex:1 1 0;min-width:150px;padding:16px 18px;border-right:1px solid rgba(200,147,68,.2)}.stripe .cell:last-child{border-right:0}
.stripe .v{font-size:28px;color:var(--cream);font-weight:300;line-height:1}.stripe .l{font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:#a99878;margin-top:8px}
section{padding:38px 0;border-top:1px solid var(--line)}
.secno{font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.22em;color:var(--gold2);text-transform:uppercase}
h2{font-size:34px;font-weight:300;margin:8px 0 14px;letter-spacing:-.012em}h2 em{font-style:italic;color:var(--gold)}
h3{font-family:'JetBrains Mono',monospace;font-size:12px;letter-spacing:.14em;text-transform:uppercase;margin:22px 0 8px}
p.lead{font-size:18px;color:#3a3128;max-width:800px;font-weight:300}
table{width:100%;border-collapse:collapse;margin:16px 0;font-size:14px}
th,td{text-align:left;padding:9px 12px;border-bottom:1px solid var(--line)}
th{background:var(--espresso);color:var(--cream);font-family:'JetBrains Mono',monospace;font-weight:500;font-size:10.5px;letter-spacing:.08em;text-transform:uppercase}
td.r,th.r{text-align:right}tr:nth-child(even) td{background:var(--cream2)}
ol.f{margin:14px 0 0 22px}ol.f li{margin:0 0 14px}ol.f b{color:var(--espresso)}
.note{background:var(--cream2);border-left:3px solid var(--gold);padding:12px 16px;margin:14px 0;font-size:15px}
.flag{color:var(--flag);font-weight:600}.pass{color:#5c6e2f;font-weight:600}
footer{background:var(--espresso);color:#bdae93;padding:30px 0;font-family:'JetBrains Mono',monospace;font-size:11px;line-height:1.8}footer .gold{color:var(--gold)}
ul.n{margin:10px 0 0 22px}ul.n li{margin:0 0 8px}
a{color:var(--gold2)}
@media (max-width:640px){.hero h1{font-size:36px}table{font-size:12px}}
"""
FONTS = ("<link rel='preconnect' href='https://fonts.googleapis.com'><link rel='preconnect' href='https://fonts.gstatic.com' crossorigin>"
         "<link href='https://fonts.googleapis.com/css2?family=Newsreader:ital,opsz,wght@0,6..72,300;0,6..72,400;0,6..72,500;0,6..72,600;1,6..72,300;1,6..72,400&family=JetBrains+Mono:wght@400;500&display=swap' rel='stylesheet'>")

def memo_html():
    e = html.escape
    head = f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow"><title>1279 Simcoe Street North — Common Ground findings</title>{FONTS}<style>{CSS}</style></head><body>
<header class="top"><div class="wrap"><span class="brand">Ubuntu Land Trust <span>&amp; Developments</span></span><span class="nav">Common Ground Initiative · Findings memo · 15 Sep 2026</span></div></header>
<div class="hero"><div class="wrap"><div class="kick">1279 Simcoe Street North, Oshawa · R01 1,000-home concept</div>
<h1>Does not close <em>at any land value</em> on benchmark costs.</h1>
<p>Under the Common Ground stack (MLI Select 50% · Build Canada Homes 40% · City 5% · landowner equity 5%, 30% of homes at 80% of median market rent) the 1,000-home masterplan fails all four feasibility tests, and still fails with the land at <b>$0</b>. The <b>$17.0M</b> ask is a for-sale townhouse number under the approved 172-unit plan, not a rental number.</p>
<div class="stripe">
<div class="cell"><div class="v">{M(m['tdc'])}</div><div class="l">Total development cost</div></div>
<div class="cell"><div class="v">{money(per_home)}</div><div class="l">Cost per home</div></div>
<div class="cell"><div class="v">{pct(m['yoc'],2)}</div><div class="l">Yield on cost vs {pct(v['cap_rate'],2)} cap</div></div>
<div class="cell"><div class="v">{xx(m['dscr'])}</div><div class="l">DSCR (test ≥ {v['min_dscr']:.2f}x)</div></div>
<div class="cell"><div class="v">{money(resid_cg)}</div><div class="l">Supportable CG appraisal</div></div>
<div class="cell"><div class="v">{money(be_hard17)}/sf</div><div class="l">Break-even hard cost at $17M</div></div>
</div></div></div>
"""
    body = ["<section><div class='wrap'><div class='secno'>01</div><h2>Findings</h2><ol class='f'>"]
    for h, t in findings: body.append(f"<li><b>{e(h)}</b> {e(t)}</li>")
    body.append("</ol></div></section>")
    body.append("<section><div class='wrap'><div class='secno'>02</div><h2>Common Ground <em>tests</em></h2><table><tr><th>Test</th><th class='r'>Result</th><th class='r'>Hurdle</th><th>Status</th></tr>")
    rows = [("1 · Consolidated DSCR at stabilization", xx(detail['dscr'][0]), xx(detail['dscr'][1]), detail['dscr'][2]),
            ("2 · Weakest phase Year-10 DSCR incl. supplemental charge", xx(detail['dscr10'][0]), "1.10x", detail['dscr10'][2]),
            ("3 · Non-profit worst annual cash flow after stabilization", M(detail['nfp'][0], 2), "≥ $0", detail['nfp'][2]),
            ("4 · Year-20 takeout refinance vs debt to clear", f"{M(detail['takeout'][0])} vs {M(detail['takeout'][1])}", "refi ≥ debt", detail['takeout'][2])]
    for a, b, c, ok in rows: body.append(f"<tr><td>{e(a)}</td><td class='r'>{b}</td><td class='r'>{c}</td><td class='{'pass' if ok else 'flag'}'>{'PASS' if ok else 'FAIL'}</td></tr>")
    body.append("</table><div class='note'>What has to be true at a $17.0M appraisal: above-ground hard cost near <b>" + money(be_hard17) + "/sf</b> (benchmark " + money(v['hard_psf']) + "/sf), or rents about <b>" + pct(be_rent, 0) + "</b> above the resale comps. At $0 appraisal the break-even hard cost is " + money(be_hard0) + "/sf.</div></div></section>")
    body.append("<section><div class='wrap'><div class='secno'>03</div><h2>Density and <em>residual</em></h2><table><tr><th>Programme</th><th class='r'>Homes</th><th class='r'>TDC</th><th class='r'>TDC/home</th><th class='r'>Stab. NOI</th><th class='r'>YoC</th><th class='r'>DSCR</th><th class='r'>CG appraisal</th><th class='r'>Conventional residual</th></tr>")
    for d in dens:
        mm = d["m"]; body.append(f"<tr><td>{e(d['meta']['label'])}</td><td class='r'>{mm['homes']:,}</td><td class='r'>{M(mm['tdc'])}</td><td class='r'>{money(mm['tdc']/mm['homes'])}</td><td class='r'>{M(mm['noi_stab'],2)}</td><td class='r'>{pct(mm['yoc'],2)}</td><td class='r'>{xx(mm['dscr'])}</td><td class='r'>{money(d['resid_cg'])}</td><td class='r'>{M(d['conv']['land_binding'])}</td></tr>")
    body.append("</table><p class='lead'>600 and 800 homes are analyst re-builds (fewer storeys, fewer phases, lower $/sf); 400 is a sensitivity, not a confirmed as-of-right yield. Conventional residuals use labelled hurdles (" + pct(v['target_yoc'], 1) + " yield on cost, " + pct(v['target_margin'], 0) + " margin), not a house standard.</p>")
    body.append("<h3>Approved 172-townhouse plan — for-sale cross-check (placeholder costs)</h3><table><tr><th>New-build sale price</th>" + "".join(f"<th class='r'>${k}/sf</th>" for k in th) + "</tr><tr><td>Residual land value at 15% margin</td>" + "".join(f"<td class='r'>{M(x)}</td>" for x in th.values()) + "</tr></table><p class='lead'>Resale comps for 50-year-old stock near the site sit around $410/sf (Arterial, 8 active listings, Jun–Jul 2026). Katanna's new-build pricing is VIP-only. On these placeholders the plan needs about " + money(th_be) + "/sf of new-build pricing before land has value; the ask rests on the vendor's own for-sale pro forma, which has not been supplied.</p></div></section>")
    body.append("<section><div class='wrap'><div class='secno'>04</div><h2>Stress <em>cases</em></h2><table><tr><th>Scenario</th><th class='r'>TDC</th><th class='r'>NOI</th><th class='r'>YoC</th><th class='r'>DSCR</th><th class='r'>Min DSCR10</th><th class='r'>NFP worst yr</th><th class='r'>Landowner all-in</th><th class='r'>Mezz @Y20</th><th>Pass</th></tr>")
    for label, ov, mm in scen:
        vv = dict(v); vv.update(ov); ok = E.passes(mm, vv)
        body.append(f"<tr><td>{e(label)}</td><td class='r'>{M(mm['tdc'])}</td><td class='r'>{M(mm['noi_stab'],2)}</td><td class='r'>{pct(mm['yoc'],2)}</td><td class='r'>{xx(mm['dscr'])}</td><td class='r'>{xx(mm['dscr10_min'])}</td><td class='r'>{M(mm['np_cf_min_year'],2)}</td><td class='r'>{M(mm['lo_total'])}</td><td class='r'>{M(mm['mezz_bal20'])}</td><td class='{'pass' if ok else 'flag'}'>{'PASS' if ok else 'FAIL'}</td></tr>")
    body.append("</table></div></section>")
    body.append("<section><div class='wrap'><div class='secno'>05</div><h2>Exposure and <em>who gets what</em></h2><table><tr><th>Item</th><th class='r'>Value</th><th>Note</th></tr>")
    ex = [("Peak cumulative public funding (City + BCH contribution + BCH mezz)", M(m['peak_public']), "front-loaded; phases fund sequentially"),
          ("Peak MLI Select balance", M(m['peak_mli']), "50% of TDC by design, not sized to coverage"),
          ("BCH mezz outstanding at Year 20 (before takeout)", M(m['mezz_bal20']), "must be cleared by the takeout refinance"),
          ("Landowner Day-1 draw / Yrs 1-10 share / Year-10 buyout", f"{M(m['land_draw'])} / {M(m['lo_cf'])} / {M(m['lo_buyout'])}", "all-in " + M(m['lo_total']) + " at a $17.0M appraisal"),
          ("Appraised value donated to affordability", M(m['donated']), "draw + retained equity already exceed the ask"),
          ("Phase 1 construction start / all phases stabilized / hold end", f"{dte(p1.m_start)} / {dte(m['stab_all'])} / {dte(m['hold_end_m'])}", f"{v['entitle_months']}-month entitlement, sequential phases")]
    for a, b, c in ex: body.append(f"<tr><td>{e(a)}</td><td class='r'>{b}</td><td>{e(c)}</td></tr>")
    body.append("</table></div></section>")
    body.append(f"<section><div class='wrap'><div class='secno'>06</div><h2>Input <em>quality</em> and next steps</h2><p class='lead'>{n_inputs} inputs; <span class='flag'>{n_flag}</span> are placeholder, unverified or pending and shown in red on the Inputs tab. The ones that decide the verdict are the hard-cost benchmark, the development-charge placeholder, the resale-listing rents, the missing CMHC median rents, the seller's ask standing in for an appraisal, and the untested lender terms.</p><ul class='n'>")
    for s in next_steps: body.append(f"<li>{e(s)}</li>")
    body.append("</ul><h3>Files</h3><ul class='n'><li><a href='1279_Simcoe_CommonGround_Model.xlsx'>1279_Simcoe_CommonGround_Model.xlsx</a> — 14-tab formula-driven model, reconciled to the engine by LibreOffice recalculation</li><li><a href='1279_Simcoe_SOURCE_REGISTER.md'>1279_Simcoe_SOURCE_REGISTER.md</a> — assumption and source register</li><li><a href='1279_Simcoe_FINDINGS_MEMO.md'>1279_Simcoe_FINDINGS_MEMO.md</a> — this memo as Markdown</li></ul></div></section>")
    foot = f"""<footer><div class="wrap"><b class="gold">Ubuntu Land Trust &amp; Developments</b> · "I am because we are"<br><br>Illustrative underwriting for {e(ADDR := '1279 Simcoe Street North, Oshawa, ON L1G 4X1')}. Programme quantities from Blender masterplan R01; construction cost, property-tax rate, leasing and sales comps from the Arterial connector; land value is the seller's ask standing in for an appraisal. Cap rate, vacancy, operating ratios and financing terms are modelling assumptions. Not a valuation, appraisal, offer or financial advice. Subject to a QS budget, verified rent schedule and third-party approvals (CMHC MLI Select, Build Canada Homes, City of Oshawa).</div></footer></body></html>"""
    return head + "".join(body) + foot

def register_md():
    L = ["# 1279 Simcoe Street North — assumption & source register\n", "Every input in `model/cg_engine.py` → Inputs tab. Flags: **PLACEHOLDER** = analyst number with no source; **UNVERIFIED** = supplied or connector number not checked to its source; **PENDING** = required input that could not be obtained in this session.\n",
         "| Key | Value | Unit | Label | Source / basis | Date | Conf. | Owner | Flag |", "|---|---:|---|---|---|---|:--:|---|---|"]
    for k, d in inp.items():
        val = d["value"]
        vs = f"{val:,.4g}" if isinstance(val, float) and abs(val) < 1 else (f"{val:,}" if isinstance(val, (int, float)) else str(val))
        L.append(f"| `{k}` | {vs} | {d['unit']} | {d['label']} | {d['source']} | {d['date']} | {d['conf']} | {d['owner']} | {('**'+d['flag']+'**') if d['flag'] else ''} |")
    L.append("\n## External sources\n")
    L.append("| # | Source | Supplied | Date | Reached | Used for |\n|---|---|---|---|---|---|")
    srcs = [
        ("City of Oshawa ED-24-20 Attachment 2 — Proposed Draft Plan of Subdivision, 1279 Simcoe St N, Katanna Simcoe Ltd. (S-O-2022-05, Z-2022-12, C-O-2022-08)", "Block 1 3.1211 ha / 172 condo townhouses; Block 2 road widening; PIN assembly; two excluded lots", "Jan 2024", "Yes (uploaded PDF)", "site_net_m2; approved-plan cross-check"),
        ("Financial Modelling Handoff (internal)", "R01 quantities, suite mix, $17M ask, 7.85 ac, ~$500k existing NOI, DC placeholder, tax rate, cost benchmark", "2026-09-15", "Yes", "Programme, scenario framework"),
        ("Arterial resolve_site", "17,014 m² single PIN; 5 buildings, 6,070 m² footprint; Regional Road 2", "2026-09-15", "Yes", "exist_footprint"),
        ("Arterial financial_context", "$290-390/sf up to 12 storeys; $230-270/sf 3-storey stacked townhouse; 0.9823% tax; NO Oshawa DC dataset; CMHC zone empty", "2026-09-15", "Yes", "hard_psf, tax_rate"),
        ("Arterial run_development_scenario (townhouse, rental)", "1-BR $1,795 (1 comp), 2-BR $2,173 (2), 3-BR $2,412 (4); negative residual as rental", "2026-09-15", "Yes", "rent_1..3"),
        ("Arterial find_unit_comps (sales, condo townhouse)", "8 active listings ~$388-411/sf, 50-year-old stock", "2026-09-15", "Yes", "for-sale cross-check"),
        ("Arterial run_hbu_development_scenario", "never returned ('computing'); result URL blocked", "2026-09-15", "No", "—"),
        ("CMHC Rental Market Survey Oct 2025 (HMIP)", "median market rents", "2025-10", "No (blocked)", "mmr_0..3 PENDING"),
        ("City of Oshawa / Region of Durham DC schedules; By-law 104-2025; DRHBA", "per-unit DC rates", "2025", "No (blocked)", "dc_per_home UNVERIFIED"),
        ("Oshawa council memo (escribe 14949); Durham Post", "approval status", "2024", "No (blocked)", "status UNCONFIRMED"),
        ("Platinum Condo Deals / LoopNet / Zolo", "165 units in 20 blocks now marketed; plaza units for lease", "2026", "Snippets only", "notes"),
        ("Ubuntu Common Ground skill, July 2026 structure", "capital stack, splits, draw, buyout, takeout, opex defaults", "2026-07", "Yes", "all Common Ground mechanics"),
    ]
    for i, s in enumerate(srcs, 1): L.append(f"| {i} | {s[0]} | {s[1]} | {s[2]} | {s[3]} | {s[4]} |")
    return "\n".join(L) + "\n"

def index_html():
    e = html.escape
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow">
<title>1279 Simcoe Street North</title>{FONTS}<style>{CSS}</style></head><body>
<header class="top"><div class="wrap"><span class="brand">Ubuntu Land Trust <span>&amp; Developments</span></span><span class="nav">Common Ground Initiative · 1279 Simcoe Street North</span></div></header>
<div class="hero"><div class="wrap"><div class="kick">Oshawa · R01 1,000-home concept · financial modelling handoff</div><h1>1279 Simcoe Street North <em>package</em></h1>
<p>Findings memo, formula-driven model and source register for the September 2026 handoff, run under the Common Ground structure. <b>Verdict: does not close at any land value on benchmark costs.</b></p></div></div>
<section><div class="wrap"><table><tr><th>Document</th><th>What it is</th></tr>
<tr><td><a href="1279_Simcoe_FINDINGS_MEMO.html">Findings memo</a></td><td>One-page findings, tests, density and stress results, next steps</td></tr>
<tr><td><a href="1279_Simcoe_CommonGround_Model.xlsx">Common Ground model (.xlsx)</a></td><td>14 tabs, live formulas, engine-computed sensitivities, checks</td></tr>
<tr><td><a href="1279_Simcoe_SOURCE_REGISTER.md">Source register</a></td><td>Every input with source, date, units, confidence, owner and flag</td></tr>
<tr><td><a href="model/">model/</a></td><td>cg_engine.py · build_workbook.py · write_memo.py · engine_summary.json</td></tr>
</table></div></section>
<footer><div class="wrap"><b class="gold">Ubuntu Land Trust &amp; Developments</b> · Preliminary scenario brief — not a valuation, appraisal, offer or financial advice.</div></footer></body></html>"""

if __name__ == "__main__":
    open(os.path.join(OUT, "1279_Simcoe_FINDINGS_MEMO.md"), "w").write(memo_md())
    open(os.path.join(OUT, "1279_Simcoe_FINDINGS_MEMO.html"), "w").write(memo_html())
    open(os.path.join(OUT, "1279_Simcoe_SOURCE_REGISTER.md"), "w").write(register_md())
    open(os.path.join(OUT, "index.html"), "w").write(index_html())
    print("memo, register, index written")
    print(json.dumps(dict(tdc=m["tdc"], per_home=per_home, noi=m["noi_stab"], yoc=m["yoc"], dscr=m["dscr"], dscr10=m["dscr10_min"], nfp=m["np_cf_min_year"],
                          takeout_gap=m["refi20"] - m["debt20"], resid_cg=resid_cg, be_hard17=be_hard17, be_hard0=be_hard0, be_rent=be_rent, be_dc=be_dc,
                          lo_total=m["lo_total"], peak_public=m["peak_public"], peak_mli=m["peak_mli"], th=th, value=m["value"]), indent=1))
