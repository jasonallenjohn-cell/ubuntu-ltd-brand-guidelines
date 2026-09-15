# 1279 Simcoe Street North — assumption & source register

Every input in `model/cg_engine.py` → Inputs tab. Flags: **PLACEHOLDER** = analyst number with no source; **UNVERIFIED** = supplied or connector number not checked to its source; **PENDING** = required input that could not be obtained in this session.

| Key | Value | Unit | Label | Source / basis | Date | Conf. | Owner | Flag |
|---|---:|---|---|---|---|:--:|---|---|
| `site_net_m2` | 31,211 | m² | Net site (R01 Block 1) | Draft plan S-O-2022-05, Attachment 2 to ED-24-20 (3.1211 ha) | 2024-01 | H | Survey |  |
| `site_gross_ac` | 7.85 | acres | Gross site per acquisition summary | Supplied acquisition summary; reconcile to survey | 2026-09 | L | Vendor | **UNVERIFIED** |
| `land_appraisal` | 17,000,000 | $ | Land value used for Common Ground draw (seller ask as placeholder) | Seller ask $17.0M; NOT an appraisal. Common Ground requires an AACI appraisal | 2026-09 | L | Vendor | **PLACEHOLDER** |
| `draw_pct` | 0.6 | % | Day-1 landowner draw as % of appraisal | Common Ground July 2026 structure (locked) | 2026-07 | H | Ubuntu |  |
| `land_eq_pct` | 0.05 | % of TDC | Landowner retained equity | Common Ground July 2026 structure (locked) | 2026-07 | H | Ubuntu |  |
| `draw_timing` | 1 | 1 = pro rata by phase homes; 0 = 100% at Phase 1 close | Land draw timing | Analyst assumption; editable | 2026-09 | M | Analyst |  |
| `existing_noi` | 500,000 | $/yr | Existing NOI (seller-supplied) | Vendor summary; rent roll, leases and expenses NOT provided | 2026-09 | L | Vendor | **UNVERIFIED** |
| `existing_noi_keep` | 0 | share | Share of existing NOI retained during entitlement (0 = zero-retained-income case) | Handoff §2: use zero-retained case unless leases support retention | 2026-09 | M | Analyst |  |
| `legal_title` | 150,000 | $ | Legal, title, lease registration, due diligence | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `ltt_applies` | 0 | 1/0 | Land transfers to SPV (Ontario LTT applies)? | Common Ground = land lease, no transfer; set 1 if structure changes | 2026-09 | M | Ubuntu |  |
| `carry_taxes` | 120,000 | $/yr | Property tax + insurance on existing site during entitlement | Analyst placeholder pending tax bill | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `t_start` | 2027-01-01 | date | Month 1 (financial close / lease commencement) | Analyst assumption; editable | 2026-09 | M | Analyst |  |
| `entitle_months` | 18 | months | Entitlement (OPA/ZBA + SPA) before Phase 1 construction | Analyst assumption; planner input required | 2026-09 | L | Planner | **PLACEHOLDER** |
| `entitle_delay` | 0 | months | Approval delay stress (added to entitlement) | Scenario lever | 2026-09 | H | Analyst |  |
| `constr_m_8` | 26 | months | Construction duration, 8-storey building | Analyst assumption; no schedule established (handoff §3) | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `constr_m_12` | 30 | months | Construction duration, 12-storey building | Analyst assumption; no schedule established (handoff §3) | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `seq_mode` | 1 | 1 = sequential; 0 = overlap by interval | Phase sequencing | Handoff §6: sequential reference case | 2026-09 | H | Analyst |  |
| `overlap_interval` | 12 | months | Start interval between phases when overlapping | Scenario lever | 2026-09 | H | Analyst |  |
| `later_defer` | 0 | months | Defer Phases 2-4 (stress) | Scenario lever | 2026-09 | H | Analyst |  |
| `lease_cap` | 25 | homes/month | Sitewide leasing capacity (all phases compete) | Analyst assumption; no absorption evidence supplied | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `lease_cap_mult` | 1.0 | × | Leasing pace multiplier (stress) | Scenario lever | 2026-09 | H | Analyst |  |
| `retail_delay` | 6 | months | Retail lease-up delay after phase completion | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `hold_after_stab` | 120 | months | Operating hold after final stabilization | Handoff §1 | 2026-09 | H | Client |  |
| `rent_0` | 1,495 | $/mo | Studio market rent | NO COMP SUPPLIED. Analyst placeholder from 1-BR $/sf scaled; verify new-build achieved rents | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `rent_1` | 1,795 | $/mo | 1-bedroom market rent | Arterial leasing comps, 1 comp (650 sf resale condo, Centennial) | 2026-09 | L | Arterial | **UNVERIFIED** |
| `rent_2` | 2,173 | $/mo | 2-bedroom market rent | Arterial leasing comps, 2 comps (avg 1,025 sf resale condo/townhouse) | 2026-09 | L | Arterial | **UNVERIFIED** |
| `rent_3` | 2,412 | $/mo | 3-bedroom market rent | Arterial leasing comps, 4 comps (avg 1,300 sf, e.g. 53 Taunton Rd E $2,400-2,450) | 2026-09 | L | Arterial | **UNVERIFIED** |
| `newbuild_prem` | 0 | % | New-build rent premium over comps | Handoff §1: do not assume a premium | 2026-09 | H | Analyst |  |
| `rent_shock` | 0 | % | Starting-rent stress (±) | Scenario lever | 2026-09 | H | Analyst |  |
| `comp_sf_1` | 650 | sf | Comp reference size 1-BR | Arterial comps | 2026-09 | M | Arterial |  |
| `comp_sf_2` | 1,025 | sf | Comp reference size 2-BR | Arterial comps | 2026-09 | M | Arterial |  |
| `comp_sf_3` | 1,300 | sf | Comp reference size 3-BR | Arterial comps | 2026-09 | M | Arterial |  |
| `mmr_0` | 1,250 | $/mo | CMHC median market rent, bachelor (Oshawa CMA) | PENDING — CMHC Rental Market Survey Oct 2025 (HMIP portal not reachable from this session); Arterial CMHC zone 'Oshawa (North)' returned no value | 2025-10 | L | CMHC | **PENDING** |
| `mmr_1` | 1,500 | $/mo | CMHC median market rent, 1-bedroom | PENDING — as above | 2025-10 | L | CMHC | **PENDING** |
| `mmr_2` | 1,750 | $/mo | CMHC median market rent, 2-bedroom | PENDING — as above | 2025-10 | L | CMHC | **PENDING** |
| `mmr_3` | 1,950 | $/mo | CMHC median market rent, 3-bedroom | PENDING — as above | 2025-10 | L | CMHC | **PENDING** |
| `aff_share` | 0.3 | % of homes | Affordable share of homes (per type) | Common Ground locked default | 2026-07 | H | Ubuntu |  |
| `aff_factor` | 0.8 | % of MMR | Affordable rent as % of CMHC MMR | Common Ground locked default | 2026-07 | H | Ubuntu |  |
| `rent_growth` | 0.03 | %/yr | Market rent growth | Common Ground locked default | 2026-07 | M | Ubuntu |  |
| `opex_growth` | 0.025 | %/yr | Operating expense growth | Common Ground locked default | 2026-07 | M | Ubuntu |  |
| `stab_occ` | 0.97 | % | Stabilized physical occupancy (vacancy 3%) | Common Ground locked default (3% vacancy) | 2026-07 | M | Ubuntu |  |
| `bad_debt` | 0.005 | % of rent | Bad debt / concessions | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `opex_util` | 1,200 | $/occupied home/yr | Utilities (variable) | Common Ground generator default | 2026-07 | M | Ubuntu |  |
| `opex_rm` | 900 | $/occupied home/yr | Repairs, maintenance, turnover (variable) | Common Ground generator default | 2026-07 | M | Ubuntu |  |
| `opex_ins` | 450 | $/home/yr | Insurance (fixed on delivered homes) | Common Ground generator default | 2026-07 | M | Ubuntu |  |
| `opex_staff` | 400 | $/home/yr | On-site staffing (fixed) | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `opex_ga` | 300 | $/home/yr | Administration (fixed) | Common Ground generator default | 2026-07 | M | Ubuntu |  |
| `mgmt_pct` | 0.04 | % of EGI | Property management | Common Ground generator default | 2026-07 | M | Ubuntu |  |
| `reserve` | 500 | $/home/yr | Replacement reserve (below valuation NOI; deducted for lender DSCR) | Common Ground generator default; treatment per handoff §4 | 2026-07 | M | Ubuntu |  |
| `tax_rate` | 0.009823 | % of assessed | Multi-residential tax rate, Oshawa 2025 | Arterial (Oshawa 2025). UNVERIFIED — confirm class & current rate | 2025 | L | Arterial | **UNVERIFIED** |
| `assess_ratio` | 0.6 | × stabilized value | Assessed value as share of stabilized value (new PBR, MPAC lag) | Analyst placeholder; not cost, not price | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `cap_rate` | 0.0425 | % | Valuation cap rate | Common Ground locked default | 2026-07 | M | Ubuntu |  |
| `cap_shock` | 0 | bps as decimal | Cap-rate stress (+0.01 = +100 bps) | Scenario lever | 2026-09 | H | Analyst |  |
| `pk_m2_space` | 32 | m²/space | Below-grade gross area per space | R01 volume allowance (unvalidated layout) | 2026-09 | L | Design | **UNVERIFIED** |
| `pk_ratio` | 0.65 | spaces/home | Parking ratio (aggregate) | R01; scenario lever | 2026-09 | M | Design |  |
| `pk_visitor` | 0.1 | % of spaces | Visitor share (not leased) | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `pk_takeup` | 0.6 | % of resident spaces | Resident take-up | Analyst placeholder; no evidence supplied | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `pk_rent` | 125 | $/space/mo | Parking rent | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `ret_eff` | 0.9 | % | Retail rentable efficiency | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `ret_rent` | 28 | $/sf/yr net | Retail net rent | Analyst placeholder; no retail comps supplied | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `ret_vac` | 0.1 | % | Retail vacancy | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `ret_ti` | 60 | $/sf | Retail TI + inducements | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `hard_psf` | 340 | $/sf above-ground gross | Above-ground construction, up to 12 storeys (midpoint) | Arterial / Altus benchmark $290-390/sf, 'Condos, up to 12 storeys'. Scope, basis and price date NOT inspected | 2026-09 | L | Arterial | **UNVERIFIED** |
| `hard_shock` | 0 | % | Hard-cost stress (±) | Scenario lever | 2026-09 | H | Analyst |  |
| `pk_psf` | 200 | $/sf below-grade gross | Underground parking incl. ramps/systems | Analyst placeholder; assumes benchmark EXCLUDES parking (set 0 if included) | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `demo_m2` | 75 | $/m² existing footprint | Demolition + hazmat | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `exist_footprint` | 6,070 | m² | Existing building footprint (5 buildings) | Arterial site characteristics (2026-07) | 2026-07 | M | Arterial |  |
| `tenant_term` | 500,000 | $ | Tenant relocation / lease termination | Analyst placeholder; leases not provided | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `remed_offsite` | 2,500,000 | $ | Remediation, earthworks, off-site & utility upgrades | Analyst placeholder; no geotech/ESA/servicing study | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `site_m2` | 150 | $/m² net site | Streets, landscaping, public realm, shared infrastructure | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `infra_p1_share` | 0.6 | % | Share of site works + remediation paid in Phase 1 (front-loaded) | Handoff §3: Phase 1 carries early shared infrastructure | 2026-09 | M | Analyst |  |
| `escalation` | 0.03 | %/yr | Escalation from benchmark date (Sept 2026) to phase mid-construction | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `cont_constr` | 0.075 | % of hard | Construction contingency | Analyst placeholder | 2026-09 | M | Analyst |  |
| `cont_design` | 0.025 | % of hard | Design contingency | Analyst placeholder | 2026-09 | M | Analyst |  |
| `design_pct` | 0.08 | % of hard | Design & consultants | Analyst placeholder | 2026-09 | M | Analyst |  |
| `permit_pct` | 0.01 | % of hard | Approvals, permits, planning fees | Analyst placeholder | 2026-09 | M | Analyst |  |
| `dc_per_home` | 102,653 | $/home | Development charges incl. Region + education (placeholder) | Supplied placeholder $102,653/home; NOT a verified payable rate. City/Region 2025 schedules not reachable from this session | 2026-09 | L | Client | **UNVERIFIED** |
| `dc_aff_exempt` | 0 | 1/0 | Affordable units DC-exempt (DC Act s.4.1)? | Default 0 (conservative). Ontario exempts qualifying affordable rental units since June 2024; eligibility to be confirmed with City/Region | 2026-09 | M | Analyst |  |
| `dc_credits` | 0 | $ | DC credits (existing units / prior payments) | None evidenced | 2026-09 | L | Analyst | **PENDING** |
| `lease_mkt` | 1,500 | $/home | Leasing, marketing, pre-opening | Analyst placeholder | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `dev_fee` | 0.03 | % of hard+soft | Development management fee (Ubuntu / NFP) | Analyst placeholder | 2026-09 | M | Analyst |  |
| `builders_risk` | 0.005 | % of hard | Builder's risk & wrap-up insurance | Analyst placeholder | 2026-09 | M | Analyst |  |
| `hst_net` | 0 | % of hard+soft | Net non-recoverable HST | PENDING tax advice — enhanced GST rental rebate (2023) and Ontario rebate eligibility not confirmed | 2026-09 | L | Tax | **PENDING** |
| `cmhc_share` | 0.5 | % of TDC | CMHC MLI Select insured loan | Common Ground July 2026 (locked) | 2026-07 | H | Ubuntu |  |
| `bch_share` | 0.4 | % of TDC | Build Canada Homes tranche | Common Ground July 2026 (locked) | 2026-07 | H | Ubuntu |  |
| `city_share` | 0.05 | % of TDC | City contribution | Common Ground July 2026 (locked) | 2026-07 | H | Ubuntu |  |
| `grant_share` | 0.545 | % of BCH | BCH conditionally-repayable contribution share | Common Ground July 2026 (~54.5%); raise if dscr10 < 1.10 | 2026-07 | M | Ubuntu |  |
| `mezz_rate` | 0.045 | % | BCH mezz rate (BoC 10-yr + 100 bps), accrues from Year 10 | Common Ground July 2026; BoC bond yield UNVERIFIED | 2026-07 | L | Ubuntu | **UNVERIFIED** |
| `mort_rate` | 0.0425 | % | MLI Select rate (construction + permanent) | Common Ground locked default; lender term sheet NOT obtained | 2026-07 | L | Ubuntu | **UNVERIFIED** |
| `rate_shock` | 0 | decimal | Borrowing-rate stress (+0.02 = +200 bps) | Scenario lever | 2026-09 | H | Analyst |  |
| `amort_yrs` | 50 | years | MLI Select amortization | Common Ground locked default | 2026-07 | H | Ubuntu |  |
| `cmhc_premium` | 0.025 | % of loan | CMHC insurance premium (MLI Select 100-point tier) | Analyst placeholder; confirm premium schedule | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `cmhc_app_fee` | 200 | $/home | CMHC application fee | CMHC schedule (analyst recollection) | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `leaseup_reserve_m` | 6 | months of MLI debt service | Lease-up operating / interest reserve capitalized in TDC | Analyst placeholder; lender completion-reserve requirement not obtained | 2026-09 | L | Analyst | **PLACEHOLDER** |
| `takeout_ltv` | 0.75 | % | Year-20 takeout refinance LTV (must clear MLI + supplemental + mezz) | Common Ground generator | 2026-07 | M | Ubuntu |  |
| `min_dscr` | 1.1 | × | Minimum DSCR (MLI Select affordability tier) | CMHC MLI Select | 2026-07 | M | CMHC |  |
| `split_lo` | 0.6 | % | Landowner share of net cash flow, Yrs 1-10 (per phase) | Common Ground July 2026 (locked) | 2026-07 | H | Ubuntu |  |
| `buyout_prem` | 0.01 | % | Supplemental charge sizing over buyout (fees) | Common Ground generator | 2026-07 | M | Ubuntu |  |
| `sell_cost` | 0.02 | % of price | Selling costs (stabilize-and-sell test) | Analyst placeholder | 2026-09 | M | Analyst |  |
| `target_yoc` | 0.055 | % | Hurdle: unlevered yield on cost (conventional cross-check) | Handoff §7: labelled hurdle, not a house standard | 2026-09 | M | Client |  |
| `target_margin` | 0.15 | % | Hurdle: development margin on cost (conventional cross-check) | Handoff §7: labelled hurdle | 2026-09 | M | Client |  |
| `disc_rate` | 0.07 | % | Discount rate for dated NPV | Handoff §7: labelled hurdle | 2026-09 | M | Client |  |

## External sources

| # | Source | Supplied | Date | Reached | Used for |
|---|---|---|---|---|---|
| 1 | City of Oshawa ED-24-20 Attachment 2 — Proposed Draft Plan of Subdivision, 1279 Simcoe St N, Katanna Simcoe Ltd. (S-O-2022-05, Z-2022-12, C-O-2022-08) | Block 1 3.1211 ha / 172 condo townhouses; Block 2 road widening; PIN assembly; two excluded lots | Jan 2024 | Yes (uploaded PDF) | site_net_m2; approved-plan cross-check |
| 2 | Financial Modelling Handoff (internal) | R01 quantities, suite mix, $17M ask, 7.85 ac, ~$500k existing NOI, DC placeholder, tax rate, cost benchmark | 2026-09-15 | Yes | Programme, scenario framework |
| 3 | Arterial resolve_site | 17,014 m² single PIN; 5 buildings, 6,070 m² footprint; Regional Road 2 | 2026-09-15 | Yes | exist_footprint |
| 4 | Arterial financial_context | $290-390/sf up to 12 storeys; $230-270/sf 3-storey stacked townhouse; 0.9823% tax; NO Oshawa DC dataset; CMHC zone empty | 2026-09-15 | Yes | hard_psf, tax_rate |
| 5 | Arterial run_development_scenario (townhouse, rental) | 1-BR $1,795 (1 comp), 2-BR $2,173 (2), 3-BR $2,412 (4); negative residual as rental | 2026-09-15 | Yes | rent_1..3 |
| 6 | Arterial find_unit_comps (sales, condo townhouse) | 8 active listings ~$388-411/sf, 50-year-old stock | 2026-09-15 | Yes | for-sale cross-check |
| 7 | Arterial run_hbu_development_scenario | never returned ('computing'); result URL blocked | 2026-09-15 | No | — |
| 8 | CMHC Rental Market Survey Oct 2025 (HMIP) | median market rents | 2025-10 | No (blocked) | mmr_0..3 PENDING |
| 9 | City of Oshawa / Region of Durham DC schedules; By-law 104-2025; DRHBA | per-unit DC rates | 2025 | No (blocked) | dc_per_home UNVERIFIED |
| 10 | Oshawa council memo (escribe 14949); Durham Post | approval status | 2024 | No (blocked) | status UNCONFIRMED |
| 11 | Platinum Condo Deals / LoopNet / Zolo | 165 units in 20 blocks now marketed; plaza units for lease | 2026 | Snippets only | notes |
| 12 | Ubuntu Common Ground skill, July 2026 structure | capital stack, splits, draw, buyout, takeout, opex defaults | 2026-07 | Yes | all Common Ground mechanics |
