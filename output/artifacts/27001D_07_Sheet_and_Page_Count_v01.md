# Design Deliverable Count Estimate (Drawings, Specs, BOD, Calcs) — FY27 CIP MINCON: Substation Generators — NSF Dahlgren, VA

**Pursuit:** 27001D · **Solicitation:** N4008027R0004 (NAVFAC Project 1921806 / PID N40080-27-RFPREQ-SP DAHL ROICC-0004)  
**Prime A/E (DOR):** Bowman, Foster & Associates, Inc. (BFA) · **DB Offeror (GC):** Encon Corporation  
**Prepared:** 09 Nov 2026 · **Status:** Working draft for internal review

<!-- INTERNAL — estimating basis per 00_admin/drawing_set_and_spec_baseline.SCHEMA.md. Sheet list, ranges and drivers are internal; only the Summary table is publishable (SCHEMA §7). -->
> **Agent-proposed ranges.** Per SCHEMA §8, the low/expected/high values and sheet justifications are a proposal for the responsible engineer to overwrite. The calibration ledger has no closed rows, so these ranges come from judgement, not data.

## Summary (publishable)

| Deliverable | Low | Expected | High | Notes |
|---|---|---|---|---|
| Drawings (sheets, all disciplines) | 32 | 36 | 43 | 7 seal holders; MEP (M+E) expected = 24; P merged onto M sheets |
| Specifications (pages) | 590 | 738 | 922 | 44 line items incl. 4 Part 5 prescriptive sections (unmodified) and product data; SpecsIntact |
| Basis of Design (pages) | 45 | 53 | 69 | Low = DD edition; Expected = Final edition; FC 1-300-09N |
| Calculations (pages) | 170 | 227 | 306 | Electrical studies ≈ 75% of volume |

Expected drawings by series: G 2, V 2, C 4, S 3, M 5, FP 1, E 19 = **36**. Spread 1.34×. SCHEMA §6.2 flags anything under 1.5× before DD as possible false precision. It is accepted here, for the engineer to confirm, because the concept drawings E-101–E-108 already fix the SLD and site content. Expect the range to tighten after the site visit and PPI responses.

## Seal floor and merge groups

| Series | Seal holder | Merge |
|---|---|---|
| G | Lead DOR — BFA Electrical PE | — |
| V | Land Surveyor (sub) | — |
| C | Civil PE (sub) | — |
| S | Structural PE (sub) | — |
| M | BFA Mechanical PE (M/P merged) | P merged into M (same seal holder) |
| FP | QFPE (sub) | — |
| E | BFA Electrical PE | — |

Disciplines are kept separate because NAVFAC Part 2 requires each sheet to be sealed by the DOR registered in that discipline (a vehicle rule), not for legibility. BFA issues G, M and E, so the MEP subtotal is not the whole set. The non-MEP sheets (V, C, S, FP) are sealed by subconsultants.

## Sheet list (internal)

| Sheet | Title | Series | Seal holder | Tag | Justification | Driver | MEP |
|---|---|---|---|---|---|---|---|
| G-001 | Cover sheet, location/vicinity maps, sheet index | G | Lead DOR — BFA Electrical PE | vehicle_mandated | FC 1-300-09N drawing set requirements; NAVFAC title block |  | N |
| G-002 | General notes, abbreviations, code summary, criteria | G | Lead DOR — BFA Electrical PE | vehicle_mandated | 01 33 10.05 20 1.2 criteria list; UFC 1-200-01 |  | N |
| V-101 | Topographic and utility survey — Substation #1 | V | Land Surveyor (sub) | vehicle_mandated | Part 3 G10: topo survey shall be included in design drawings |  | N |
| V-102 | Topographic and utility survey — Substation #6 | V | Land Surveyor (sub) | vehicle_mandated | Part 3 G10 |  | N |
| C-001 | Civil general notes, legend, ESC narrative | C | Civil PE (sub) | vehicle_mandated | Part 4 G10.1/3; UFGS 31 23 00.00 20 |  | N |
| C-101 | Site, grading and access plan — Substation #1 (asphalt, open) | C | Civil PE (sub) | scope_mandated | Part 3 3.1, G1030 (grade to drain); trailer access |  | N |
| C-102 | Site, grading and access plan — Substation #6 (concrete, fenced) | C | Civil PE (sub) | scope_mandated | Part 3 3.1, G1030; fence/gate modifications for trailer entry |  | N |
| C-501 | Civil details — pavement repair, bollards, ESC | C | Civil PE (sub) | design_implied | Asphalt cut/patch at Sub #1; bollard protection of JB | RD-3 | N |
| S-001 | Structural general notes, design criteria (120 mph, CVSZ seismic) | S | Structural PE (sub) | vehicle_mandated | Part 3 1.11.17, 1.7.19; UFC 3-301-01 |  | N |
| S-101 | Foundation plans — trailer/tank pads, JB and DEF pads (both sites) | S | Structural PE (sub) | scope_mandated | Part 3 1.18.1.1.1 tank secured to a concrete pad | RD-2 | N |
| S-501 | Foundation details — pads, jacking-leg bearing, tie-downs, anchorage | S | Structural PE (sub) | scope_mandated | Part 3 1.7.19 seismic; 1.11.17 wind |  | N |
| M-001 | Mechanical legend, notes, equipment schedules (DEF, pumps, tanks) | M | BFA Mechanical PE (M/P merged) | vehicle_mandated | FC 1-300-09N; schedules for DEF equipment |  | Y |
| M-101 | Mechanical site plan — Substation #1 (DEF, fuel fill, purge/air provisions) | M | BFA Mechanical PE (M/P merged) | scope_mandated | Sol. p.1 DEF storage/supply/containment; Part 3 1.2.4 | RD-7 | Y |
| M-102 | Mechanical site plan — Substation #6 | M | BFA Mechanical PE (M/P merged) | scope_mandated | Sol. p.1 DEF; Part 3 1.2.4 | RD-7 | Y |
| M-201 | Flow diagrams — DEF and off-skid fuel/air/purge systems | M | BFA Mechanical PE (M/P merged) | scope_mandated | Part 3 1.18, 1.35 fuel controls; DEF level monitoring |  | Y |
| M-501 | Mechanical details — containment, tank, piping, supports | M | BFA Mechanical PE (M/P merged) | scope_mandated | 40 CFR 112 / 9VAC25-91 secondary containment; NFPA 30 |  | Y |
| FP-001 | Fire protection code analysis — NFPA 37 separation, NFPA 12/30 summary | FP | QFPE (sub) | vehicle_mandated | UFC 3-600-01 QFPE; Part 3 1.27; FP is a DOR-approved submittal |  | N |
| E-001 | Electrical legend, symbols, abbreviations, general notes | E | BFA Electrical PE | vehicle_mandated | FC 1-300-09N |  | Y |
| E-002 | Conduit and cable schedules; quick-connect connector schedule | E | BFA Electrical PE | scope_mandated | E-101/E-102 note 3; Part 3 1.2.6/1.2.7 |  | Y |
| E-101 | Electrical site plan, new work — Substation #1 | E | BFA Electrical PE | scope_mandated | Part 6 E-101 notes 1–4 |  | Y |
| E-102 | Electrical site plan, new work — Substation #6 | E | BFA Electrical PE | scope_mandated | Part 6 E-102 notes 1–4 |  | Y |
| E-111 | Grounding plan — Substation #1 | E | BFA Electrical PE | scope_mandated | Part 3 1.34.11.6.2.b (≤25 Ω, IEEE 81); D5010 | RD-7 | Y |
| E-112 | Grounding plan — Substation #6 | E | BFA Electrical PE | scope_mandated | Part 3 1.34.11.6.2.b | RD-7 | Y |
| E-201 | Single-line diagram — Substation #1 with GTG | E | BFA Electrical PE | scope_mandated | Part 6 E-105 |  | Y |
| E-202 | Single-line diagram — Substation #6 with GTG | E | BFA Electrical PE | scope_mandated | Part 6 E-106 |  | Y |
| E-203 | Three-line, metering and relaying diagram — Substation #1 | E | BFA Electrical PE | scope_mandated | E-105 87B bus differential, 7SJ80, SEL-700G |  | Y |
| E-204 | Three-line, metering and relaying diagram — Substation #6 | E | BFA Electrical PE | scope_mandated | E-106 |  | Y |
| E-205 | GTG–switchgear interconnection and control block diagram | E | BFA Electrical PE | scope_mandated | Part 3 1.21, 1.24; black start ATS 1.2.5 |  | Y |
| E-206 | Control network / SCADA architecture and cyber boundary | E | BFA Electrical PE | scope_mandated | Part 3 2.4 (NIST 800-82); 1.21 Modbus TCP |  | Y |
| E-207 | Switchgear elevations and modifications — Substation #1 | E | BFA Electrical PE | scope_mandated | E-103 'Feeder Future' cell conversion |  | Y |
| E-208 | Switchgear elevations and modifications — Substation #6 | E | BFA Electrical PE | scope_mandated | E-104; spare cell unconfirmed | RD-1 | Y |
| E-209 | Sequence of operations and I/O summary | E | BFA Electrical PE | scope_mandated | Part 3 1.26 modes of operation; 1.21 I/O list for SCADA |  | Y |
| E-501 | Electrical details — JB, receptacles, conduit supports, penetrations | E | BFA Electrical PE | scope_mandated | E-101 note 2/3; stainless hardware G4010 |  | Y |
| E-502 | Grounding and NGR details | E | BFA Electrical PE | scope_mandated | Part 3 1.20.3 NGR; IEEE 142 |  | Y |
| E-601 | Protective device settings and coordination summary | E | BFA Electrical PE | scope_mandated | Part 3 1.34.11.11.1 relay settings per coordination study |  | Y |
| E-602 | Short-circuit / arc-flash summary and label schedule | E | BFA Electrical PE | vehicle_mandated | UFC 3-501-01, UFC 3-560-01, NFPA 70E |  | Y |
| C-111 | Erosion and sediment control plan — both sites | C | Civil PE (sub) | design_implied | VA ESC/CBPA: land disturbance ≥2,500 SF in a CBPA area | **Conditional — RD-3** | N |
| FP-101 | Fire protection site plans — separation distances, both sites | FP | QFPE (sub) | design_implied | NFPA 37 separation to Bldgs 115/248 and 1500 | **Conditional — RD-5** | N |
| E-301 | Underground duct bank plan and sections (if underground route) | E | BFA Electrical PE | design_implied | Part 4 D5010.1 underground Sch 40 PVC | **Conditional — RD-4** | Y |
| E-503 | Utility-parallel interconnection protection details | E | BFA Electrical PE | design_implied | Part 3 1.3.17 grid-parallel operation | **Conditional — RD-6** | Y |

## Count range by series (internal)

| Series | Low | Expected | High |
|---|---|---|---|
| G | 2 | 2 | 3 |
| V | 2 | 2 | 2 |
| C | 3 | 4 | 5 |
| S | 2 | 3 | 4 |
| M | 4 | 5 | 5 |
| FP | 1 | 1 | 2 |
| E | 18 | 19 | 22 |
| **Total** | 32 | 36 | 43 |

## Range drivers (internal)

| ID | Driver | Affects | Resolved by | Swing (+/−) |
|---|---|---|---|---|
| RD-1 | Does Sub #6 have a usable spare cubicle, or is a new switchgear section needed? | E-208 (+1 new-section elevation) | PPI-15 / site visit | +1 / 0 |
| RD-2 | Concrete trailer pad required at both sites, or tank/JB pads only? | S-101 splits per site (+1), S-501 | PPI-13 | +1 / −1 |
| RD-3 | Land disturbance ≥2,500 SF in CBPA (dedicated ESC plan)? | C-111 (conditional, +1); C-501 (−1 if ESC rides on C-101/102) | PPI-23 / civil | +1 / −1 |
| RD-4 | Underground duct bank vs above-grade conduit? | E-301 (conditional) | PPI-14 | +1 / 0 |
| RD-5 | Do NFPA 37 separations need dedicated plans? | FP-101 (conditional) | QFPE at site visit | +1 / 0 |
| RD-6 | Utility-parallel operation required? | E-503 (conditional) | PPI-16 | +1 / 0 |
| RD-7 | Separate per-site sheets vs combined plans | M-101/102 (−1), E-111/112 (−1) | Production judgement | 0 / −2 |
| RD-8 | Cover/general split into a third G sheet (criteria + PROD summary) | G-003 | Production judgement | +1 / 0 |

Reconciliation: the positive swings sum to +7, against high − expected = 7. The negative swings sum to −4, against expected − low = 4. **Dominant driver: RD-1 / PPI-15 (switchgear spare cell at Sub #6).** If no spare exists, a new switchgear section adds a sheet, ≈40 engineering hours and ≈$150–300k of construction cost. Settle it at the 17 Nov site visit.

## Specifications (internal)

| Section | Title | Discipline | Tier | Est. pages |
|---|---|---|---|---|
| 00 01 10 | Specification cover (sealed) and table of contents | All | Package | 4 |
| 01 31 23.13 20 | eCMS — Part 5 prescriptive, unmodified | Div 01 | Front-end | 12 |
| 01 32 17.00 20 | Cost-loaded NAS — Part 5 prescriptive, unmodified | Div 01 | Front-end | 30 |
| 01 33 29 | Sustainability Requirements and Reporting — Part 5 prescriptive, unmodified | Div 01 | Front-end | 25 |
| 01 35 26 | Governmental Safety Requirements — Part 5 prescriptive, unmodified | Div 01 | Front-end | 30 |
| 01 11 00 | Summary of Work | Div 01 | Front-end | 4 |
| 01 14 00 | Work Restrictions (outage and tie-in sequencing) | Div 01 | Front-end | 6 |
| 01 33 00.05 20 | Construction Submittal Procedures | Div 01 | Front-end | 14 |
| 01 45 00.05 20 | Quality Control | Div 01 | Front-end | 20 |
| 01 57 19 | Temporary Environmental Controls | Div 01 | Front-end | 14 |
| 01 74 19 | Construction Waste Management and Disposal | Div 01 | Front-end | 8 |
| 01 78 00 | Closeout Submittals | Div 01 | Front-end | 8 |
| 01 78 23 | Operation and Maintenance Data | Div 01 | Front-end | 16 |
| 03 30 00 | Cast-in-Place Concrete | Structural | Technical | 22 |
| 05 05 23.16 | Structural Welding (referenced by Part 3 1.7.15) | Structural | Technical | 10 |
| 05 50 13 | Miscellaneous Metal Fabrications (access platforms, bollards) | Structural | Technical | 8 |
| 13 48 73 | Seismic Control for Nonstructural Components | Structural | Technical | 10 |
| 31 00 00 / 31 23 00.00 20 | Earthwork; Excavation and Fill | Civil | Technical | 18 |
| 31 32 11 | Soil Surface Erosion Control (conditional, RD-3) | Civil | Technical | 6 |
| 32 12 16 | Asphalt Paving (pavement repair, Sub #1) | Civil | Technical | 10 |
| 22 05 00 / 22 11 00 | Common Work Results for Plumbing; Facility Water Distribution (purge supply, backflow) | Plumbing | Technical | 14 |
| 23 05 48.19 | Seismic Bracing for Mechanical Systems | Mechanical | Technical | 8 |
| 23 05 53 | Identification for Mechanical Piping and Equipment | Mechanical | Technical | 4 |
| 23 13 xx (custom) | Diesel Exhaust Fluid Storage and Supply System | Mechanical | Technical | 10 |
| 33 56 10 (edited) | Factory-Fabricated Fuel Storage Tanks — trailer tanks, appurtenances, tightness test | Mechanical | Technical | 16 |
| 40 05 13 | Common Work Results for Process Piping (fuel/DEF field piping) | Mechanical | Technical | 10 |
| (within 26 32 15) | CO2 total-flooding suppression — NFPA 12 requirements carried in GTG section | Fire Protection | Technical | 0 |
| 25 05 11 | Cybersecurity for Facility-Related Control Systems | Cyber | Technical | 22 |
| 26 05 00.00 40 | Common Work Results for Electrical | Electrical | Technical | 10 |
| 26 05 13.00 40 | Medium-Voltage Cables (quick-connect power cables, terminations) | Electrical | Technical | 14 |
| 26 05 19.00 10 | Insulated Wire and Cable (LV, control) | Electrical | Technical | 10 |
| 26 05 26 | Grounding and Bonding | Electrical | Technical | 8 |
| 26 05 33 | Raceways and Boxes (metal conduit, JB) | Electrical | Technical | 10 |
| 26 05 48 | Seismic Protection for Electrical Equipment | Electrical | Technical | 8 |
| 26 08 00 | Apparatus Inspection and Testing | Electrical | Technical | 22 |
| 26 13 xx (MV switchgear) | Medium-Voltage Metal-Clad Switchgear — modifications and new breaker (UFGS number per DOR) | Electrical | Technical | 20 |
| 26 20 00 | Interior Distribution System (aux panelboard, black-start receptacle) | Electrical | Technical | 14 |
| 26 28 01.00 10 | Coordinated Power System Protection | Electrical | Technical | 12 |
| 26 32 15.00 (edited for GTG) | Mobile Gas Turbine Generator Sets (Part 3 D5090 content; Part 4 D5090.3) | Electrical | Technical | 75 |
| 26 36 23 | Automatic Transfer Switches (black-start ATS, package) | Electrical | Technical | 14 |
| 26 41 00 | Lightning Protection (conditional) | Electrical | Technical | 8 |
| 33 71 02 | Underground Electrical Distribution (conditional, RD-4) | Electrical | Technical | 22 |
| — | Submittal register (SpecsIntact-generated) | All | Package | 12 |
| — | Annotated manufacturer product data (GTG, switchgear, relays, DEF) | All | Package | 120 |
|  | **Total** |  |  | 738 |

On-drawing specifications are not accepted (01 33 10.05 20 1.3.e), so specification effort is priced separately.

## Basis of Design (internal)

| Chapter | DD pages | Final pages |
|---|---|---|
| General: project description, criteria, PROD summary, deviations | 4 | 5 |
| Civil (site, grading, ESC, CBPA) | 3 | 3 |
| Geotechnical summary (report appended) | 2 | 2 |
| Structural (pads, wind/seismic anchorage) | 3 | 4 |
| Mechanical (DEF, fuel, containment, purge/air) | 5 | 6 |
| Plumbing (purge water, backflow) | 1 | 1 |
| Fire protection (NFPA 37/12/30, QFPE) | 4 | 4 |
| Electrical (interconnection, protection, grounding, studies summary) | 10 | 12 |
| Controls and cybersecurity (SCADA, RMF approach) | 4 | 5 |
| Environmental and air permitting | 3 | 3 |
| Appendix: equipment data, PROD form, DD 1354 data | 6 | 8 |
| **Total** | 45 | 53 |

## Calculations (internal)

| Discipline | Calculation | Est. pages |
|---|---|---|
| Electrical | Short-circuit study (both substations, GTG and utility sources) | 40 |
| Electrical | Protective device coordination study and relay settings (7SJ80, SEL-700G, 87B) | 60 |
| Electrical | Arc-flash hazard analysis and label data | 30 |
| Electrical | Load flow / voltage drop / cable sizing (MV and LV, black start) | 18 |
| Electrical | NGR, CT/PT and aux-transformer sizing | 10 |
| Electrical | Grounding (ground resistance, step/touch screening) | 10 |
| Mechanical | DEF consumption, storage and runtime; fuel runtime verification | 6 |
| Mechanical | Secondary containment volume (40 CFR 112, 110%) and piping sizing | 6 |
| Plumbing | Purge-water demand and backflow device sizing | 3 |
| Structural | Pad design, bearing, jacking legs, tie-downs (wind/seismic) | 30 |
| Civil | Drainage and ESC computations | 8 |
| Fire Protection | NFPA 37 separation and exposure analysis | 6 |
|  | **Total** | 227 |
