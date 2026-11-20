# Technical Narrative — A/E Design Services

**Project:** FY27 CIP MINCON: Substation Generators, NSF Dahlgren, VA
**Solicitation:** N4008027R0004 · NAVFAC Project 1921806
**Designer of Record:** Bowman, Foster & Associates, Inc. (BFA) · **DB Offeror:** Encon Corporation
**Date:** 9 Nov 2026 · Working draft

---

## 1. Scope Understanding

### 1.1 Project

| Item | Requirement | RFP ref |
|---|---|---|
| Equipment | 2 × 3,400 kWe, 4,160 V, 3-ph, 60 Hz trailer-mounted mobile gas turbine generators (GTG), liquid fuel (No. 2 FO) | Sol. p.1; Part 3 1.0 |
| Locations | 1 GTG at Substation #1 (asphalt, open); 1 GTG at Substation #6 (concrete pad, fenced) | Part 3 3.1, D5090; E-101/E-102 |
| Package | Turbine, gearbox, generator, MV switchgear, NGR, aux transformer 4,160–480 V, MCC, black-start ATS, PLC, HMI, DECS-250, SEL-700G, DEIF AGC 200, easYgen 3200XT, SEL-735, F&G detection, CO2 suppression, enclosure, 8 × 550 gal SS double-wall tanks | Part 3 D5090, 1.12–1.35 |
| Trailer limit | 48 ft L × 8.5 ft W × 13.5 ft H | Sol. p.1 |
| Site interconnection | Existing 4,160 V substation switchgear → new above-ground junction box → quick-connect power and control cabling in metal conduit → GTG | E-101/E-102 notes 1–4 |
| Switchgear work | New 1,200 A generator breaker, SIPROTEC 7SJ80 relay, bus differential (87B) CTs and relay, 15 kV load-break switch, MOV arresters | E-105/E-106 |
| Emissions | DEF system: storage, supply piping, containment, level monitoring, power/control conduits. NSPS KKKK; BACT; VDEQ | Sol. p.1; Part 3 1.3.19, 1.32 |
| Controls | Modbus TCP; I/O list for SCADA; DoDI 8500.01/8510.01 RMF; NIST SP 800-82 App. G | Part 3 1.21, 2.4 |
| Operating modes | Standby, isochronous island, load-following, baseload, grid-parallel | Part 3 1.3.15–1.3.17, 1.26 |
| Duration | 24 months from award; LDs $4,450/day | 52.211-10/-12 |

### 1.2 Existing conditions (concept drawings)

| Element | Substation #1 | Substation #6 |
|---|---|---|
| Switchgear | 7 sections, 4.16 kV, 1,200 A bus, 25 kA, VCBs; Feeders #1–#9; 3 'Feeder Future' cells | 4 sections, 4.16 kV, 1,200 A, 25 kA; Feeders #1–#6; spare not shown |
| Surface | Asphalt, open | Concrete pad, fenced |
| Adjacent structures | Bldgs 115, 248, 248A; hatched water/wetland area west | Bldgs 1500, 1501, 591, 1520F |
| GTG footprint | Enclosure ≈24'-10½" × 8'-4" on double drop-deck trailer; tank modules 10 ft each end (≈52 ft overall) | Same |

### 1.3 BFA role

- Designer of Record for all site and interconnection design, specifications, BOD, calculations and studies.
- Specifies the OEM GTG package in an edited UFGS 26 32 15.00; reviews and approves GTG shop drawings, factory test reports and site test reports as DOR (01 33 00.05 20 1.1.a).
- The OEM designs the package internals. BFA does not redesign them.
- Evaluation is price only. Design scope is held to the minimum compliant solution.

### 1.4 Open scope conditions (PPIs to submit)

| PPI | Condition | Design impact |
|---|---|---|
| 01 | Subs #1A and #4 named in the Solicitation description only | Site count |
| 07 | DEF/'Tier IV' vs a gas turbine; SCR need undefined | DEF design; package length |
| 11 | Nuclear Class 1E standards (IEEE 308/323/344/387, NQA-1) | Qualification basis |
| 13 | Tank 'secured to concrete pad' vs trailer-mounted tanks | Structural pad scope |
| 14 | Metal conduit above grade vs Part 4 underground Sch 40 PVC | Route, civil scope |
| 15 | Spare cubicle and switchgear OEM at Sub #6 | Retrofit vs new section |
| 16 | Grid-parallel operation | Interconnect protection |
| 17 | RMF authorization path and SCADA owner | Cyber deliverables |
| 19 | 4-wire multi-grounded wye (G4010) vs HRG/NGR (GTG) | Grounding scheme |
| 20 | 480 V black-start source at each site | New LV source |

---

## 2. Technical Approach by Discipline

### 2.1 Electrical (BFA — Lead DOR)

**Criteria:** UFC 3-501-01, UFC 3-560-01, UFC 1-200-01, NFPA 70, NFPA 70E, IEEE C2, IEEE 142, IEEE 81, IEEE C37.20.2, C37.101/102, C57.13, NEMA MG-1.

**Interconnection**

- One generator breaker per substation in an existing spare cell (Sub #1 'Feeder Future'; Sub #6 confirmed at site visit). Breaker 1,200 A, VCB, 25 kA, matched to the existing lineup.
- Bus differential (87B) added. Existing feeder CTs verified for 87B zone; new CTs where class or ratio is inadequate.
- 15 kV-class manual load-break switch and MOV arresters (15 kV MCOV) at the junction box per E-105/E-106.
- Above-ground stainless junction box (G4010) with dead-break MV receptacles for GTG output, control/communication receptacles, ground bus and grounding receptacle. Tool-free quick connection per Part 3 1.2.6/1.2.7.
- MV output cable from switchgear to JB in metal conduit (route type per PPI-14). Portable trailer-to-JB cables provided with the package.
- 480 V black-start source: new circuit from substation station service or a new pad-mounted source to a 480 V receptacle within 50 ft of trailer park position (Part 3 1.2.4.1). Source identified at site visit (PPI-20).

**Protection and studies**

| Study | Method | Output |
|---|---|---|
| Short circuit | IEEE 551 / ANSI; utility + GTG sources; both substations | Duty check of existing 25 kA gear; CT saturation check |
| Coordination | Time-current curves: utility upstream, 7SJ80 feeders, 87B, SEL-700G, GTG breaker, NGR | Relay settings files (7SJ80, SEL-700G, 87B); E-601 |
| Arc flash | IEEE 1584-2018 | Incident energy and labels, E-602 |
| Load flow / voltage drop | Steady state, island and parallel; black-start feeder | Cable sizing; aux transformer tap |
| NGR / CT / PT | IEEE C62.92.2, C57.13, C57.32 | NGR rating (A, s), CT class |
| Grounding | IEEE 80 screening; IEEE 81 test criteria ≤25 Ω | Ground ring and rods at trailer park position |

- System grounding: confirm existing 4.16 kV grounding (PPI-19). If solidly grounded wye, the GTG connects through its HRG/NGR with 100% stator ground protection and 87B/51N coordination.
- Grid-parallel: if confirmed (PPI-16), add 32, 81O/U, 27/59, 25 sync check and anti-islanding logic, and coordinate with the utility service owner.

**Controls and SCADA**

- Interconnection block diagram, sequence of operations (Part 3 1.26 modes) and I/O summary (E-205, E-209).
- Modbus TCP from the package PLC/HMI to the installation SCADA. Network architecture and cyber boundary on E-206, coordinated with the cyber subconsultant.
- Hard-wired permissives between the GTG breaker, 87B lockout (86) and the generator controller.

**Specifications:** edited UFGS 26 32 15.00 for the gas-turbine prime mover (all Part 3 D5090 requirements, testing 1.34, submittals 1.6); 26 05 xx; 26 08 00; 26 13 xx; 26 20 00; 26 28 01.00 10; 26 36 23; 26 41 00 (if required); 25 05 11; 33 71 02 (if underground).

### 2.2 Mechanical (BFA)

**Criteria:** NFPA 30, NFPA 37, 40 CFR 112, 9VAC25-91, ASME B31.1/B31.3, UFC 3-401-01 (as applicable).

- **DEF system (if SCR confirmed, PPI-07):** site DEF tank (polyethylene or stainless, heat-traced for 13 °F minimum), secondary containment ≥110% of the largest tank, transfer pump, stainless or HDPE supply piping to the trailer connection, quick-disconnect coupling, level transmitter to the GTG PLC, dedicated power and control conduits. Tank sized for DEF consumption at full load × 14.6 h plus margin.
- **Fuel:** tanks are trailer-mounted (OEM). BFA specifies tank appurtenances, overfill alarm, spill bucket, interstitial leak detection, tightness test (NFPA 30 21.5.2), fill-point access platform (<4 ft rule) and labeling. Runtime basis per PPI-08.
- **Atomizing air and purge water (PPI-02):** if site-supplied, provide the connection point, isolation, filtration to the OEM specification and a backflow preventer (Part 2 01 50 00 1.2.h).
- **Exhaust:** verify plume discharge height and clearance to Bldgs 115/248 and 1500 air intakes and openings; coordinate with the QFPE on NFPA 37 separation.
- **Piping tests:** flush and pressure test per Part 3 1.34.11.6.1 (150% / ≥150 psi, 2 h).
- **Specifications:** 23 13 xx DEF (custom), 33 56 10 (edited), 40 05 13, 23 05 48.19, 23 05 53.

### 2.3 Plumbing (BFA)

- Purge-water supply from the nearest potable source with RPZ backflow preventer, freeze protection and isolation, only if required by PPI-02.
- Containment drainage: manual valved drain with oil-sheen inspection; no connection to storm without approval.
- Specifications: 22 05 00, 22 11 00.

### 2.4 Fire Protection (QFPE subconsultant)

- UFC 3-600-01 QFPE review. Code analysis sheet FP-001.
- NFPA 37 separation of the GTG enclosure and fuel tanks from Bldgs 115, 248, 248A (Sub #1) and 1500, 1501 (Sub #6), plus exposure analysis.
- NFPA 12 CO2 system: design concentration 37% for 20 min, extended discharge, isolation dampers, manual release, warning signage, life-safety lockout. Full discharge test procedure (PPI-24).
- NFPA 30 tank and fuel review. Review of F&G detection (MIR, thermal, IR gas) and backup shutdown.
- DOR approval of FP construction submittals; attend the NAVFAC FPE meeting at PAK.

### 2.5 Structural (subconsultant)

- **Criteria:** UFC 3-301-01; ASCE 7 (120 mph design wind per Part 3 1.11.17; Central Virginia Seismic Zone); ACI 318.
- **Trailer park pad:** reinforced concrete pad sized for GTG (≈32,000 lb) + full fuel (4,400 gal ≈ 31,000 lb) + trailer, with jacking-leg point loads and leveling tolerance. Asphalt removal at Sub #1. At Sub #6, confirm capacity of the existing pad (PPI-13).
- **Restraint:** tie-down anchors for wind and seismic (13 48 73), designed for removable connection so the unit stays mobile.
- **Foundations:** junction box, DEF tank and containment, black-start receptacle pedestal, conduit supports, bollards.
- **Specifications:** 03 30 00, 05 05 23.16, 05 50 13, 13 48 73.

### 2.6 Civil (subconsultant)

- Site, grading and access plans per site: trailer delivery path and turning (WB-62-class vehicle), gate modifications at Sub #6, positive drainage away from the pad (G1030).
- Pavement removal and repair at Sub #1. Limit of work and laydown.
- Land disturbance check against the CBPA 2,500 SF threshold. ESC plan and VSMP coordination if exceeded; RPA buffer and wetland screen at Sub #1 (PPI-23).
- **Specifications:** 31 00 00 / 31 23 00.00 20, 32 12 16, 31 32 11 (if required).

### 2.7 Geotechnical (subconsultant)

- Two borings per site (pad and JB/DEF area), 15–20 ft depth. Lab classification, bearing capacity, groundwater, pavement section.
- Report sealed and submitted as a DOR-approved geotechnical submittal (01 33 00.05 20 1.1.a.2).

### 2.8 Survey / SUE (subconsultant)

- Topographic survey of each site at 1" = 20' (Part 3 G10), tied to installation datum.
- SUE Quality Level B designation of utilities in the conduit route and pad area. Locate existing switchgear, duct banks, grounding, fence and buildings.

### 2.9 Environmental and Air Permitting (subconsultant)

- Air permit applicability: 2 × ≈45 MMBtu/h liquid-fuel turbines. NSPS KKKK NOx limit 96 ppm @15% O2 (non-gas fuel); BACT; VDEQ minor NSR or exemption determination with the NSF Dahlgren Air Program Manager. Application support and an emissions test protocol.
- SPCC plan amendment (40 CFR 112) for 8,800 gal of new aboveground oil storage. 9VAC25-91 AST applicability.
- PROD form with the DD submittal (01 30 00 1.6).

### 2.10 Controls / Cybersecurity (subconsultant)

- Categorize the FRCS with the installation ISSM (PPI-17); select NIST SP 800-82 App. G baseline controls.
- Network architecture, boundary protection for Modbus TCP, PLC/HMI/relay hardening, account management, logging.
- Artifacts for eMASS: system description, hardware/software list, ports-protocols-services, control implementation statements. UFGS 25 05 11 edit.

---

## 3. Deliverables

### 3.1 Design submittals

| # | Submittal | Timing (from award) | Content |
|---|---|---|---|
| 0 | Concept presentation at PAK | ≤ 21 d | Concept SLDs, site layouts, interconnection approach, PPI resolutions; NAVFAC FPE meeting |
| 1 | Design Development | ≈ 75 d | BOD; half-size drawings; online specs; calcs; DQC docs; PROD form; submittal register |
| 2 | Final Design | ≈ 135 d | BOD; drawings; specs + annotated cut sheets; PIF/bid schedule; calcs; DQC docs; register |
| 3 | Released for Construction | ≈ 170 d | Signed and sealed final; comment responses closed |
| Opt. | Early OEM release | At PAK | Edited 26 32 15.00 + SLD for long-lead GTG order |

Government review: 21 days before each in-progress review (01 33 10.05 20 1.3.j).

### 3.2 Deliverable quantities (expected)

| Deliverable | Quantity |
|---|---|
| Drawings | 36 sheets (range 32–43); 7 seal holders; MEP 24 |
| Specifications | ≈ 738 pages, 44 line items, SpecsIntact, sealed cover |
| Basis of Design | 45 pages (DD) / 53 pages (Final) |
| Calculations and studies | ≈ 227 pages |
| Relay settings files | 7SJ80 (×2 sites), 87B (×2), SEL-700G review (×2) |
| Other | PROD form; submittal register (Excel/SpecsIntact); DD 1354 asset data; DQC records; comment response matrices (3) |

### 3.3 Drawing index (expected)

| Series | Sheets | Content |
|---|---|---|
| G | G-001, G-002 | Cover, index, general notes, criteria |
| V | V-101, V-102 | Topographic and utility survey, each site |
| C | C-001, C-101, C-102, C-501 | Notes, site/grading/access plans, details |
| S | S-001, S-101, S-501 | Criteria, foundation plans, pad and tie-down details |
| M | M-001, M-101, M-102, M-201, M-501 | Schedules, site plans, flow diagrams, details (P merged) |
| FP | FP-001 | Code analysis, NFPA 37/12/30 |
| E | E-001, E-002, E-101, E-102, E-111, E-112, E-201–E-209, E-501, E-502, E-601, E-602 | Legend, schedules, site plans, grounding, SLDs, three-lines, control/network, switchgear elevations, sequence/I-O, details, settings, arc flash |

### 3.4 Design meetings

5 formal (PAK/concept, FPE, DD IPR, Final IPR, RFC closeout) and 8 working (6 biweekly design coordination, 2 RMF sessions).

---

## 4. Assumptions

1. Scope is limited to Substations #1 and #6, one GTG each (PPI-01).
2. Each GTG connects through one new generator breaker in an existing spare cell of the existing switchgear. No new switchgear lineup (PPI-15).
3. Existing switchgear is rated for the added fault contribution of one 4.4 MVA GTG. Replacing switchgear to cure a duty deficiency is a change.
4. Operation is islanded on the substation bus with closed-transition capability only. Utility-parallel export and interconnection studies with the serving utility are excluded unless confirmed (PPI-16).
5. Conduit route is above grade in metal conduit per E-101/E-102, ≈250 LF per site (PPI-14).
6. Nuclear Class 1E standards and qualification criteria (IEEE 308/323/344/387/603, NQA-1) do not apply (PPI-11).
7. The OEM provides package design, package drawings, package calculations (torsional, BMEP, capacity, battery, instrument), F&G and CO2 design, and factory testing.
8. The DEF system is designed only if SCR is confirmed. Otherwise DEF content is deleted (PPI-07).
9. Existing utility and record drawings for both substations (switchgear manufacturer drawings, relay settings, grounding, station service) are provided by the Government at award.
10. One design site visit per discipline after award, in addition to the 17 Nov 2026 proposal site visit.
11. Government review periods are 21 days, with no more than two review cycles per submittal.
12. The DQC Manager and the QC reviewer are BFA staff independent of the design team.
13. Air permit application is support only. Encon or the Navy is the permittee; permit fees are paid by Encon.
14. RMF scope is limited to design-phase artifacts and controls specification. The ATO decision and ISSM effort are by the Government.
15. Design follows criteria current at award; no re-design for criteria changes after award.

## 5. Exclusions

1. Design of OEM package internals: turbine, gearbox, generator, enclosure, F&G, CO2 system hydraulics, PLC programming, trailer.
2. Construction-phase services: submittal reviews, RFIs, site observation, factory and site test witnessing, weekly QC meetings, record-drawing CAD updates. These are offered as a separate line item.
3. Commissioning authority services (not required by RFP).
4. NETA field testing, relay setting implementation and emissions source testing (construction trades); BFA provides settings and specifications.
5. Utility interconnection application, studies or upgrades with the serving utility.
6. Hazardous material surveys and abatement; wetland delineation and USACE/VMRC permitting if required.
7. SUE Quality Level A test pits.
8. Replacement of existing switchgear, transformers or upstream distribution.
9. LEED or other third-party sustainability certification; HPSB checklist beyond design-phase data.
10. Cost estimating, scheduling (NAS) and schedule of prices preparation (by Encon).
11. Permit, plan review and application fees.
12. Security clearances above base access (DBIDS).
