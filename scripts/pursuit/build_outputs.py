#!/usr/bin/env python3
"""scripts/pursuit/build_outputs.py — 27001D (N4008027R0004 Substation Generators)

Builds the data-driven proposal outputs into a staging folder
(``_scratch/stage`` by default):

    artifacts/*.md              shred, disciplines, schedule, ROM, fee, counts, submittals
    published/*.xlsx            same content as live-formula workbooks
    published/calendar/*.ics    Outlook calendar events

Hand-written artifacts (technical narrative, executive summary) are placed in
``_scratch/stage/artifacts`` and converted with ``--docx`` (pandoc +
templates/docx_reference/reference.docx). Publishing (versioning + git) is
done by scripts/publish/publish_outputs.py.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts" / "lib"))
from guard_cwd import require_proposal_or_engine_root  # noqa: E402

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

P = "27001D"
TITLE = "FY27 CIP MINCON: Substation Generators — NSF Dahlgren, VA"
SOL = "N4008027R0004 (NAVFAC Project 1921806 / PID N40080-27-RFPREQ-SP DAHL ROICC-0004)"
TODAY = date(2026, 11, 9)
SCHED_ASOF = date(2026, 11, 24)  # schedule status date; bump when statuses change
ENCON_SRC = "Encon ITB email, T. Colvin, 5 Nov 2026 (rfp_intake/original_rfp/*.msg)"
HDR = (f"**Pursuit:** {P} · **Solicitation:** {SOL}  \n**Prime A/E (DOR):** Bowman, Foster & Associates, Inc. (BFA) · "
       f"**DB Offeror (GC):** Encon Corporation  \n**Prepared:** {TODAY:%d %b %Y} · **Status:** Working draft for internal review")

# ════════════════════════════════════════════════════════════════════
# 1. RFP SHRED  (id, rfp_ref, requirement, category, responsible, a/e note)
# ════════════════════════════════════════════════════════════════════
SHRED = [
    ("R-001", "Sol. p.1", "All MACC holders shall propose or submit a written no-propose determination to the KO within 5 days of RFP receipt.", "Proposal", "GC", "Encon obligation; none for BFA."),
    ("R-002", "Sol. p.4", "Basis of award: price only, lowest evaluated price, FAR 16.505 fair opportunity; award without discussions intended.", "Proposal", "GC", "No technical volume is scored. A/E value = lowest compliant design scope and fee; no gold-plating."),
    ("R-003", "Sol. p.4", "Submit: cover page (solicitation no./title, prime name, address, phone, UEI, CAGE, POC), Price Proposal Form, amendment acknowledgments, bid guarantee.", "Proposal", "GC", "BFA provides fee + scope letter to Encon only."),
    ("R-004", "Sol. p.2", "Electronic bid bond with proposal; hard copy within 1 business day of request. Missing e-bond = non-responsive.", "Proposal", "GC", "—"),
    ("R-005", "Sol. p.3", "Proposal due Wed 16 Dec 2026, 1400 EST, by email to kayla.m.liverman.civ@us.navy.mil; subject line = PID, title, contractor. Late proposals not considered.", "Proposal", "GC", "Encon requires BFA proposal Tue 15 Dec 2026, 1400 (Encon ITB)."),
    ("R-006", "Sol. p.3", "PPIs in writing via email to Kayla Liverman NLT Wed 2 Dec 2026, 1400 EST, as Excel (Attachment C, no tables). No duplicates.", "Proposal", "GC", "Encon internal deadline 1100 EST. BFA PPIs to Encon by 23 Nov (see schedule)."),
    ("R-007", "Sol. p.2", "Site visit Tue 17 Nov 2026, 1300 EST (single visit). Base access forms (SECNAV 5512/1) to Brian Hornaman NLT Mon 9 Nov 2026, 1300 EST.", "Proposal", "BFA/GC", "Encon asked for BFA 5512 by 9 Nov 1000 (password-protected); BFA sent 9 Nov 1046."),
    ("R-008", "Sol. p.4", "Evaluated price = total lump sum incl. all options (FAR 52.217-5). Price Schedule CLIN 0001, firm fixed price, all-or-none.", "Price", "GC", "Single CLIN; no options listed."),
    ("R-009", "Sol. p.1", "Magnitude of construction: $5,000,000–$10,000,000 (FAR 36.204).", "Price", "GC", "ROM expected exceeds upper bound; see ROM."),
    ("R-010", "Part 1", "Bid schedule per Part 6 Attachment 1 (blank Excel). Proposals not accepted without it. Contractor may add DFOWs.", "Price", "GC", "Encon: BFA price breakdown per Attachment 1. A/E fee rolls into 'Engineering and Project Management'."),
    ("R-011", "Part 1", "Overall project schedule: award, design kickoff, concept workshop, all design phases and review times, construction submittals, fabrication, weather days, construction.", "Proposal", "GC", "BFA supplies design-phase durations and review periods."),
    ("R-012", "Sol. p.2", "Proposal acceptance period 120 days from receipt of offers.", "Proposal", "GC", "Expires 15 Apr 2027; BFA fee validity must match."),
    ("R-013", "Sol. p.2 / 52.211-10", "Commence within 15 days of NTP; complete within 24 months of award, incl. 15 days for bonds/insurance.", "Contract", "GC", "GTG lead time is the critical path; design must release OEM shop drawings early."),
    ("R-014", "52.211-12", "Liquidated damages $4,450 per calendar day.", "Contract", "GC", "—"),
    ("R-015", "Sol. p.2", "Performance and payment bonds 20% of offer or $3,000,000, whichever is less, within 10 days of award.", "Contract", "GC", "—"),
    ("R-016", "Sol. p.2", "Davis-Bacon WD VA20260007 (Mod 1, 18 May 2026), Building construction type.", "Contract", "GC", "—"),
    ("R-017", "DFARS 252.227-7023", "All designs, drawings and specifications become Government property; DOR retains copies, furnishes on request for 3 years.", "Contract", "BFA", "Standard; no reuse restriction."),
    ("R-018", "5252.209-9300", "OCI: preparers of SOW/specs are barred from related work.", "Contract", "BFA", "Confirm BFA had no role in RFP preparation."),
    ("R-019", "252.204-7000", "No public release of contract information without KO approval (10 business days).", "Contract", "BFA", "No marketing use of project data without KO approval."),
    ("R-020", "Part 2 1.6", "Warranty 1 year after final acceptance.", "Contract", "GC", "—"),
    ("R-021", "Part 2 01 30 00 1.3", "Post-Award Kickoff (PAK) within 21 days of award: present concept design for acceptance. DOR, DQC Manager and major subs attend. If fire protection work is included, DOR meets the NAVFAC FPE.", "Design", "BFA", "Concept design must exist by about day 21 → BFA concept effort front-loaded. CO2 suppression triggers the FPE meeting."),
    ("R-022", "Part 2 01 33 00 1.3.f / k", "Baseline schedule before PAK; initial schedule of prices due 21 days after award; detailed schedule before construction. Design/construction split required.", "Admin", "GC", "BFA provides design-phase schedule of prices."),
    ("R-023", "Part 2 01 32 17.00 20", "Cost-loaded NAS per UFGS 01 32 17.00 20 (Part 5, prescriptive).", "Admin", "GC", "—"),
    ("R-024", "Part 2 01 31 23.13 20", "eCMS per UFGS 01 31 23.13 20 (Part 5).", "Admin", "GC/BFA", "Design submittals via eCMS."),
    ("R-025", "Part 2 01 45 00 1.1", "QC Plan incl. design QC: names/quals, outside A/E firms, initial design + construction submittal register, testing plan, DFOWs, communication plan (meeting frequency, design decision points, peer review, comment resolution), IBC Ch.17 special inspections.", "QC", "GC/BFA", "BFA writes design sections and communication plan."),
    ("R-026", "Part 2 01 33 10.05 20 1.3.d", "Design effort overseen by a DQC Manager; DQC Plan required. Sealing per FC 1-300-09N.", "QC", "BFA", "QC hours in fee = 15% of engineering."),
    ("R-027", "Part 2 01 45 00 1.3.e", "Weekly QC meetings with DOR, Superintendent and Government technical team, as appropriate to the phase.", "QC", "BFA/GC", "DOR attendance during construction is outside the design-only fee; carry separately."),
    ("R-028", "Part 2 01 35 26 1.1", "Ergonomic design per MIL-STD-1472F for maintenance access.", "Design", "BFA", "Address JB and receptacle heights and access clearances."),
    ("R-029", "Part 2 00 00 00 1.2", "Barrier-free design per UFC 1-200-01.", "Design", "BFA", "Not applicable to an unoccupied outdoor utility installation; state in BOD."),
    ("R-030", "Part 2 01 14 00 1.1", "Third-party utility locates before construction; dig permits 30 days before excavation.", "Site", "GC", "Survey/SUE sub supports design."),
    ("R-031", "Part 2 01 14 00 1.2.h; 01 50 00 1.2.d", "Work hours 0700–1530 M–F; 48-hour notice for off-hours. Utility tie-in/outage notice 15 days.", "Site", "GC", "Design must sequence the switchgear tie-in outage."),
    ("R-032", "Part 2 01 30 00 1.6", "Obtain all permits; submit a complete PROD form with the first design submittal; pay permit fees.", "Permits", "BFA/GC", "BFA prepares the PROD form at DD (VDEQ air, ESC, AST/SPCC)."),
    ("R-033", "Part 2 01 33 10.05 20 1.3.a", "Design documents: BOD, drawings, specifications, reports, submittal register per FC 1-300-09N.", "Design", "BFA", "—"),
    ("R-034", "Part 2 01 33 10.05 20 1.3.e", "Contractor-originated specifications: UFGS section for every product/material/system, CSI MasterFormat, sealed cover sheet, TOC; Part 5 prescriptive sections included unmodified.", "Specs", "BFA", "On-drawing specs not accepted → spec effort priced separately."),
    ("R-035", "Part 2 01 33 10.05 20 1.3.e.6", "Cumulative coordinated submittal register with each design package; DOR identifies DOR-approved items; all fields complete at final.", "Specs", "BFA", "—"),
    ("R-036", "Part 2 01 33 10.05 20 1.3.f", "Annotated manufacturer product data appended to the applicable UFGS section where final selections are made.", "Specs", "BFA/OEM", "GTG cut sheets attached to the GTG section."),
    ("R-037", "Part 2 01 33 10.05 20 1.3.g", "Final spec source in MS Word or SpecsIntact; submittal register in Excel or SpecsIntact.", "Specs", "BFA", "Use SpecsIntact."),
    ("R-038", "Part 2 01 33 10.05 20 1.3.j", "Design submittals: Design Development (BOD, half-size drawings, online specs, calcs, DQC docs) and Final Design (adds cut sheets, PIF with bid schedule). Electronic copies 21 days before the in-progress review meeting.", "Design", "BFA", "Two formal review submittals + released-for-construction set."),
    ("R-039", "Part 2 01 33 10.05 20 1.3.k", "Final design signed and sealed by the DOR before construction start. Separate final packages only if agreed at PAK.", "Design", "BFA", "Propose an early 'site/OEM release' package at PAK for long-lead procurement."),
    ("R-040", "Part 2 01 33 10.05 20 1.3.l", "Submit a draft DD Form 1354.", "Closeout", "GC/BFA", "BFA provides asset data."),
    ("R-041", "Part 2 01 33 10.05 20 1.2", "Comply with UFC 1-200-01, 1-200-02, 1-300-08, FC 1-300-09N, UFC 3-560-01, 3-600-01, 3-810-01N and Core UFCs. Advisory provisions are mandatory.", "Criteria", "BFA", "—"),
    ("R-042", "Part 2 01 33 10.05 20 1.4", "Order of precedence: proposal/final design that exceeds solicitation governs; then SF1442/price schedule/WD, Part 1, Part 2, Part 3, Part 6 (excl. concept dwgs), Part 5, Part 4, concept drawings.", "Criteria", "BFA", "Do not over-promise in proposal text; it becomes contract."),
    ("R-043", "Part 2 01 33 00.05 20 1.1.a", "DOR approval required: fire protection submittals, earthwork/geotechnical, GTG shop drawings, GTG factory test reports, GTG site test reports.", "CA", "BFA", "DOR review of OEM package is a major construction-phase effort (not in design fee)."),
    ("R-044", "Part 2 01 33 00.05 20 1.3.p", "Construction submittals listed in contractor specs; 15-day review assumed; 3 hard copies to ROICC.", "CA", "GC", "—"),
    ("R-045", "Part 3 D5090 1.6", "GTG submittals for Government review and approval: shop drawings (21 items), product data (36), design data (14), test reports (10), certificates (13), manufacturer field reports (3), O&M, closeout.", "CA", "GC/OEM/BFA", "About 100 enumerated GTG submittal items."),
    ("R-046", "Part 2 01 33 00 1.7", "Record drawings: electronic as-builts on contractor DWG files; PDF per FC 1-300-09N.", "Closeout", "GC/BFA", "BFA provides CAD; as-built drafting by others unless added."),
    ("R-047", "Part 3 1.0 / D5090", "Provide and install two (2) trailer-mounted 3,400 kWe, 4,160 V liquid-fuel mobile GTGs: one at Substation #1 and one at Substation #6, with quick-connect power and control cables, terminations, accessories.", "Scope", "GC/OEM/BFA", "Solicitation description also names Substations #1A and #4 → PPI-01."),
    ("R-048", "Part 3 D5090", "Complete package: generator, turbine, MV switchgear, air filtration, starting, controller, protection, NGR, turbine PLC, gen protection relay, DECS, DEIF load management/shedding, dedicated gen PLC, meter, step-down transformer, LV panelboard, HMI, 15 kV MCOV arresters, CTs/PTs, lube, fuel, cooling, enclosure, exhaust.", "Scope", "OEM", "OEM package design; BFA specifies, reviews and integrates."),
    ("R-049", "Part 3 1.1 / 1.2", "Single-frame mobile power plant; hydraulic leveling; park-plug-play; 50 ft black-start cable; plugs/receptacles for high-amp quick, tool-free connection; 480 V black-start ATS; 4,160–480 V aux transformer.", "Electrical", "OEM/BFA", "BFA designs the site side of every connection (black start, output power, controls, ground)."),
    ("R-050", "Part 3 1.3", "Guaranteed power, fuel consumption and NOx/CO emissions at 59 °F, 60% RH, No. 2 fuel oil; part-load data at 25/50/75%; load-following, baseload, isochronous and grid-parallel operation.", "Performance", "OEM", "Grid-parallel operation drives utility-interconnect protection → PPI-16."),
    ("R-051", "Part 3 1.3.19 / 1.32 / 1.34.11.8.6", "Meet NSPS 40 CFR 60 Subpart KKKK and BACT; comply with VDEQ requirements via the NSF Dahlgren Air Program Manager (T. Valentour); manufacturer emissions certification incl. HAPs; field emission testing.", "Environmental", "GC/OEM/BFA", "Air permit applicability determination and support needed (env. sub)."),
    ("R-052", "Part 3 1.3.20 / 1.11", "Design criteria: 98.2 °F DB / 80.8 °F WB max, 13 °F min; 120 mph wind; Central Virginia Seismic Zone; elevation 15 ft (1.3.20) vs 112 ft (1.11).", "Criteria", "BFA", "Elevation conflict → PPI-10."),
    ("R-053", "Part 3 1.5", "Maintenance schedule and spares to support 97% minimum online availability.", "O&M", "OEM", "—"),
    ("R-054", "Part 3 1.7.19", "Seismic per UFC 3-301-01 and UFGS 13 48 73, 23 05 48.19, 26 05 48.", "Structural", "BFA/Sub", "Trailer tie-down and anchorage design at each pad."),
    ("R-055", "Part 3 1.7.17", "Torsional analysis certifying no damaging vibration at ±10% synchronous speed.", "Mechanical", "OEM", "DOR reviews."),
    ("R-056", "Part 3 1.12.14", "Unit to perform as a redundant standby supply per IEEE 308; design-basis-event, 60-year, radiation dose and ECCS criteria.", "Criteria", "OEM", "Nuclear Class 1E carry-over language → PPI-11. Major cost driver if enforced."),
    ("R-057", "Part 3 1.12.15", "Parameter schedule: 4,160 V; 4,434 kVA / 3,547 kW at 0.8 pf; 1,800 rpm; 60 Hz; IP21; DECS; island/parallel/grid; dead-break connectors; high-resistance grounding; trailer 48 × 8.5 × 13.5 ft.", "Electrical", "OEM", "Table labels are misaligned in the RFP → PPI-06."),
    ("R-058", "Part 3 1.18 / 1.35", "Liquid fuel: 8 stainless double-wall UL tanks × 550 gal per GTG (4 on fifth-wheel deck, 4 at rear); overfill prevention, spill bucket, interstitial leak detection; tightness test per NFPA 30; fuel lock (by NSFDL Env.); 14.6 h full-load runtime; carbon absorption unit; spill detector; duplex strainers and 5-µm filter/coalescers; lead/lag pumps.", "Mechanical", "OEM/BFA", "Runtime math does not close (4,400 gal ÷ 350 gal/h = 12.6 h) → PPI-08. Tank 'secured to concrete pad' → PPI-13."),
    ("R-059", "Sol. p.1 / Part 3 1.0", "Provide, install and integrate the DEF system: storage tanks, supply piping, containment, level monitoring, dedicated power/control conduits.", "Mechanical", "GC/BFA", "DEF implies SCR on a gas turbine, which is unusual → PPI-07. BFA designs the site DEF storage, containment and conduits."),
    ("R-060", "Part 3 1.20.2 / 1.20.3", "Generator protection relay (ANSI 59/27/50/51/87/81/40/32/46/24; 100% stator ground); NGR sized to limit L-G fault; CT class per fault and relay studies.", "Electrical", "OEM/BFA", "BFA performs short-circuit, coordination, NGR and CT sizing studies."),
    ("R-061", "Part 3 1.21", "Control system: Modbus TCP; submit I/O list with software addresses for SCADA integration; vibration and ventilation monitoring.", "Controls", "OEM/BFA", "SCADA head-end identity and ownership → PPI-17."),
    ("R-062", "Part 3 2.4", "Cybersecurity: control systems per DoDI 8500.01, DoDI 8510.01 (RMF) and FEC/installation policy; NIST SP 800-82 App. G, Table G-1 controls.", "Cyber", "GC/BFA/Sub", "RMF package scope (ATO/IATT) and ISSM unknown → PPI-17. Cyber sub required."),
    ("R-063", "Part 3 1.27", "Enclosure: 50 psf roof, 120 mph wind; separate turbine and generator compartments; F&G detection; CO2 total flooding + extended discharge per NFPA 12 (37% for 20 min); full discharge test at site.", "Fire Protection", "OEM/FPE", "FP submittals require DOR (QFPE) approval; NAVFAC FPE meeting at PAK."),
    ("R-064", "Part 3 1.31", "Sound limits at 22.9 ft radial (octave bands 31.5 Hz–8 kHz; 60–87 dB).", "Acoustic", "OEM", "Site-adjacent building impacts are not addressed by RFP."),
    ("R-065", "Part 3 1.34", "Field acceptance tests; factory tests incl. power-section, engine assembly, load transient (50/75/100/110%), 150 start-and-load tests with zero failures, margin tests, 22-h continuous + 2-h 110% load capability, electrical, functional, overspeed; procedures 10 days before tests.", "Testing", "OEM/GC", "150-start qualification test is an IEEE 387 nuclear-grade requirement → PPI-12."),
    ("R-066", "Part 3 1.34.11.6", "Construction tests: piping flush and pressure test (150% / ≥150 psi, 2 h); MV cable insulation tests; ground resistance ≤25 Ω by fall-of-potential (IEEE 81); switchgear and breaker testing.", "Testing", "GC", "BFA specifies in 26 08 00 and the piping section."),
    ("R-067", "Part 3 1.34.11.11", "Pre-operational tests: protective relays set per the coordination study; IEEE 43 insulation; coupling alignment. Safety run tests incl. overspeed and day-tank level alarms.", "Testing", "GC/BFA", "Coordination study and relay settings are BFA deliverables."),
    ("R-068", "Part 3 1.36", "On-site O&M training by factory-authorized personnel.", "Closeout", "OEM", "—"),
    ("R-069", "Part 3 G10", "Topographic survey of each site included in the design drawings; physically verify existing utilities.", "Site", "BFA/Sub", "Survey sub; SUE Quality Level B."),
    ("R-070", "Part 3 G1030", "Grade so the site drains; no blasting; surplus and unsuitable material disposed off Government property.", "Civil", "BFA/GC", "—"),
    ("R-071", "Part 4 G10 3", "Geotechnical engineer to perform subsurface exploration the DOR deems necessary (pads, utility structure foundations).", "Geotech", "Sub", "Two-site boring program for trailer/tank pads."),
    ("R-072", "Part 4 G1030", "UFGS 31 23 00.00 20 edited and submitted; PE-designed shoring/dewatering if required.", "Civil", "BFA/Sub", "—"),
    ("R-073", "Part 3 G4010 / Part 4 G40", "Electrical per UFC 3-501-01; primary distribution 4-wire multi-grounded wye; stainless enclosures and hardware outdoors; conduits and cables as shown on Part 6 drawings; underground primary per UFC 3-501-01.", "Electrical", "BFA", "4-wire solidly grounded wye conflicts with the GTG HRG/NGR → PPI-19."),
    ("R-074", "Part 3 D5010", "Insulated equipment grounding conductor in all raceways >50 V.", "Electrical", "BFA", "—"),
    ("R-075", "Part 4 D5090", "Use UFGS 26 32 15.00 (generators), 26 36 23 (ATS), 26 41 00 (lightning), 26 29 23 (VFD) where applicable; submit edited sections with design.", "Specs", "BFA", "Edit 26 32 15.00 for gas-turbine prime mover."),
    ("R-076", "Part 6 E-101/E-102", "Concept site plans, Sub #1 (asphalt, open) and Sub #6 (concrete pad, fenced): existing 4,160 V switchgear (1); new above-ground junction box (2); quick-connect power/control cabling in metal conduit between switchgear and GTG (3); new trailer-mounted GTG (4).", "Electrical", "BFA", "Defines BFA's core site scope."),
    ("R-077", "Part 6 E-103/E-104", "Existing switchgear: Sub #1 7-section, 4.16 kV, 1,200 A bus, 25 kA, VCBs, Feeders #1–#9 with three 'Feeder Future' cells; Sub #6 4-section, Feeders #1–#6.", "Electrical", "BFA", "Confirm spare cubicle availability at Sub #6 → PPI-15."),
    ("R-078", "Part 6 E-105/E-106", "Concept SLDs: new 1,200 A generator breaker in existing switchgear with SIPROTEC 7SJ80 relay and bus differential (87B); 15 kV manual load-break switch; MOV; 45 kVA 4.16 kV–208/120 V aux transformer; SEL-700G, SEL-735, easYgen 3200XT, DECS-250, DEIF AGC 200, PLC.", "Electrical", "BFA/OEM", "Bus differential retrofit and relay settings in BFA design scope."),
    ("R-079", "Part 6 E-107/E-108", "GTG on double drop-deck trailer: enclosure about 24'-10½\" × 8'-4\" × 11'-7\"; fuel tank modules 10'-0\" each end of a 32'-0\" deck.", "Electrical", "OEM", "Overall length about 52 ft vs 48 ft maximum → PPI-09."),
    ("R-080", "Part 3 2.2 / 01 33 29", "Sustainability per UFGS 01 33 29; HPSB checklist attached to interim DD 1354.", "Sustainability", "GC/BFA", "Applicability to a non-building project → PPI-22."),
]

# PPI / conflict log: (id, rfp_ref, issue, impact, proposed question)
PPIS = [
    ("PPI-01", "Sol. p.1 vs Part 3 D5090, E-101/E-102", "Description names Substations #1, #1A, #4 and #6; scope and drawings place one GTG each at #1 and #6 only.", "Number of interconnection sites drives the design and site-work price.", "Confirm interconnection is required only at Substations #1 and #6, and that #1A and #4 are not in scope."),
    ("PPI-02", "Part 3 1.1.8 vs 1.2.4", "1.1.8 says four main connections; 1.2.4 says six and lists five (black start, output, purge water, atomizing air, diesel), citing ES 9-98.", "Purge-water and atomizing-air supplies may need site utilities.", "Confirm required site connections. Are purge water and atomizing air on-board (off-skid module) or site-supplied? Is ES 9-98 (an OEM standard) mandatory?"),
    ("PPI-03", "Part 3 1.1.2, 1.14 vs 1.12.15", "SoLoNOx and 'wide Wobbe range' are gas-fuel/single-OEM terms; 1.14 specifies single-shaft but 1.12.15 says split shaft.", "Possible proprietary specification; affects competition and the design basis.", "Confirm liquid-fuel-only operation, single- vs split-shaft, and that equal products from other OEMs are acceptable."),
    ("PPI-04", "Part 3 1.16.1 vs 2.1 refs", "Generator called 'salient pole' while ANSI C50.13 (cylindrical rotor ≥10 MVA) is referenced.", "Minor; clarify which standard governs.", "Confirm the generator standard (NEMA MG-1 salient pole at 1,800 rpm)."),
    ("PPI-05", "Sol. p.1 / Part 3 1.12.15", "Rating: 3,400 kWe stated; parameter table shows 4,434 kVA / 3,547 kW at 0.8 pf.", "Basis for guarantee testing.", "Confirm guaranteed net output (3,400 kWe net at the generator breaker at 59 °F)."),
    ("PPI-06", "Part 3 1.12.15", "Parameter schedule columns are shifted (e.g., 'Voltage 4434 kVA', 'Shaft speed 60 Hz', 'Heat rate 101,880 lb/hr').", "Ambiguous contract data.", "Request a corrected parameter schedule via amendment."),
    ("PPI-07", "Sol. p.1, Part 3 1.0, 1.12.15", "DEF system and 'Tier IV compliant' are specified. Tier IV applies to nonroad diesel engines; a gas turbine needs DEF only with SCR. NSPS KKKK and BACT are also cited.", "SCR on a mobile 3.4 MW GTG is a major cost, weight and length driver.", "Is SCR required? If so, state the NOx limit. If not, delete the DEF requirement. Who holds the VDEQ permit?"),
    ("PPI-08", "Part 3 1.35 vs 1.12.15", "8 × 550 gal = 4,400 gal per unit; at 350 gal/h this gives 12.6 h, not the 14.6 h required.", "Tank capacity vs runtime conflict.", "Which governs: tank quantity/size or 14.6 h runtime at full load?"),
    ("PPI-09", "Sol. p.1 vs E-107", "Maximum trailer 48 ft L; concept E-107 shows 10 ft + 32 ft + 10 ft ≈ 52 ft with tank modules.", "Trailer configuration and DOT permitting.", "Confirm the 48 ft limit includes the tank modules, or accept the E-107 configuration."),
    ("PPI-10", "Part 3 1.3.20.3 vs 1.11.15", "Site elevation 15 ft MSL vs 112 ft MSL.", "Derate basis (minor).", "Confirm design elevation for each site."),
    ("PPI-11", "Part 3 2.1 (IEEE 308/323/344/384/387/603, NQA-1) and 1.12.14", "Nuclear Class 1E criteria: 60-year qualification, radiation dose, design-basis events, ECCS actuation.", "Qualification to Class 1E is not commercially available for mobile GTGs; large cost and schedule risk.", "Confirm nuclear Class 1E standards and qualification criteria are not applicable (commercial/industrial standby duty)."),
    ("PPI-12", "Part 3 1.34.2.5–1.34.2.7", "150 consecutive start-and-load tests with zero failures, plus a 22-h load test, per unit at the factory.", "Large FAT cost and schedule.", "Is the 150-start reliability test required per unit, or can prototype or type-test data be accepted?"),
    ("PPI-13", "Part 3 1.18.1.1.1", "Fuel tanks must be 'secured to a concrete pad'; tanks are trailer-mounted. Sub #1 is asphalt.", "Defines whether new concrete pads and anchorage are required.", "Is a concrete trailer pad required at each site? Provide any existing pad data for Sub #6."),
    ("PPI-14", "Part 4 D5010.1 vs E-101/E-102 note 3", "Part 4 prohibits exposed exterior conduit and requires underground Sch 40 PVC; concept notes require metal conduit.", "Above-ground vs underground route; trenching scope.", "Confirm the conduit route type (above-grade RGC on supports, or underground duct bank)."),
    ("PPI-15", "E-103/E-104 vs E-105/E-106", "New 1,200 A generator breaker and bus differential shown in existing switchgear. Sub #1 has 'Feeder Future' cells; Sub #6 shows no spare.", "Switchgear retrofit vs new section; OEM involvement; outage.", "Identify the existing switchgear manufacturer/model and the spare cubicle to be used at each site. Are bus differential relays and CTs new?"),
    ("PPI-16", "Part 3 1.3.17", "Operation 'synchronized to the utility grid' / in parallel with the local utility.", "Utility interconnection study and protection (anti-islanding, 32/81/27/59) and utility coordination.", "Is utility-parallel operation required, or only islanded operation on the substation bus? Who owns the upstream service?"),
    ("PPI-17", "Part 3 2.4, 1.21", "Cybersecurity per DoDI 8510.01 (RMF); Modbus TCP SCADA integration.", "RMF package level of effort varies by 5–10×.", "Identify the SCADA system and owner, the required authorization (ATO/IATT/type authorization) and the ISSM. Is the Contractor responsible for the RMF package?"),
    ("PPI-18", "Part 3 1.34.2.6 Table 1.17.2.6.1", "Margin-test load list totals more than 7 MW starting kW against a 3.4 MW unit.", "Test basis undefined.", "Provide the intended margin-test load profile."),
    ("PPI-19", "Part 3 G4010 vs 1.12.15/1.20.3", "G4010 requires 4-wire multi-grounded wye; the GTG is high-resistance grounded through an NGR (3-wire).", "System grounding and relay scheme.", "Confirm existing 4.16 kV system grounding at Subs #1 and #6 and the intended GTG grounding method."),
    ("PPI-20", "Part 3 1.2.5, 1.2.4.1", "Black start from a 480 V source via ATS with a 50 ft cable; no site 480 V source is shown.", "New LV source/receptacle at each site.", "Identify the black-start power source at each substation."),
    ("PPI-21", "Sol. p.3 vs Encon ITB", "PPI due 1400 EST (Solicitation) vs 1100 EST (Encon ITB).", "Internal deadline.", "Internal only: use 1100 EST (Encon). Do not submit."),
    ("PPI-22", "Part 3 2.2, 01 33 29", "Sustainability notebook and HPSB checklist apply to buildings; this project has no building.", "Admin effort.", "Confirm UFGS 01 33 29 deliverables are limited or not applicable."),
    ("PPI-23", "Part 3 G10", "Wetlands statement contains an unresolved bracket: 'have [not] been identified'. E-101 shows a hatched water/wetland area west of Sub #1.", "Permitting (USACE/VMRC) and Chesapeake Bay Act RPA buffer.", "Confirm wetland/RPA status at each site; provide any delineation."),
    ("PPI-24", "Part 3 1.27", "Full CO2 discharge test at site per NFPA 12, for each unit.", "Recharge cost; life-safety controls during the test.", "Confirm a full discharge test is required for both units."),
]

# ════════════════════════════════════════════════════════════════════
# 2. DISCIPLINES (discipline, source, scope, RFP driver, seal)
# ════════════════════════════════════════════════════════════════════
DISCIPLINES = [
    ("Electrical (Lead DOR)", "BFA in-house", "Interconnection design (2 sites); generator breaker/bus differential retrofit; SLD/three-line; quick-connect JB; MV/LV cable and conduit; grounding; NGR/CT/PT sizing; short-circuit, coordination, arc flash, load-flow/voltage-drop studies; relay settings; SCADA/Modbus I/O integration; GTG UFGS edit; DOR review of OEM package.", "Part 3 D5090, G4010; Part 6 E-101–E-106; 01 33 00 1.1.a (DOR approval of GTG submittals)", "Yes — PE (VA)"),
    ("Mechanical", "BFA in-house", "DEF storage/supply/containment/level monitoring; fuel fill and appurtenances; secondary-containment sizing; off-skid fuel module, atomizing air and purge provisions; exhaust plume and clearances; piping test requirements.", "Sol. p.1 DEF scope; Part 3 1.18, 1.35, 1.34.11.6", "Yes — PE (VA)"),
    ("Plumbing", "BFA in-house", "Purge-water supply and backflow prevention (if site-supplied); drainage of containment.", "Part 3 1.2.4.3; Part 2 01 50 00 1.2.h (backflow)", "Shares Mechanical seal"),
    ("Controls / Cybersecurity (ICS-RMF)", "Subconsultant", "RMF categorization support, NIST SP 800-82 App. G control implementation, network architecture, eMASS artifacts, hardening of PLC/HMI/relays, Modbus TCP boundary protection.", "Part 3 2.4 (DoDI 8500.01/8510.01, NIST 800-82)", "No"),
    ("Fire Protection (QFPE)", "Subconsultant", "UFC 3-600-01 QFPE: NFPA 37 separation to Bldgs 115/248 (Sub #1) and 1500 (Sub #6); NFPA 12 CO2 package review; NFPA 30 tank review; FP submittal approval (DOR); NAVFAC FPE meeting.", "01 30 00 1.3 (FPE meeting); 01 33 00 1.1.a.1; Part 3 1.27", "Yes — FP PE"),
    ("Structural", "Subconsultant (or BFA in-house if available)", "Trailer and tank pad design; jacking-leg bearing; tie-downs for 120 mph wind and seismic; JB and DEF tank foundations; conduit supports; bollards.", "Part 3 1.18.1.1.1 (pad), 1.7.19 (seismic), 1.11.17 (wind)", "Yes — PE/SE"),
    ("Civil", "Subconsultant", "Site plans, grading/drainage, trailer delivery access (turning), pavement repair at Sub #1, ESC/CBPA compliance, limit of work, laydown.", "Part 3 G10/G1030; Part 4 G10", "Yes — PE"),
    ("Geotechnical", "Subconsultant", "Borings at both pads, bearing capacity, pavement section, groundwater; geotech report (DOR-approved submittal).", "Part 4 G10.3; 01 33 00 1.1.a.2", "Yes — PE (report)"),
    ("Land Surveying / SUE", "Subconsultant", "Topographic survey of both sites; utility designation (SUE QL-B) and verification.", "Part 3 G10 (topo required in drawings); 01 14 00 1.1", "Yes — LS"),
    ("Environmental / Air Permitting", "Subconsultant", "VDEQ air permit applicability and application support (NSPS KKKK, BACT); SPCC/40 CFR 112 and 9VAC25-91 AST compliance; PROD form; CBPA/wetland screen.", "Part 3 1.3.19, 1.18.1.1.1, 1.32; Part 2 01 30 00 1.6 (PROD)", "No"),
    ("Acoustical (optional)", "Subconsultant (optional)", "Site noise assessment at adjacent occupied buildings using OEM sound data.", "Part 3 1.31 (sound limits); not mandated", "No"),
]
NOT_REQUIRED = [
    ("Architecture", "Part 3 B10–C30 'Not Used'; no building."),
    ("HVAC / interior MEP", "D30 'Not Used'; enclosure ventilation is OEM package."),
    ("Landscape", "G20 'Not Used'."),
    ("Telecom / security", "D5030/G4030 'Not Used' (SCADA/Modbus link is under Electrical/Cyber)."),
    ("Commissioning Authority", "No CxA required by RFP; OEM commissioning and contractor testing per 1.33/1.34."),
    ("Cost estimating", "By Encon (DB Offeror)."),
    ("NETA acceptance testing", "Construction trade subcontract to Encon; BFA specifies."),
]

# ════════════════════════════════════════════════════════════════════
# 3. SCHEDULE (key, start, minutes, title, owner, source, hard, done)
#    Due dates only (user, 24 Nov 2026): no internal meetings or color-team reviews.
# ════════════════════════════════════════════════════════════════════
SCHEDULE = [
    ("5512e", "2026-11-09 10:00", 0, "SECNAV 5512 + docs to Encon (password-protected) — sent 1046", "BFA", ENCON_SRC, True, True),
    ("5512g", "2026-11-09 13:00", 0, "Base access forms to Gov (Brian Hornaman)", "Encon", "Sol. p.2", True, True),
    ("site", "2026-11-17 13:00", 180, "Government site visit — NSF Dahlgren (Subs #1 and #6) — attended (BFA, Kelly)", "Gov/Encon/BFA", "Sol. p.2; Encon ITB", True, True),
    ("ppie", "2026-11-23 12:00", 30, "BFA PPI list to Encon (Attachment C format) — sent Fri 20 Nov 1542", "BFA", "Internal → Encon", True, True),
    ("subq", "2026-11-30 17:00", 30, "A/E subconsultant questions due to BFA (feeds PPIs)", "Subs", "Internal", False, False),
    ("ppi1", "2026-12-02 11:00", 30, "PPI deadline — Encon (Pre-Bid RFIs due)", "Encon", ENCON_SRC, True, False),
    ("ppi2", "2026-12-02 14:00", 30, "PPI deadline — Government (Kayla Liverman)", "Encon", "Sol. p.3", True, False),
    ("subs", "2026-12-10 17:00", 30, "Subconsultant fee proposals due to BFA (3 business days before BFA proposal due)", "Subs", "Internal", True, False),
    ("due1", "2026-12-15 14:00", 30, "BFA PROPOSAL DUE TO ENCON (price per Attachment 1)", "BFA", ENCON_SRC, True, False),
    ("due2", "2026-12-16 14:00", 30, "ENCON PROPOSAL DUE TO NAVFAC (email, Kayla Liverman)", "Encon", "Sol. p.3", True, False),
    ("acc", "2027-04-15", 0, "Proposal acceptance period expires (120 days)", "—", "Sol. p.2", False, False),
]
POST_AWARD = [
    ("Award (A)", "A", "—"), ("Bonds to KO", "A+10 d", "Sol. p.2"), ("Certificates of insurance", "A+15 d", "01 30 00 1.2"),
    ("Commence work", "NTP+15 d", "52.211-10"), ("PAK / concept presentation; NAVFAC FPE meeting", "≤ A+21 d", "01 30 00 1.3"),
    ("Initial schedule of prices", "A+21 d", "01 33 00 1.3.k"), ("Design Development submittal (+PROD form)", "≈ A+75 d", "01 33 10 1.3.j; 01 30 00 1.6"),
    ("DD in-progress review meeting", "submittal + 21 d", "01 33 10 1.3.j"), ("Final Design submittal", "≈ A+135 d", "01 33 10 1.3.j"),
    ("Final review meeting / comment resolution", "submittal + 21 d", "01 33 10 1.3.j"), ("Released-for-construction set (signed and sealed)", "≈ A+170 d", "01 33 10 1.3.k"),
    ("Construction submittal reviews", "15 d each", "01 33 00 1.3.p"), ("O&M information", "BOD − 30 d", "01 33 00 1.3.m"),
    ("Completion", "A+24 mo", "52.211-10"),
]

# ════════════════════════════════════════════════════════════════════
# 4. SHEETS / SPECS / BOD / CALCS
# ════════════════════════════════════════════════════════════════════
SEALS = {
    "G": "Lead DOR — BFA Electrical PE", "V": "Land Surveyor (sub)", "C": "Civil PE (sub)", "S": "Structural PE (sub)",
    "M": "BFA Mechanical PE (M/P merged)", "FP": "QFPE (sub)", "E": "BFA Electrical PE",
}
# (sheet, title, series, tag, justification, driver)  — driver in COND_IN_EXP counts in expected
COND_IN_EXP = ("RD-1", "RD-2", "RD-3", "RD-7")
SHEETS = [
    ("G-001", "Cover sheet, location/vicinity maps, sheet index", "G", "vehicle", "FC 1-300-09N drawing set requirements; NAVFAC title block", ""),
    ("G-002", "General notes, abbreviations, code summary, criteria", "G", "vehicle", "01 33 10.05 20 1.2 criteria list; UFC 1-200-01", ""),
    ("V-101", "Topographic and utility survey — Substation #1", "V", "vehicle", "Part 3 G10: topo survey shall be included in design drawings", ""),
    ("V-102", "Topographic and utility survey — Substation #6", "V", "vehicle", "Part 3 G10", ""),
    ("C-001", "Civil general notes, legend, ESC narrative", "C", "vehicle", "Part 4 G10.1/3; UFGS 31 23 00.00 20", ""),
    ("C-101", "Site, grading and access plan — Substation #1 (asphalt, open)", "C", "scope", "Part 3 3.1, G1030 (grade to drain); trailer access", ""),
    ("C-102", "Site, grading and access plan — Substation #6 (concrete, fenced)", "C", "scope", "Part 3 3.1, G1030; fence/gate modifications for trailer entry", ""),
    ("C-501", "Civil details — pavement repair, bollards, ESC", "C", "design", "Asphalt cut/patch at Sub #1; bollard protection of JB", "RD-3"),
    ("S-001", "Structural general notes, design criteria (120 mph, CVSZ seismic)", "S", "vehicle", "Part 3 1.11.17, 1.7.19; UFC 3-301-01", ""),
    ("S-101", "Foundation plans — trailer/tank pads, JB and DEF pads (both sites)", "S", "scope", "Part 3 1.18.1.1.1 tank secured to a concrete pad", "RD-2"),
    ("S-501", "Foundation details — pads, jacking-leg bearing, tie-downs, anchorage", "S", "scope", "Part 3 1.7.19 seismic; 1.11.17 wind", ""),
    ("M-001", "Mechanical legend, notes, equipment schedules (DEF, pumps, tanks)", "M", "vehicle", "FC 1-300-09N; schedules for DEF equipment", ""),
    ("M-101", "Mechanical site plan — Substation #1 (DEF, fuel fill, purge/air provisions)", "M", "scope", "Sol. p.1 DEF storage/supply/containment; Part 3 1.2.4", "RD-7"),
    ("M-102", "Mechanical site plan — Substation #6", "M", "scope", "Sol. p.1 DEF; Part 3 1.2.4", "RD-7"),
    ("M-201", "Flow diagrams — DEF and off-skid fuel/air/purge systems", "M", "scope", "Part 3 1.18, 1.35 fuel controls; DEF level monitoring", ""),
    ("M-501", "Mechanical details — containment, tank, piping, supports", "M", "scope", "40 CFR 112 / 9VAC25-91 secondary containment; NFPA 30", ""),
    ("FP-001", "Fire protection code analysis — NFPA 37 separation, NFPA 12/30 summary", "FP", "vehicle", "UFC 3-600-01 QFPE; Part 3 1.27; FP is a DOR-approved submittal", ""),
    ("E-001", "Electrical legend, symbols, abbreviations, general notes", "E", "vehicle", "FC 1-300-09N", ""),
    ("E-002", "Conduit and cable schedules; quick-connect connector schedule", "E", "scope", "E-101/E-102 note 3; Part 3 1.2.6/1.2.7", ""),
    ("E-101", "Electrical site plan, new work — Substation #1", "E", "scope", "Part 6 E-101 notes 1–4", ""),
    ("E-102", "Electrical site plan, new work — Substation #6", "E", "scope", "Part 6 E-102 notes 1–4", ""),
    ("E-111", "Grounding plan — Substation #1", "E", "scope", "Part 3 1.34.11.6.2.b (≤25 Ω, IEEE 81); D5010", "RD-7"),
    ("E-112", "Grounding plan — Substation #6", "E", "scope", "Part 3 1.34.11.6.2.b", "RD-7"),
    ("E-201", "Single-line diagram — Substation #1 with GTG", "E", "scope", "Part 6 E-105", ""),
    ("E-202", "Single-line diagram — Substation #6 with GTG", "E", "scope", "Part 6 E-106", ""),
    ("E-203", "Three-line, metering and relaying diagram — Substation #1", "E", "scope", "E-105 87B bus differential, 7SJ80, SEL-700G", ""),
    ("E-204", "Three-line, metering and relaying diagram — Substation #6", "E", "scope", "E-106", ""),
    ("E-205", "GTG–switchgear interconnection and control block diagram", "E", "scope", "Part 3 1.21, 1.24; black start ATS 1.2.5", ""),
    ("E-206", "Control network / SCADA architecture and cyber boundary", "E", "scope", "Part 3 2.4 (NIST 800-82); 1.21 Modbus TCP", ""),
    ("E-207", "Switchgear elevations and modifications — Substation #1", "E", "scope", "E-103 'Feeder Future' cell conversion", ""),
    ("E-208", "Switchgear elevations and modifications — Substation #6", "E", "scope", "E-104; spare cell unconfirmed", "RD-1"),
    ("E-209", "Sequence of operations and I/O summary", "E", "scope", "Part 3 1.26 modes of operation; 1.21 I/O list for SCADA", ""),
    ("E-501", "Electrical details — JB, receptacles, conduit supports, penetrations", "E", "scope", "E-101 note 2/3; stainless hardware G4010", ""),
    ("E-502", "Grounding and NGR details", "E", "scope", "Part 3 1.20.3 NGR; IEEE 142", ""),
    ("E-601", "Protective device settings and coordination summary", "E", "scope", "Part 3 1.34.11.11.1 relay settings per coordination study", ""),
    ("E-602", "Short-circuit / arc-flash summary and label schedule", "E", "vehicle", "UFC 3-501-01, UFC 3-560-01, NFPA 70E", ""),
    # conditional (not in expected)
    ("C-111", "Erosion and sediment control plan — both sites", "C", "design", "VA ESC/CBPA: land disturbance ≥2,500 SF in a CBPA area", "RD-3+"),
    ("FP-101", "Fire protection site plans — separation distances, both sites", "FP", "design", "NFPA 37 separation to Bldgs 115/248 and 1500", "RD-5"),
    ("E-301", "Underground duct bank plan and sections (if underground route)", "E", "design", "Part 4 D5010.1 underground Sch 40 PVC", "RD-4"),
    ("E-503", "Utility-parallel interconnection protection details", "E", "design", "Part 3 1.3.17 grid-parallel operation", "RD-6"),
]
COUNT_RANGE = [("G", 2, 2, 3), ("V", 2, 2, 2), ("C", 3, 4, 5), ("S", 2, 3, 4), ("M", 4, 5, 5), ("FP", 1, 1, 2), ("E", 18, 19, 22)]
DRIVERS = [
    ("RD-1", "Does Sub #6 have a usable spare cubicle, or is a new switchgear section needed?", "E-208 (+1 new-section elevation)", "PPI-15 / site visit", "+1 / 0"),
    ("RD-2", "Concrete trailer pad required at both sites, or tank/JB pads only?", "S-101 splits per site (+1), S-501", "PPI-13", "+1 / −1"),
    ("RD-3", "Land disturbance ≥2,500 SF in CBPA (dedicated ESC plan)?", "C-111 (conditional, +1); C-501 (−1 if ESC rides on C-101/102)", "PPI-23 / civil", "+1 / −1"),
    ("RD-4", "Underground duct bank vs above-grade conduit?", "E-301 (conditional)", "PPI-14", "+1 / 0"),
    ("RD-5", "Do NFPA 37 separations need dedicated plans?", "FP-101 (conditional)", "QFPE at site visit", "+1 / 0"),
    ("RD-6", "Utility-parallel operation required?", "E-503 (conditional)", "PPI-16", "+1 / 0"),
    ("RD-7", "Separate per-site sheets vs combined plans", "M-101/102 (−1), E-111/112 (−1)", "Production judgement", "0 / −2"),
    ("RD-8", "Cover/general split into a third G sheet (criteria + PROD summary)", "G-003", "Production judgement", "+1 / 0"),
]
SPECS = [
    ("00 01 10", "Specification cover (sealed) and table of contents", "All", "Package", 4),
    ("01 31 23.13 20", "eCMS — Part 5 prescriptive, unmodified", "Div 01", "Front-end", 12),
    ("01 32 17.00 20", "Cost-loaded NAS — Part 5 prescriptive, unmodified", "Div 01", "Front-end", 30),
    ("01 33 29", "Sustainability Requirements and Reporting — Part 5 prescriptive, unmodified", "Div 01", "Front-end", 25),
    ("01 35 26", "Governmental Safety Requirements — Part 5 prescriptive, unmodified", "Div 01", "Front-end", 30),
    ("01 11 00", "Summary of Work", "Div 01", "Front-end", 4),
    ("01 14 00", "Work Restrictions (outage and tie-in sequencing)", "Div 01", "Front-end", 6),
    ("01 33 00.05 20", "Construction Submittal Procedures", "Div 01", "Front-end", 14),
    ("01 45 00.05 20", "Quality Control", "Div 01", "Front-end", 20),
    ("01 57 19", "Temporary Environmental Controls", "Div 01", "Front-end", 14),
    ("01 74 19", "Construction Waste Management and Disposal", "Div 01", "Front-end", 8),
    ("01 78 00", "Closeout Submittals", "Div 01", "Front-end", 8),
    ("01 78 23", "Operation and Maintenance Data", "Div 01", "Front-end", 16),
    ("03 30 00", "Cast-in-Place Concrete", "Structural", "Technical", 22),
    ("05 05 23.16", "Structural Welding (referenced by Part 3 1.7.15)", "Structural", "Technical", 10),
    ("05 50 13", "Miscellaneous Metal Fabrications (access platforms, bollards)", "Structural", "Technical", 8),
    ("13 48 73", "Seismic Control for Nonstructural Components", "Structural", "Technical", 10),
    ("31 00 00 / 31 23 00.00 20", "Earthwork; Excavation and Fill", "Civil", "Technical", 18),
    ("31 32 11", "Soil Surface Erosion Control (conditional, RD-3)", "Civil", "Technical", 6),
    ("32 12 16", "Asphalt Paving (pavement repair, Sub #1)", "Civil", "Technical", 10),
    ("22 05 00 / 22 11 00", "Common Work Results for Plumbing; Facility Water Distribution (purge supply, backflow)", "Plumbing", "Technical", 14),
    ("23 05 48.19", "Seismic Bracing for Mechanical Systems", "Mechanical", "Technical", 8),
    ("23 05 53", "Identification for Mechanical Piping and Equipment", "Mechanical", "Technical", 4),
    ("23 13 xx (custom)", "Diesel Exhaust Fluid Storage and Supply System", "Mechanical", "Technical", 10),
    ("33 56 10 (edited)", "Factory-Fabricated Fuel Storage Tanks — trailer tanks, appurtenances, tightness test", "Mechanical", "Technical", 16),
    ("40 05 13", "Common Work Results for Process Piping (fuel/DEF field piping)", "Mechanical", "Technical", 10),
    ("(within 26 32 15)", "CO2 total-flooding suppression — NFPA 12 requirements carried in GTG section", "Fire Protection", "Technical", 0),
    ("25 05 11", "Cybersecurity for Facility-Related Control Systems", "Cyber", "Technical", 22),
    ("26 05 00.00 40", "Common Work Results for Electrical", "Electrical", "Technical", 10),
    ("26 05 13.00 40", "Medium-Voltage Cables (quick-connect power cables, terminations)", "Electrical", "Technical", 14),
    ("26 05 19.00 10", "Insulated Wire and Cable (LV, control)", "Electrical", "Technical", 10),
    ("26 05 26", "Grounding and Bonding", "Electrical", "Technical", 8),
    ("26 05 33", "Raceways and Boxes (metal conduit, JB)", "Electrical", "Technical", 10),
    ("26 05 48", "Seismic Protection for Electrical Equipment", "Electrical", "Technical", 8),
    ("26 08 00", "Apparatus Inspection and Testing", "Electrical", "Technical", 22),
    ("26 13 xx (MV switchgear)", "Medium-Voltage Metal-Clad Switchgear — modifications and new breaker (UFGS number per DOR)", "Electrical", "Technical", 20),
    ("26 20 00", "Interior Distribution System (aux panelboard, black-start receptacle)", "Electrical", "Technical", 14),
    ("26 28 01.00 10", "Coordinated Power System Protection", "Electrical", "Technical", 12),
    ("26 32 15.00 (edited for GTG)", "Mobile Gas Turbine Generator Sets (Part 3 D5090 content; Part 4 D5090.3)", "Electrical", "Technical", 75),
    ("26 36 23", "Automatic Transfer Switches (black-start ATS, package)", "Electrical", "Technical", 14),
    ("26 41 00", "Lightning Protection (conditional)", "Electrical", "Technical", 8),
    ("33 71 02", "Underground Electrical Distribution (conditional, RD-4)", "Electrical", "Technical", 22),
    ("—", "Submittal register (SpecsIntact-generated)", "All", "Package", 12),
    ("—", "Annotated manufacturer product data (GTG, switchgear, relays, DEF)", "All", "Package", 120),
]
BOD = [
    ("General: project description, criteria, PROD summary, deviations", 4, 5),
    ("Civil (site, grading, ESC, CBPA)", 3, 3),
    ("Geotechnical summary (report appended)", 2, 2),
    ("Structural (pads, wind/seismic anchorage)", 3, 4),
    ("Mechanical (DEF, fuel, containment, purge/air)", 5, 6),
    ("Plumbing (purge water, backflow)", 1, 1),
    ("Fire protection (NFPA 37/12/30, QFPE)", 4, 4),
    ("Electrical (interconnection, protection, grounding, studies summary)", 10, 12),
    ("Controls and cybersecurity (SCADA, RMF approach)", 4, 5),
    ("Environmental and air permitting", 3, 3),
    ("Appendix: equipment data, PROD form, DD 1354 data", 6, 8),
]
CALCS = [
    ("Electrical", "Short-circuit study (both substations, GTG and utility sources)", 40),
    ("Electrical", "Protective device coordination study and relay settings (7SJ80, SEL-700G, 87B)", 60),
    ("Electrical", "Arc-flash hazard analysis and label data", 30),
    ("Electrical", "Load flow / voltage drop / cable sizing (MV and LV, black start)", 18),
    ("Electrical", "NGR, CT/PT and aux-transformer sizing", 10),
    ("Electrical", "Grounding (ground resistance, step/touch screening)", 10),
    ("Mechanical", "DEF consumption, storage and runtime; fuel runtime verification", 6),
    ("Mechanical", "Secondary containment volume (40 CFR 112, 110%) and piping sizing", 6),
    ("Plumbing", "Purge-water demand and backflow device sizing", 3),
    ("Structural", "Pad design, bearing, jacking legs, tie-downs (wind/seismic)", 30),
    ("Civil", "Drainage and ESC computations", 8),
    ("Fire Protection", "NFPA 37 separation and exposure analysis", 6),
]

# ════════════════════════════════════════════════════════════════════
# 5. MEP FEE — engineering hours per line; Tech = 25%, QC = 15% of Eng
# ════════════════════════════════════════════════════════════════════
TECH_PCT, QC_PCT = 0.25, 0.15
RATES = {"Engineer": 175.0, "Technician": 110.0, "QC Reviewer": 210.0}
_EHRS = {"E-001": 16, "E-002": 24, "E-101": 32, "E-102": 32, "E-111": 20, "E-112": 20, "E-201": 36, "E-202": 36,
         "E-203": 40, "E-204": 40, "E-205": 40, "E-206": 32, "E-207": 24, "E-208": 24, "E-209": 32, "E-501": 28,
         "E-502": 20, "E-601": 32, "E-602": 20, "G-001": 8, "G-002": 12}
_MHRS = {"M-001": 12, "M-101": 24, "M-102": 20, "M-201": 32, "M-501": 24}
FEE = []
for s, t, ser, *_ in SHEETS:
    if s in _EHRS:
        FEE.append(("Electrical", "Design Drawings", s, t + (" (lead DOR)" if ser == "G" else ""), _EHRS[s]))
    if s in _MHRS:
        FEE.append(("Mechanical", "Design Drawings", s, t, _MHRS[s]))
FEE += [
    ("Plumbing", "Design Drawings", "M-101/M-102/M-501", "Purge-water supply and backflow content on mechanical sheets", 16),
    ("Electrical", "Specifications", "26 32 15.00", "Mobile GTG section — heavy edit from Part 3 D5090 (largest section)", 80),
    ("Electrical", "Specifications", "26 13 xx", "MV metal-clad switchgear modifications and new breaker", 16),
    ("Electrical", "Specifications", "26 08 00 / 26 28 01", "Apparatus testing; coordinated power system protection", 20),
    ("Electrical", "Specifications", "26 05 xx (6 sections)", "Common work, MV cable, LV wire, grounding, raceways, seismic", 36),
    ("Electrical", "Specifications", "26 20 00 / 26 36 23 / 26 41 00", "Aux distribution, black-start ATS, lightning protection", 20),
    ("Electrical", "Specifications", "25 05 11", "Cybersecurity for FRCS — coordinate with cyber sub", 16),
    ("Electrical", "Specifications", "Div 01 + package", "Div 01 edits (lead DOR), sealed cover, TOC, SpecsIntact submittal register, product data", 32),
    ("Mechanical", "Specifications", "23 13 xx", "DEF storage and supply system (custom section)", 16),
    ("Mechanical", "Specifications", "33 56 10 / 40 05 13", "Trailer fuel tank appurtenances and tightness test; field process piping", 20),
    ("Mechanical", "Specifications", "23 05 48.19 / 23 05 53", "Seismic bracing; identification", 8),
    ("Plumbing", "Specifications", "22 05 00 / 22 11 00", "Purge-water piping and backflow preventer", 12),
    ("Electrical", "Calculations", "Short-circuit", "Short-circuit study — 2 substations, utility + GTG sources", 40),
    ("Electrical", "Calculations", "Coordination", "Protective device coordination and relay settings (7SJ80, SEL-700G, 87B)", 64),
    ("Electrical", "Calculations", "Arc flash", "Arc-flash analysis and labels", 32),
    ("Electrical", "Calculations", "Load flow / VD", "Load flow, voltage drop, MV/LV cable sizing, black start", 24),
    ("Electrical", "Calculations", "NGR / CT / PT", "NGR, instrument transformer and aux transformer sizing", 16),
    ("Electrical", "Calculations", "Grounding", "Ground resistance and step/touch screening", 16),
    ("Mechanical", "Calculations", "DEF / fuel", "DEF consumption and storage; fuel runtime verification", 14),
    ("Mechanical", "Calculations", "Containment / piping", "Secondary containment (110%) and piping sizing", 12),
    ("Plumbing", "Calculations", "Purge water", "Purge-water demand and backflow sizing", 6),
    ("Electrical", "Basis of Design", "BOD — Electrical", "Electrical + controls/cyber chapters (DD, updated at Final)", 48),
    ("Electrical", "Basis of Design", "BOD — General (lead DOR)", "General chapter, PROD summary, assembly of multi-discipline BOD", 16),
    ("Mechanical", "Basis of Design", "BOD — Mechanical", "Mechanical chapter (DD, updated at Final)", 20),
    ("Plumbing", "Basis of Design", "BOD — Plumbing", "Plumbing chapter", 6),
]

# ════════════════════════════════════════════════════════════════════
# 6. ROM (Attachment 1 aligned): (group, item, qty, unit, unit_cost, low_f, high_f, material share)
# ════════════════════════════════════════════════════════════════════
ROM = [
    ("A. GTG package (Attachment 1 items) — per unit", "Primary gas turbine with reduction gearbox and driveline", 2, "EA", 1_150_000, .80, 1.30, 1.00),
    ("A", "PLC controller for turbine — hardware, software, HMI", 2, "EA", 140_000, .80, 1.30, .70),
    ("A", "3.4 MW rated generator — 4,160 V", 2, "EA", 350_000, .80, 1.30, 1.00),
    ("A", "Switchgear, power conditioning, DECS-250, Woodward, Siemens, controls, cabinets, protection, NGR, aux xfmr", 2, "EA", 420_000, .80, 1.30, .85),
    ("A", "Digital fuel control system with nozzles, harness and integration", 2, "EA", 110_000, .80, 1.30, .90),
    ("A", "Exhaust system — emissions mitigation (silencer; SCR/DEF if required, PPI-07)", 2, "EA", 260_000, .60, 2.20, .85),
    ("A", "Trailer, fuel tanks (8 × 550 gal SS double-wall), integration", 2, "EA", 350_000, .80, 1.35, .80),
    ("A", "Custom enclosure (conex), partitions, air inlet, ventilation, F&G detection, CO2 suppression", 2, "EA", 300_000, .80, 1.30, .80),
    ("A", "Engineering and project management (OEM/packager)", 2, "EA", 160_000, .80, 1.30, 0),
    ("A", "Labor — genset construction, integration, fabrication, packaging", 2, "EA", 280_000, .80, 1.30, 0),
    ("A", "Load tests and FAT (150-start test, 22-h load), documentation, manuals", 2, "EA", 220_000, .55, 1.60, .10),
    ("A", "Commissioning of genset at site", 2, "EA", 100_000, .80, 1.40, .05),
    ("B. Electrical site work — per substation", "Generator breaker / retrofit in existing 4.16 kV switchgear incl. 87B CTs, relays, wiring", 2, "SITE", 220_000, .70, 1.70, .70),
    ("B", "Above-ground junction box with quick-connect MV receptacles and control connectors", 2, "SITE", 140_000, .80, 1.30, .80),
    ("B", "MV power and control cable in metal conduit (≈250 LF route)", 2, "SITE", 75_000, .70, 1.60, .60),
    ("B", "Grounding system and testing", 2, "SITE", 25_000, .80, 1.30, .50),
    ("B", "480 V black-start / aux source and receptacle", 2, "SITE", 30_000, .70, 1.60, .60),
    ("B", "SCADA / Modbus TCP integration and communications", 2, "SITE", 40_000, .70, 1.60, .50),
    ("B", "NETA acceptance testing; relay-setting implementation", 2, "SITE", 45_000, .80, 1.30, .05),
    ("C. Civil / structural / mechanical site work — per substation", "Survey layout, utility locates", 2, "SITE", 15_000, .80, 1.30, 0),
    ("C", "Concrete trailer/tank pad and jacking pads", 2, "SITE", 85_000, .50, 1.60, .45),
    ("C", "DEF storage tank, containment, piping, level monitoring, conduits", 2, "SITE", 45_000, .60, 1.60, .60),
    ("C", "Bollards, fence/gate modifications, asphalt repair, ESC", 2, "SITE", 30_000, .70, 1.50, .50),
    ("C", "Purge-water / atomizing-air provisions (PPI-02)", 2, "SITE", 15_000, .30, 2.00, .50),
    ("D. Project-wide", "Design services — all disciplines, DD/Final/RFC, design-phase QC (A/E)", 1, "LS", 550_000, .85, 1.30, 0),
    ("D", "Cybersecurity / RMF package and control hardening (PPI-17)", 1, "LS", 150_000, .50, 2.20, .10),
    ("D", "Air permitting support and emissions source testing (VDEQ)", 1, "LS", 90_000, .60, 1.80, 0),
    ("D", "Transportation, logistics, oversize permits (2 units)", 1, "LS", 80_000, .70, 1.50, 0),
    ("D", "Training, spare parts, special tools, extra filters, paint", 1, "LS", 120_000, .80, 1.30, .70),
    ("D", "GC field general conditions (PM, superintendent, QC manager, SSHO, eCMS, NAS) — 24 mo, part-time", 1, "LS", 250_000, .80, 1.40, .05),
    ("D", "Fuel for site testing; initial fill", 1, "LS", 40_000, .70, 1.50, 1.00),
]
TAX, OH, PROFIT, BOND = 0.053, 0.08, 0.06, 0.01

# ════════════════════════════════════════════════════════════════════
# 7. SUBMITTALS & MEETINGS
# ════════════════════════════════════════════════════════════════════
DESIGN_SUBMITTALS = [
    ("0", "Concept design presentation at PAK (≤ A+21 d)", "Concept SLDs, site layouts, approach; NAVFAC FPE meeting", "Acceptance at PAK", "01 30 00 1.3"),
    ("1", "Design Development (DD) submittal", "BOD, half-size drawings, online specs, calcs, DQC docs, PROD form, initial submittal register", "21-day Gov review + IPR meeting", "01 33 10 1.3.j; 01 30 00 1.6"),
    ("2", "Final Design submittal", "BOD, drawings, specs + cut sheets, PIF/bid schedule, calcs, DQC docs, register", "21-day Gov review + IPR meeting", "01 33 10 1.3.j"),
    ("3", "Released-for-construction (corrected final, signed and sealed)", "Backchecked final, comment responses closed", "Acceptance; DOR seal before construction", "01 33 10 1.3.k"),
    ("Opt.", "Early OEM/long-lead release package (if approved at PAK)", "GTG UFGS section + SLD for OEM order", "Gov acceptance at PAK", "01 33 10 1.3.k"),
]
MEETINGS = [
    ("Post-Award Kickoff (PAK) + concept presentation + initial partnering", 1, "≤ A+21 d", "Formal"),
    ("NAVFAC Fire Protection Engineer meeting (CO2/NFPA 12)", 1, "At PAK", "Formal"),
    ("DD in-progress review (IPR) meeting", 1, "DD + 21 d", "Formal"),
    ("Final design IPR / comment resolution meeting", 1, "Final + 21 d", "Formal"),
    ("RFC backcheck / comment closeout", 1, "≈ A+165 d", "Formal (virtual)"),
    ("Design coordination meetings with Gov (per QC Plan communication plan)", 6, "Biweekly during design", "Working"),
    ("Cybersecurity/RMF working sessions with ISSM", 2, "DD and Final", "Working"),
]
CONST_SUBS = [
    ("Div 01 — administrative (QC plan, APP/AHA, EPP, schedules, SOP, SDS, waste, sustainability, HPSB, DD1354, permits, O&M, record dwgs, warranties)", 26, 0, 26, "Contractor/Gov"),
    ("26 32 15 — GTG package (Part 3 1.6: SD 21, PD 36, DD 14, TR 10, Cert 13, MFR 3, O&M 1, CO 2)", 100, 31, 100, "Gov approval per 1.6; DOR approves shop dwgs, factory and site test reports"),
    ("26 13 xx — MV switchgear modifications / generator breaker", 8, 8, 2, "DOR"),
    ("26 05 13 / 26 05 19 — MV and LV cable, terminations, connectors", 6, 6, 0, "DOR"),
    ("26 05 26 / 26 05 33 — grounding, raceways, junction box", 6, 6, 0, "DOR"),
    ("26 08 00 / 26 28 01 — test procedures/reports, coordination study, relay settings", 8, 8, 2, "DOR"),
    ("26 20 00 / 26 36 23 / 26 41 00 — aux panel, ATS, lightning protection", 6, 6, 0, "DOR"),
    ("25 05 11 — cybersecurity artifacts (FRCS)", 8, 2, 8, "Gov/ISSM"),
    ("23 13 xx / 33 56 10 / 40 05 13 — DEF, tank appurtenances, piping, tightness test", 10, 10, 2, "DOR"),
    ("22 11 00 — purge water, backflow certification", 3, 3, 0, "DOR"),
    ("Seismic 13 48 73 / 23 05 48.19 / 26 05 48", 4, 4, 0, "DOR"),
    ("Fire protection (CO2 system data, F&G detection, discharge test)", 6, 6, 6, "DOR (QFPE) + Gov surveillance"),
    ("03 30 00 / 05 05 23.16 / 05 50 13 — concrete, welding, misc metals", 10, 10, 2, "DOR"),
    ("31 00 00 / 31 23 00.00 20 / 32 12 16 / 31 32 11 — earthwork, geotech, paving, ESC", 10, 10, 4, "DOR (geotech) + Gov surveillance"),
]
RESUBMIT = 0.20


# ════════════════════════════════════════════════════════════════════
# helpers
# ════════════════════════════════════════════════════════════════════
def md_table(headers, rows):
    out = ["| " + " | ".join(headers) + " |", "|" + "|".join("---" for _ in headers) + "|"]
    for r in rows:
        out.append("| " + " | ".join(str(c).replace("|", "/").replace("\n", " ") for c in r) + " |")
    return "\n".join(out)


def money(x):
    return f"${x:,.0f}"


HFILL = PatternFill("solid", fgColor="1F3864")
SUBFILL = PatternFill("solid", fgColor="D9E1F2")
TOTFILL = PatternFill("solid", fgColor="FFF2CC")
INFILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def sheet(wb, title, headers, rows, widths=None, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title[:31]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HFILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BOX
    for r in rows:
        ws.append(list(r))
    style_body(ws)
    for i, w in enumerate(widths or [], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    return ws


def style_body(ws):
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BOX


def bold_row(ws, r, fill=TOTFILL):
    for c in ws[r]:
        c.font = Font(bold=True)
        c.fill = fill
        c.border = BOX


# ════════════════════════════════════════════════════════════════════
# builders
# ════════════════════════════════════════════════════════════════════
def build_shred(A, Pb):
    md = [f"# RFP Shred — {TITLE}", "", HDR, "",
          "<!-- INTERNAL: requirement shred for the A/E scope. Sources: Solicitation (9 pp), SOW Parts 1–6 (210 pp incl. concept drawings E-101–E-108), Attachment 1, Price Schedule, WD VA20260007, PPI form, Encon ITB email 5 Nov 2026. -->",
          "## Key facts", "",
          md_table(["Item", "Value"], [
              ("Client / contracting office", "NAVFAC Washington — ROICC South Potomac, Dahlgren (KO POC: Kayla Liverman)"),
              ("Vehicle", "MACC N40080-24-D-0006 to -0013; FFP task order (fair opportunity, FAR 16.505)"),
              ("Delivery method", "Design-build; BFA = Designer of Record under Encon Corporation"),
              ("Evaluation", "Price only — lowest evaluated price; no technical factors scored"),
              ("Magnitude", "$5M–$10M (FAR 36.204)"),
              ("NAICS", "236220"),
              ("Performance period", "24 months from award, incl. 15 days for bonds/insurance; LDs $4,450/day"),
              ("Site visit", "Tue 17 Nov 2026, 1300 EST (single visit)"),
              ("PPIs", "Wed 2 Dec 2026, 1400 EST (Gov) / 1100 EST (Encon)"),
              ("Proposal due", "Encon → NAVFAC Wed 16 Dec 2026, 1400 EST; BFA → Encon Tue 15 Dec 2026, 1400"),
              ("Encon instructions", "Price breakdown per Attachment 1; documents on Encon SharePoint; 5512 by 9 Nov 1000"),
              ("Acceptance period", "120 days (to 15 Apr 2027)"),
              ("Scope", "2 × 3,400 kWe, 4,160 V trailer-mounted liquid-fuel mobile GTGs: Sub #1 and Sub #6; quick-connect interconnection to existing 4.16 kV switchgear; DEF system"),
              ("Design phases", "Concept at PAK → Design Development → Final → RFC (sealed)"),
              ("Gov RFP preparer", "Electrical only (K. D. Patel, PE); Civil/Struct/Mech/FP listed N/A"),
          ]), "",
          "## Requirements", "",
          "Responsible: **BFA** = A/E (DOR) · **GC** = Encon · **OEM** = GTG manufacturer/packager · **Sub** = BFA subconsultant.", "",
          md_table(["ID", "RFP ref", "Requirement", "Category", "Resp.", "A/E note"], SHRED), "",
          "## Conflicts, gaps and recommended PPIs", "",
          md_table(["ID", "RFP ref", "Issue", "Impact", "Proposed PPI question"], PPIS), ""]
    (A / f"{P}_01_RFP_Shred.md").write_text("\n".join(md), encoding="utf-8")
    wb = Workbook()
    sheet(wb, "Requirements", ["ID", "RFP Ref", "Requirement", "Category", "Responsible", "A/E Note", "Compliant? (Y/N)", "Owner Notes"],
          [r + ("", "") for r in SHRED], [8, 20, 70, 14, 12, 50, 12, 30], first=True)
    sheet(wb, "Conflicts & PPIs", ["ID", "RFP Ref", "Issue", "Impact", "Proposed PPI Question", "Submit? (Y/N)", "Gov Response"],
          [r + ("", "") for r in PPIS], [8, 24, 55, 35, 60, 10, 40])
    sheet(wb, "Attachment C format", ["Question No.", "Date Received", "Page", "Section", "Para.", "QUESTION", "GOVERNMENT RESPONSE"],
          [(i + 1, "", "", r[1], "", r[4], "") for i, r in enumerate(p for p in PPIS if p[0] != "PPI-21")], [10, 12, 8, 30, 8, 80, 40])
    wb.save(Pb / f"{P}_01_RFP_Shred.xlsx")


def build_disciplines(A, Pb):
    subs = [d for d in DISCIPLINES if d[1].startswith("Sub")]
    md = [f"# A/E Team — Subconsultant Disciplines — {TITLE}", "", HDR, "",
          "## Required disciplines", "",
          md_table(["Discipline", "Source", "Scope", "RFP driver", "Seal"], DISCIPLINES), "",
          "## Subconsultants to solicit (RFQ 10 Nov 2026)", "",
          "\n".join(f"{i+1}. **{d[0]}** — {d[3]}" for i, d in enumerate(subs)), "",
          "## Disciplines not required", "",
          md_table(["Discipline", "Basis"], NOT_REQUIRED), "",
          "## Notes", "",
          "- The RFP cover identifies only an Electrical preparer (K. D. Patel, PE). Civil, Structural, Mechanical and FP are marked N/A, but Part 3/4 requirements (topo survey, geotech, pads, DEF containment, CO2 suppression) force those disciplines into the DOR team.",
          "- GTG package internal design (turbine, generator, enclosure, CO2, F&G, controls) is by the OEM. BFA specifies, integrates and approves as DOR (01 33 00.05 20 1.1.a).",
          "- Structural and Civil may be self-performed if BFA has in-house staff; the list assumes subconsultants.", ""]
    (A / f"{P}_02_Subconsultant_Disciplines.md").write_text("\n".join(md), encoding="utf-8")
    wb = Workbook()
    sheet(wb, "Disciplines", ["Discipline", "Source", "Scope", "RFP Driver", "Seal Required", "Firm (TBD)", "RFQ Sent", "Fee Rec'd"],
          [d + ("", "", "") for d in DISCIPLINES], [26, 18, 70, 40, 18, 22, 10, 10], first=True)
    sheet(wb, "Not Required", ["Discipline", "Basis"], NOT_REQUIRED, [26, 80])
    wb.save(Pb / f"{P}_02_Subconsultant_Disciplines.xlsx")


def _dt(s):
    return datetime.strptime(s, "%Y-%m-%d %H:%M") if " " in s else datetime.strptime(s, "%Y-%m-%d")


def _esc(s):
    return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def _fold(line):
    if len(line.encode("utf-8")) <= 74:
        return [line]
    out, cur = [], ""
    for ch in line:
        if len((cur + ch).encode("utf-8")) > 73:
            out.append(cur)
            cur = " " + ch
        else:
            cur += ch
    out.append(cur)
    return out


VTZ = ["BEGIN:VTIMEZONE", "TZID:America/New_York", "BEGIN:DAYLIGHT", "TZOFFSETFROM:-0500", "TZOFFSETTO:-0400",
       "TZNAME:EDT", "DTSTART:19700308T020000", "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU", "END:DAYLIGHT",
       "BEGIN:STANDARD", "TZOFFSETFROM:-0400", "TZOFFSETTO:-0500", "TZNAME:EST", "DTSTART:19701101T020000",
       "RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU", "END:STANDARD", "END:VTIMEZONE"]


def _vevent(k, s, mins, t, owner, src, hard, stamp):
    d = _dt(s)
    loc = "NSF Dahlgren, VA (Substations #1 and #6)" if k == "site" else ("Teams / BFA office" if owner.startswith("BFA") else "")
    desc = f"{SOL}\nOwner: {owner}\nSource: {src}\nPursuit {P} — Encon (DB Offeror) / BFA (DOR)."
    if k == "site":
        desc += ("\nVerify: switchgear make/model and spare cells, nameplates, conduit route, pad/asphalt condition, "
                 "480 V black-start source, gate/fence access for a 48-52 ft trailer, distances to Bldgs 115/248 and 1500, "
                 "wetland edge at Sub #1. Coordinate photography with the KO.")
    ev = ["BEGIN:VEVENT", f"UID:{P}-{k}@bfa-eng.com", f"DTSTAMP:{stamp}"]
    if " " in s:
        end = d + timedelta(minutes=mins or 30)
        ev += [f"DTSTART;TZID=America/New_York:{d:%Y%m%dT%H%M%S}", f"DTEND;TZID=America/New_York:{end:%Y%m%dT%H%M%S}"]
    else:
        ev += [f"DTSTART;VALUE=DATE:{d:%Y%m%d}", f"DTEND;VALUE=DATE:{(d + timedelta(days=1)):%Y%m%d}", "TRANSP:TRANSPARENT"]
    ev += [f"SUMMARY:{_esc(('[HARD] ' if hard else '') + P + ' — ' + t)}", f"DESCRIPTION:{_esc(desc)}"]
    if loc:
        ev.append(f"LOCATION:{_esc(loc)}")
    ev += ["CATEGORIES:Proposal", "BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{_esc(t)}",
           "TRIGGER:-P1D" if hard else "TRIGGER:-PT30M", "END:VALARM"]
    if hard and " " in s:
        ev += ["BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{_esc(t)}", "TRIGGER:-PT2H", "END:VALARM"]
    ev.append("END:VEVENT")
    return ev


def _write_cal(path, events):
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//BFA//27001D Proposal//EN", "CALSCALE:GREGORIAN",
             "METHOD:PUBLISH", f"X-WR-CALNAME:{P} Substation Generators"] + VTZ
    for e in events:
        lines += e
    lines.append("END:VCALENDAR")
    out = []
    for ln in lines:
        out += _fold(ln)
    path.write_bytes(("\r\n".join(out) + "\r\n").encode("utf-8"))


def build_schedule(A, Pb):
    rows = []
    for k, s, mins, t, owner, src, hard, done in SCHEDULE:
        d = _dt(s)
        when = d.strftime("%a %d %b %Y") + (d.strftime(" %H%M") if " " in s else "")
        status = "Complete" if done else ("Past" if d.date() < SCHED_ASOF else "Open")
        rows.append((when, t, owner, src, "HARD" if hard else "", status))
    md = [f"# Proposal Schedule — {TITLE}", "", HDR.replace(f"{TODAY:%d %b %Y}", f"{SCHED_ASOF:%d %b %Y}"), "",
          f"All times Eastern. Due dates only. Sources: Solicitation (Sol.), {ENCON_SRC}, and BFA internal deadlines.", "",
          "## Pre-award milestones", "", md_table(["Date / time", "Milestone", "Owner", "Source", "Gate", "Status"], rows), "",
          "## Post-award contract durations (for the DB schedule)", "", md_table(["Event", "Timing", "RFP ref"], POST_AWARD), "",
          "## Notes", "",
          "- Federal holiday in the window: Thanksgiving 26 Nov (27 Nov office closure is typical).",
          "- Subconsultant fee proposals are due 3 business days before the BFA proposal is due to Encon.",
          "- The PPI deadline used is Encon's 1100 EST, which is earlier than the Government's 1400 EST.",
          "- Watch for amendments (SF-30, by email to MACC holders) after the PPI deadline. Re-check fee and scope before Gold team.",
          "- Calendar files are in `output/published/calendar/`. Import the combined `27001D_Proposal_Calendar` .ics into Outlook (double-click, or File → Open & Export → Import/Export → iCalendar). Single-event files are also provided.", ""]
    (A / f"{P}_03_Proposal_Schedule.md").write_text("\n".join(md), encoding="utf-8")
    wb = Workbook()
    sheet(wb, "Milestones", ["Date / Time (ET)", "Milestone", "Owner", "Source", "Gate", "Status"], rows, [24, 70, 18, 30, 8, 10], first=True)
    sheet(wb, "Post-Award", ["Event", "Timing", "RFP Ref"], POST_AWARD, [50, 20, 30])
    wb.save(Pb / f"{P}_03_Proposal_Schedule.xlsx")
    cal = Pb / "calendar"
    cal.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    events = []
    for k, s, mins, t, owner, src, hard, done in SCHEDULE:
        if done or k in ("rfp", "vet", "tgiv"):
            continue
        ev = _vevent(k, s, mins, t, owner, src, hard, stamp)
        events.append(ev)
        _write_cal(cal / f"{P}_{_dt(s):%Y%m%d}_{k}.ics", [ev])
    _write_cal(cal / f"{P}_Proposal_Calendar.ics", events)
    return len(events)


def build_counts(A, Pb):
    low = sum(r[1] for r in COUNT_RANGE); exp = sum(r[2] for r in COUNT_RANGE); high = sum(r[3] for r in COUNT_RANGE)
    listed_exp = {}
    for s, t, ser, tag, j, cond in SHEETS:
        if not cond or cond in COND_IN_EXP:
            listed_exp[ser] = listed_exp.get(ser, 0) + 1
    for ser, lo, ex, hi in COUNT_RANGE:  # integrity: sheet list must reconcile to expected
        assert listed_exp.get(ser, 0) == ex, (ser, listed_exp.get(ser), ex)
    spec_pages = sum(s[4] for s in SPECS)
    bod_dd = sum(b[1] for b in BOD); bod_f = sum(b[2] for b in BOD)
    calc_pages = sum(c[2] for c in CALCS)
    tagmap = {"vehicle": "vehicle_mandated", "scope": "scope_mandated", "design": "design_implied"}
    sheet_rows = []
    for s, t, ser, tag, j, cond in SHEETS:
        if cond and cond not in COND_IN_EXP:
            c = f"**Conditional — {cond.rstrip('+')}**"
        else:
            c = cond
        sheet_rows.append((s, t, ser, SEALS[ser], tagmap[tag], j, c, "Y" if ser in ("M", "E") else "N"))
    mep_exp = sum(r[2] for r in COUNT_RANGE if r[0] in ("M", "E"))
    md = [f"# Design Deliverable Count Estimate (Drawings, Specs, BOD, Calcs) — {TITLE}", "", HDR, "",
          "<!-- INTERNAL — estimating basis per 00_admin/drawing_set_and_spec_baseline.SCHEMA.md. Sheet list, ranges and drivers are internal; only the Summary table is publishable (SCHEMA §7). -->",
          "> **Agent-proposed ranges.** Per SCHEMA §8, the low/expected/high values and sheet justifications are a proposal for the responsible engineer to overwrite. The calibration ledger has no closed rows, so these ranges come from judgement, not data.", "",
          "## Summary (publishable)", "",
          md_table(["Deliverable", "Low", "Expected", "High", "Notes"], [
              ("Drawings (sheets, all disciplines)", low, exp, high, f"{len(SEALS)} seal holders; MEP (M+E) expected = {mep_exp}; P merged onto M sheets"),
              ("Specifications (pages)", round(spec_pages * .8), spec_pages, round(spec_pages * 1.25), f"{len(SPECS)} line items incl. 4 Part 5 prescriptive sections (unmodified) and product data; SpecsIntact"),
              ("Basis of Design (pages)", bod_dd, bod_f, round(bod_f * 1.3), "Low = DD edition; Expected = Final edition; FC 1-300-09N"),
              ("Calculations (pages)", round(calc_pages * .75), calc_pages, round(calc_pages * 1.35), "Electrical studies ≈ 75% of volume"),
          ]), "",
          "Expected drawings by series: " + ", ".join(f"{r[0]} {r[2]}" for r in COUNT_RANGE) + f" = **{exp}**. Spread {high/low:.2f}×. SCHEMA §6.2 flags anything under 1.5× before DD as possible false precision. It is accepted here, for the engineer to confirm, because the concept drawings E-101–E-108 already fix the SLD and site content. Expect the range to tighten after the site visit and PPI responses.", "",
          "## Seal floor and merge groups", "",
          md_table(["Series", "Seal holder", "Merge"], [(k, v, "P merged into M (same seal holder)" if k == "M" else "—") for k, v in SEALS.items()]), "",
          "Disciplines are kept separate because NAVFAC Part 2 requires each sheet to be sealed by the DOR registered in that discipline (a vehicle rule), not for legibility. BFA issues G, M and E, so the MEP subtotal is not the whole set. The non-MEP sheets (V, C, S, FP) are sealed by subconsultants.", "",
          "## Sheet list (internal)", "",
          md_table(["Sheet", "Title", "Series", "Seal holder", "Tag", "Justification", "Driver", "MEP"], sheet_rows), "",
          "## Count range by series (internal)", "",
          md_table(["Series", "Low", "Expected", "High"], COUNT_RANGE + [("**Total**", low, exp, high)]), "",
          "## Range drivers (internal)", "",
          md_table(["ID", "Driver", "Affects", "Resolved by", "Swing (+/−)"], DRIVERS), "",
          f"Reconciliation: the positive swings sum to +7, against high − expected = {high-exp}. The negative swings sum to −4, against expected − low = {exp-low}. **Dominant driver: RD-1 / PPI-15 (switchgear spare cell at Sub #6).** If no spare exists, a new switchgear section adds a sheet, ≈40 engineering hours and ≈$150–300k of construction cost. Settle it at the 17 Nov site visit.", "",
          "## Specifications (internal)", "",
          md_table(["Section", "Title", "Discipline", "Tier", "Est. pages"], SPECS + [("", "**Total**", "", "", spec_pages)]), "",
          "On-drawing specifications are not accepted (01 33 10.05 20 1.3.e), so specification effort is priced separately.", "",
          "## Basis of Design (internal)", "",
          md_table(["Chapter", "DD pages", "Final pages"], BOD + [("**Total**", bod_dd, bod_f)]), "",
          "## Calculations (internal)", "",
          md_table(["Discipline", "Calculation", "Est. pages"], CALCS + [("", "**Total**", calc_pages)]), ""]
    (A / f"{P}_07_Sheet_and_Page_Count.md").write_text("\n".join(md), encoding="utf-8")
    wb = Workbook()
    ws = sheet(wb, "Summary", ["Deliverable", "Low", "Expected", "High", "Notes"], [], [36, 10, 10, 10, 70], first=True)
    nr = len(COUNT_RANGE) + 1
    ws.append(["Drawings (sheets)", f"=SUM('Count Range'!B2:B{nr})", f"=SUM('Count Range'!C2:C{nr})", f"=SUM('Count Range'!D2:D{nr})", "All disciplines; see Sheet List"])
    ws.append(["Specifications (pages)", "=ROUND(C3*0.8,0)", f"=SUM(Specifications!E2:E{len(SPECS)+1})", "=ROUND(C3*1.25,0)", "SpecsIntact; Part 5 prescriptive unmodified"])
    ws.append(["Basis of Design (pages)", f"=SUM(BOD!B2:B{len(BOD)+1})", f"=SUM(BOD!C2:C{len(BOD)+1})", "=ROUND(C4*1.3,0)", "Low = DD edition; Expected = Final edition"])
    ws.append(["Calculations (pages)", "=ROUND(C5*0.75,0)", f"=SUM(Calculations!C2:C{len(CALCS)+1})", "=ROUND(C5*1.35,0)", ""])
    style_body(ws)
    sheet(wb, "Sheet List", ["Sheet", "Title", "Series", "Seal Holder", "Tag", "Justification", "Driver", "MEP"],
          [tuple(str(c).replace("**", "") for c in r) for r in sheet_rows], [9, 55, 7, 28, 16, 55, 18, 6])
    ws2 = sheet(wb, "Count Range", ["Series", "Low", "Expected", "High"], COUNT_RANGE, [10, 10, 10, 10])
    ws2.append(["Total", f"=SUM(B2:B{nr})", f"=SUM(C2:C{nr})", f"=SUM(D2:D{nr})"]); bold_row(ws2, ws2.max_row)
    sheet(wb, "Range Drivers", ["ID", "Driver", "Affects", "Resolved By", "Swing"], DRIVERS, [7, 60, 40, 22, 10])
    ws3 = sheet(wb, "Specifications", ["Section", "Title", "Discipline", "Tier", "Est. Pages"], SPECS, [24, 70, 16, 12, 10])
    ws3.append(["", "Total", "", "", f"=SUM(E2:E{len(SPECS)+1})"]); bold_row(ws3, ws3.max_row)
    ws4 = sheet(wb, "BOD", ["Chapter", "DD Pages", "Final Pages"], BOD, [60, 10, 10])
    ws4.append(["Total", f"=SUM(B2:B{len(BOD)+1})", f"=SUM(C2:C{len(BOD)+1})"]); bold_row(ws4, ws4.max_row)
    ws5 = sheet(wb, "Calculations", ["Discipline", "Calculation", "Est. Pages"], CALCS, [16, 70, 10])
    ws5.append(["", "Total", f"=SUM(C2:C{len(CALCS)+1})"]); bold_row(ws5, ws5.max_row)
    wb.save(Pb / f"{P}_07_Sheet_and_Page_Count.xlsx")
    return dict(low=low, exp=exp, high=high, spec=spec_pages, bod_dd=bod_dd, bod_f=bod_f, calc=calc_pages, mep=mep_exp)


def build_fee(A, Pb):
    rows, tot, by = [], {"eng": 0, "tech": 0, "qc": 0, "cost": 0}, {}
    for tr, dl, ref, desc, e in FEE:
        t, q = e * TECH_PCT, e * QC_PCT
        cost = e * RATES["Engineer"] + t * RATES["Technician"] + q * RATES["QC Reviewer"]
        rows.append((tr, dl, ref, desc, e, t, q, e + t + q, cost))
        for k, v in (("eng", e), ("tech", t), ("qc", q), ("cost", cost)):
            tot[k] += v
        b = by.setdefault((tr, dl), [0, 0, 0, 0])
        b[0] += e; b[1] += t; b[2] += q; b[3] += cost
    trades = ["Mechanical", "Plumbing", "Electrical"]
    dels = ["Design Drawings", "Specifications", "Calculations", "Basis of Design"]
    summ = []
    for tr in trades:
        for dl in dels:
            b = by.get((tr, dl), [0, 0, 0, 0])
            summ.append((tr, dl, f"{b[0]:.0f}", f"{b[1]:.1f}", f"{b[2]:.1f}", f"{sum(b[:3]):.1f}", money(b[3])))
    trade_tot = []
    for tr in trades:
        e = sum(by.get((tr, d), [0] * 4)[0] for d in dels)
        c = sum(by.get((tr, d), [0] * 4)[3] for d in dels)
        trade_tot.append((tr, f"{e:.0f}", f"{e*(1+TECH_PCT+QC_PCT):.1f}", money(c)))
    md = [f"# MEP Design Fee — Construction Documents — {TITLE}", "", HDR, "",
          "## Basis", "",
          "- Design deliverables only: Design Drawings, Specifications, Calculations, Basis of Design. Hours cover the DD, Final and RFC editions.",
          f"- Technician = **{TECH_PCT:.0%}** of engineering hours; QC = **{QC_PCT:.0%}** of engineering hours, on every line.",
          f"- Rates are placeholders to replace with BFA billing rates on the workbook 'Rates' tab: Engineer ${RATES['Engineer']:.0f}/h, Technician ${RATES['Technician']:.0f}/h, QC ${RATES['QC Reviewer']:.0f}/h. Every $ cell recalculates.",
          "- Electrical carries the lead-DOR G-series sheets and Div 01/package specification effort.",
          "- Excluded from these hours: meetings (PAK, IPRs, FPE, coordination), PM, site visits, construction-phase services (DOR submittal approval, RFIs, factory and site test witnessing, record CAD), subconsultant disciplines, RMF package, permitting.", "",
          "## Totals", "",
          md_table(["Engineering h", "Technician h", "QC h", "Total h", "Fee"], [(f"{tot['eng']:.0f}", f"{tot['tech']:.1f}", f"{tot['qc']:.1f}", f"{tot['eng']+tot['tech']+tot['qc']:.1f}", money(tot['cost']))]), "",
          "## By trade", "", md_table(["Trade", "Engineering h", "Total h", "Fee"], trade_tot), "",
          "## By trade and deliverable", "", md_table(["Trade", "Deliverable", "Eng h", "Tech h", "QC h", "Total h", "Fee"], summ), "",
          "## Line items", "",
          md_table(["Trade", "Deliverable", "Ref", "Description", "Eng h", "Tech h", "QC h", "Total h", "Fee"],
                   [(r[0], r[1], r[2], r[3], f"{r[4]:.0f}", f"{r[5]:.1f}", f"{r[6]:.1f}", f"{r[7]:.1f}", money(r[8])) for r in rows]), ""]
    (A / f"{P}_06_MEP_Design_Fee.md").write_text("\n".join(md), encoding="utf-8")
    wb = Workbook()
    ws = sheet(wb, "Rates", ["Parameter", "Value", "Note"], [], [30, 12, 60], first=True)
    for row in (["Engineer rate ($/h)", RATES["Engineer"], "Placeholder — replace with BFA rate"],
                ["Technician rate ($/h)", RATES["Technician"], "Placeholder — replace with BFA rate"],
                ["QC Reviewer rate ($/h)", RATES["QC Reviewer"], "Placeholder — replace with BFA rate"],
                ["Technician % of Engineering", TECH_PCT, "Per instruction"],
                ["QC % of Engineering", QC_PCT, "Per instruction"]):
        ws.append(row)
    style_body(ws)
    for r in range(2, 7):
        ws.cell(r, 2).fill = INFILL
    for r in (2, 3, 4):
        ws.cell(r, 2).number_format = "$#,##0.00"
    ws["B5"].number_format = ws["B6"].number_format = "0%"
    d = sheet(wb, "Fee Detail", ["Trade", "Deliverable", "Ref (Sheet/Section)", "Description", "Engineering Hrs", "Technician Hrs", "QC Hrs", "Total Hrs",
                                  "Engineering $", "Technician $", "QC $", "Total $"], [], [12, 16, 22, 55, 12, 12, 10, 10, 13, 13, 12, 13])
    for i, (tr, dl, ref, desc, e) in enumerate(FEE, start=2):
        d.append([tr, dl, ref, desc, e, f"=E{i}*Rates!$B$5", f"=E{i}*Rates!$B$6", f"=SUM(E{i}:G{i})",
                  f"=E{i}*Rates!$B$2", f"=F{i}*Rates!$B$3", f"=G{i}*Rates!$B$4", f"=SUM(I{i}:K{i})"])
    style_body(d)
    for i in range(2, len(FEE) + 2):
        d.cell(i, 5).fill = INFILL
    n = len(FEE) + 1
    d.append(["TOTAL", "", "", ""] + [f"=SUM({c}2:{c}{n})" for c in "EFGHIJKL"]); bold_row(d, d.max_row)
    for row in d.iter_rows(min_row=2):
        for c in row[4:8]:
            c.number_format = "0.0"
        for c in row[8:12]:
            c.number_format = "$#,##0"
    s = sheet(wb, "Summary", ["Trade", "Deliverable", "Engineering Hrs", "Technician Hrs", "QC Hrs", "Total Hrs", "Total $"], [], [20, 18, 14, 14, 10, 10, 14])
    r = 2
    for tr in trades:
        for dl in dels:
            s.append([tr, dl] + [f"=SUMIFS('Fee Detail'!{c}:{c},'Fee Detail'!$A:$A,$A{r},'Fee Detail'!$B:$B,$B{r})" for c in "EFGHL"])
            r += 1
        s.append([f"{tr} subtotal", ""] + [f"=SUMIFS('Fee Detail'!{c}:{c},'Fee Detail'!$A:$A,\"{tr}\")" for c in "EFGHL"])
        r += 1
    s.append(["GRAND TOTAL", ""] + [f"='Fee Detail'!{c}{n+1}" for c in "EFGHL"])
    style_body(s)
    for rr in range(2, s.max_row + 1):
        if str(s.cell(rr, 1).value).endswith("subtotal"):
            bold_row(s, rr, SUBFILL)
    bold_row(s, s.max_row)
    for row in s.iter_rows(min_row=2):
        for c in row[2:6]:
            c.number_format = "0.0"
        row[6].number_format = "$#,##0"
    wb.save(Pb / f"{P}_06_MEP_Design_Fee.xlsx")
    return tot, trade_tot


def build_rom(A, Pb):
    cur, rows = None, []
    base = {"low": 0, "exp": 0, "high": 0}; mat = {"low": 0, "exp": 0, "high": 0}; grp = {}
    for g, item, q, u, uc, lf, hf, m in ROM:
        if len(g) > 1:
            cur = g
        ext = q * uc
        for k, f in (("low", lf), ("exp", 1), ("high", hf)):
            base[k] += ext * f; mat[k] += ext * f * m
        rows.append((cur.split(".")[0], item, q, u, money(uc), money(ext * lf), money(ext), money(ext * hf), f"{m:.0%}"))
        grp[cur] = grp.get(cur, 0) + ext

    def roll(k):
        tax = mat[k] * TAX; sub = base[k] + tax; oh = sub * OH; pr = (sub + oh) * PROFIT; bond = (sub + oh + pr) * BOND
        return dict(base=base[k], tax=tax, sub=sub, oh=oh, pr=pr, bond=bond, total=sub + oh + pr + bond)
    L, E, H = roll("low"), roll("exp"), roll("high")
    summ = [(lbl, money(L[k]), money(E[k]), money(H[k])) for lbl, k in (
        ("Base cost", "base"), (f"Material sales tax ({TAX:.1%} × material share)", "tax"), ("Subtotal", "sub"),
        (f"Contractor's overhead ({OH:.0%})", "oh"), (f"Contractor's profit ({PROFIT:.0%})", "pr"), (f"Contractor's bond ({BOND:.0%})", "bond"),
        ("**Construction cost = Project cost (design included)**", "total"))]
    gsum = [(g.split(" — ")[0], money(v), f"{v/E['base']:.0%}") for g, v in grp.items()]
    md = [f"# ROM Construction Cost — {TITLE}", "", HDR, "",
          "> **Class 5 order-of-magnitude estimate (AACE 18R-97, −30%/+50%).** Concept-level judgement with no OEM or subcontractor quotes. Structured to Encon's Attachment 1 bid schedule so it maps to the GC price breakdown. Encon's OEM quotes supersede group A.", "",
          "## Summary", "", md_table(["", "Low", "Expected", "High"], summ), "",
          "## By group (expected)", "", md_table(["Group", "Base cost", "% of base"], gsum), "",
          f"**FAR 36.204 magnitude: $5M–$10M.** The expected value ({money(E['total'])}) is **{E['total']/10e6-1:.0%} above the upper bound**. The low case ({money(L['total'])}) is inside it. The OEM package (group A) is {list(grp.values())[0]/E['base']:.0%} of base cost, so OEM pricing decides where the bid lands.", "",
          "## Cost drivers above the Government range (PPI levers)", "",
          "1. **SCR/DEF** (PPI-07): emissions aftertreatment on a mobile GTG. The high case carries +120% on exhaust.",
          "2. **Factory test regime** (PPI-12): 150 start-and-load tests and a 22-h load test per unit.",
          "3. **Nuclear Class 1E criteria** (PPI-11): not priced. If enforced, this estimate is not valid.",
          "4. **RMF/ATO** (PPI-17): $75k–$330k depending on the authorization path.",
          "5. **Switchgear retrofit** (PPI-15): spare-cell retrofit vs a new section with OEM field service.", "",
          "## Detail (Attachment 1 aligned)", "",
          md_table(["Grp", "Item", "Qty", "Unit", "Unit cost", "Low ext.", "Expected ext.", "High ext.", "Material %"], rows), "",
          "## Basis and assumptions", "",
          "- 2026 dollars; Dahlgren, VA; Davis-Bacon WD VA20260007. No escalation to midpoint (fabrication falls in 2027).",
          "- One GTG per site at Subs #1 and #6 (PPI-01). Above-grade metal conduit route ≈250 LF per site (PPI-14).",
          "- New generator breaker in an existing spare cell; new bus differential CTs and relays (PPI-15).",
          "- No utility-parallel interconnection study or utility upgrades (PPI-16). No hazardous-materials abatement.",
          "- Design services carried at $550k for all disciplines: MEP from the fee workbook plus subconsultants and meetings. Construction-phase DOR services are excluded.",
          f"- Markups follow the Attachment 1 'Cost Estimate Summary' order: sales tax {TAX:.1%} on material share, OH {OH:.0%}, profit {PROFIT:.0%}, bond {BOND:.0%}. OH, profit and bond are placeholders for Encon to replace.", ""]
    (A / f"{P}_05_ROM_Construction_Cost.md").write_text("\n".join(md), encoding="utf-8")
    wb = Workbook()
    d = sheet(wb, "ROM Detail", ["Group", "Item (Attachment 1 aligned)", "Qty", "Unit", "Unit Cost", "Low Factor", "High Factor", "Material %",
                                 "Low Ext.", "Expected Ext.", "High Ext.", "Material $ (Exp.)"], [], [10, 70, 6, 7, 13, 8, 8, 9, 14, 14, 14, 14], first=True)
    cur = None
    for i, (g, item, q, u, uc, lf, hf, m) in enumerate(ROM, start=2):
        if len(g) > 1:
            cur = g
        d.append([cur.split(".")[0], item, q, u, uc, lf, hf, m, f"=C{i}*E{i}*F{i}", f"=C{i}*E{i}", f"=C{i}*E{i}*G{i}", f"=J{i}*H{i}"])
    style_body(d)
    for i in range(2, len(ROM) + 2):
        for c in (3, 5, 6, 7, 8):
            d.cell(i, c).fill = INFILL
    n = len(ROM) + 1
    d.append(["", "BASE COST", "", "", "", "", "", "", f"=SUM(I2:I{n})", f"=SUM(J2:J{n})", f"=SUM(K2:K{n})", f"=SUM(L2:L{n})"]); bold_row(d, d.max_row)
    d.append(["", "MATERIAL $ (by case)", "", "", "", "", "", "", f"=SUMPRODUCT(I2:I{n},H2:H{n})", f"=L{n+1}", f"=SUMPRODUCT(K2:K{n},H2:H{n})", ""]); bold_row(d, d.max_row, SUBFILL)
    for row in d.iter_rows(min_row=2):
        row[4].number_format = "$#,##0"; row[7].number_format = "0%"
        for c in row[8:12]:
            c.number_format = "$#,##0"
    s = sheet(wb, "Summary", ["Cost Estimate Summary (Attachment 1)", "Rate", "Low", "Expected", "High"], [], [48, 10, 16, 16, 16])
    B, M = n + 1, n + 2
    for row in (["Base Cost", "", f"='ROM Detail'!I{B}", f"='ROM Detail'!J{B}", f"='ROM Detail'!K{B}"],
                ["Material Sales Tax (on material share)", TAX, f"='ROM Detail'!I{M}*$B3", f"='ROM Detail'!J{M}*$B3", f"='ROM Detail'!K{M}*$B3"],
                ["Subtotal", "", "=C2+C3", "=D2+D3", "=E2+E3"],
                ["Contractor's Overhead", OH, "=C4*$B5", "=D4*$B5", "=E4*$B5"],
                ["Contractor's Profit", PROFIT, "=(C4+C5)*$B6", "=(D4+D5)*$B6", "=(E4+E5)*$B6"],
                ["Contractor's Bond", BOND, "=(C4+C5+C6)*$B7", "=(D4+D5+D6)*$B7", "=(E4+E5+E6)*$B7"],
                ["Construction Cost", "", "=C4+C5+C6+C7", "=D4+D5+D6+D7", "=E4+E5+E6+E7"],
                ["Project Cost (design included in base)", "", "=C8", "=D8", "=E8"],
                ["FAR 36.204 range — upper bound", "", 10_000_000, 10_000_000, 10_000_000],
                ["Variance vs upper bound", "", "=C9-C10", "=D9-D10", "=E9-E10"]):
        s.append(row)
    style_body(s)
    for r in (3, 5, 6, 7):
        s.cell(r, 2).fill = INFILL; s.cell(r, 2).number_format = "0.0%"
    for row in s.iter_rows(min_row=2):
        for c in row[2:5]:
            c.number_format = "$#,##0"
    bold_row(s, 8); bold_row(s, 9)
    sheet(wb, "Basis", ["Basis / Assumption"], [(x,) for x in [
        "Class 5 ROM (AACE 18R-97), −30%/+50%; concept-level; no quotes.",
        "Structured to Encon Attachment 1; OEM quotes from Encon supersede group A.",
        "Green cells are inputs. Low/High factors are per-line judgement.",
        "Excludes escalation, Class 1E qualification, utility upgrades, hazmat abatement, utility interconnection studies, construction-phase DOR services.",
        "See 27001D_01_RFP_Shred PPIs 07, 11, 12, 15, 16, 17 for the largest cost levers."]], [110])
    wb.save(Pb / f"{P}_05_ROM_Construction_Cost.xlsx")
    return L, E, H


def build_submittals(A, Pb):
    items = sum(c[1] for c in CONST_SUBS); dor = sum(c[2] for c in CONST_SUBS); gov = sum(c[3] for c in CONST_SUBS)
    pk = round(items / 2.6)
    formal = sum(m[1] for m in MEETINGS if m[3].startswith("Formal")); working = sum(m[1] for m in MEETINGS if m[3] == "Working")
    md = [f"# Design Submittals, Design Review Meetings and Construction Submittals — {TITLE}", "", HDR, "",
          "## Design submittals", "",
          "**2 formal review submittals (DD, Final), 3 design issues in total (DD, Final, RFC), plus a concept presentation at PAK.** An optional early OEM-release package can be added if agreed at PAK.", "",
          md_table(["#", "Submittal", "Content", "Review", "RFP ref"], DESIGN_SUBMITTALS), "",
          "## Design review meetings", "",
          f"**{formal} formal design review meetings** (PAK/concept, FPE, DD IPR, Final IPR, RFC closeout) **and {working} working coordination meetings.**", "",
          md_table(["Meeting", "No.", "Timing", "Type"], MEETINGS), "",
          "## Construction submittals and shop drawings (estimate)", "",
          f"**≈{items} submittal items** (≈{round(items*(1+RESUBMIT))} including {RESUBMIT:.0%} resubmittals) in about **{pk} transmittal packages**. The DOR approves ≈{dor} items; the Government approves or surveils ≈{gov}.", "",
          md_table(["Trade / spec", "Est. items", "DOR approval", "Gov approval / surveillance", "Approving authority"], CONST_SUBS + [("**Total**", items, dor, gov, "")]), "",
          "Notes: the GTG count follows the Part 3 D5090 1.6 enumeration, with one submittal per item covering both units. Unit-specific factory and site test reports may double ≈13 items. A 15-day Government review is assumed (01 33 00.05 20 1.3.p).", ""]
    (A / f"{P}_08_Submittals_and_Meetings.md").write_text("\n".join(md), encoding="utf-8")
    wb = Workbook()
    sheet(wb, "Design Submittals", ["#", "Submittal", "Content", "Review", "RFP Ref"], DESIGN_SUBMITTALS, [6, 50, 60, 30, 26], first=True)
    m = sheet(wb, "Design Review Meetings", ["Meeting", "No.", "Timing", "Type"], MEETINGS, [70, 6, 22, 16])
    m.append(["Total", f"=SUM(B2:B{len(MEETINGS)+1})", "", ""]); bold_row(m, m.max_row)
    c = sheet(wb, "Construction Submittals", ["Trade / Spec", "Est. Items", "DOR Approval", "Gov Approval/Surveillance", "Approving Authority"], CONST_SUBS, [80, 10, 12, 14, 36])
    n = len(CONST_SUBS) + 1
    c.append(["Total", f"=SUM(B2:B{n})", f"=SUM(C2:C{n})", f"=SUM(D2:D{n})", ""]); bold_row(c, c.max_row)
    c.append(["Incl. resubmittal allowance (20%)", f"=ROUND(B{n+1}*1.2,0)", "", "", ""]); style_body(c)
    wb.save(Pb / f"{P}_08_Submittals_and_Meetings.xlsx")
    return items, dor, gov, pk, formal, working


def to_docx(A, Pb):
    ref = REPO_ROOT / "templates" / "docx_reference" / "reference.docx"
    n = 0
    for md in sorted(A.glob("*.md")):
        out = Pb / (md.stem + ".docx")
        cmd = ["pandoc", str(md), "-f", "gfm", "-o", str(out)]
        if ref.exists():
            cmd += ["--reference-doc", str(ref)]
        subprocess.run(cmd, check=True)
        n += 1
    return n


def main():
    require_proposal_or_engine_root(REPO_ROOT)
    ap = argparse.ArgumentParser()
    ap.add_argument("--stage", default="_scratch/stage")
    ap.add_argument("--docx", action="store_true", help="only convert staged .md to .docx")
    ap.add_argument("--schedule-only", action="store_true", help="only build the proposal schedule (.md/.xlsx/.ics)")
    a = ap.parse_args()
    st = REPO_ROOT / a.stage
    A, Pb = st / "artifacts", st / "published"
    A.mkdir(parents=True, exist_ok=True); Pb.mkdir(parents=True, exist_ok=True)
    if a.docx:
        print(f"docx: {to_docx(A, Pb)}")
        return
    if a.schedule_only:
        print(f"ics events {build_schedule(A, Pb)}")
        return
    build_shred(A, Pb); build_disciplines(A, Pb)
    nev = build_schedule(A, Pb)
    cnt = build_counts(A, Pb)
    tot, trade_tot = build_fee(A, Pb)
    L, E, H = build_rom(A, Pb)
    subs = build_submittals(A, Pb)
    print(f"shred {len(SHRED)} reqs / {len(PPIS)} PPIs; ics events {nev}")
    print(f"counts {cnt}")
    print(f"fee {tot} {trade_tot}")
    print(f"ROM low {L['total']:,.0f} exp {E['total']:,.0f} high {H['total']:,.0f}; base exp {E['base']:,.0f}")
    print(f"subs items/dor/gov/packages/formal/working {subs}")


if __name__ == "__main__":
    main()
