#!/usr/bin/env python3
"""scripts/pursuit/build_ppi_form.py — 27001D Pre-Proposal Inquiry (PPI) form.

Fills the Government's Pre-Proposal Inquiry Form (Attachment C,
00_admin/rfp_intake/attachments/"Pre- Proposal Inquiry Form (3).xlsx") with
BFA's PPIs for transmittal to Encon. The Government asked for an Excel file
with no Excel tables, so the form's own cell grid is filled in place.

Source of the questions: "Conflicts, gaps and recommended PPIs" in the RFP
shred (output/artifacts/27001D_01_RFP_Shred_vNN.md). PPI-21 is internal only
and is not included.

Reference columns: Page = printed page within the cited document/Part
(SOW Part 3 printed page = PDF page - 23; Part 4 = PDF page - 112).
"Date Received" and "Government Response" are left blank for the Government.

    python scripts/pursuit/build_ppi_form.py [--stage _scratch/stage_ppi]
"""
from __future__ import annotations

import argparse
import copy
import math
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts" / "lib"))
from guard_cwd import require_proposal_or_engine_root  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment  # noqa: E402

FORM = REPO_ROOT / "00_admin" / "rfp_intake" / "attachments" / "Pre- Proposal Inquiry Form (3).xlsx"
OUT_NAME = "27001D_11_PPI_Form.xlsx"
FIRST_ROW = 3

# (page, section, para, question) — shred PPI-01..24 order, PPI-21 omitted.
PPIS: list[tuple[str, str, str, str]] = [
    ("Sol. 1; P3-1",
     "Sol.; Part 3; Part 6",
     "Desc.; 1.0; E-101, E-102",
     "The Solicitation project description states the generators will support Substations #1, #1A, #4 "
     "and #6. Part 3 Section 1.0 and concept drawings E-101 and E-102 locate one GTG at Substation #1 "
     "and one GTG at Substation #6 only. Please confirm that generator interconnection is required only "
     "at Substations #1 and #6, and that no work is required at Substations #1A and #4."),
    ("9",
     "Part 3",
     "1.1.8; 1.2.4",
     "Paragraph 1.1.8 states the standby mobile GTG shall have four main connections. Paragraph 1.2.4 "
     "states there shall be six main connections and lists five (black start, output power, purge water, "
     "atomizing air, liquid fuel). Please confirm the required connections. Are the water-purge and "
     "atomizing-air supplies to be provided on board the GTG trailer, or must they be supplied from the "
     "site? Is compliance with ES 9-98 mandatory? If so, please provide a copy of the standard."),
    ("9; 23; 26; 31",
     "Part 3",
     "1.1.2; 1.12.15; 1.14; 1.17",
     "Paragraph 1.1.2 requires the GTG to be SoLoNOx equipped and capable of running on a very wide Wobbe "
     "range; these are gas-fuel combustion terms. Paragraph 1.14 describes a single-shaft engine, while "
     "the GTG Parameter Schedule (1.12.15) and paragraph 1.17 require a split shaft. Please confirm: "
     "(a) the GTG is liquid-fuel only (No. 2 fuel oil) with no gas-fuel capability required; "
     "(b) whether single-shaft or split-shaft is required, or whether either is acceptable; and "
     "(c) that equivalent products from other manufacturers meeting the stated performance requirements "
     "are acceptable."),
    ("4; 27",
     "Part 3",
     "2.1 (1.1.2, 1.1.9); 1.16.1",
     "Paragraph 1.16.1 requires a salient-pole synchronous generator per NEMA MG-1. The referenced "
     "standards also include ANSI C50.13, which covers cylindrical-rotor synchronous generators. Please "
     "confirm that NEMA MG-1 governs the generator design and that ANSI C50.13 does not apply."),
    ("Sol. 1; P3-22",
     "Sol.; Part 3",
     "Desc.; 1.12.15",
     "The Solicitation states a rating of 3,400 kWe per GTG. The GTG Parameter Schedule (1.12.15) lists "
     "4,434 kVA / 3,547 kW at 0.8 power factor. Please confirm the guaranteed output to be demonstrated in "
     "testing is 3,400 kWe net at the generator output terminals at the paragraph 1.3 rating conditions "
     "(59 deg F, 60% RH, No. 2 fuel oil)."),
    ("22-24",
     "Part 3",
     "1.12.15",
     "The GTG Parameter Schedule appears to have values offset from their labels (for example, "
     "'Voltage' shows 4434 kVA, 'Shaft speed' shows 60 Hz, and 'Heat rate' shows 101,880 lb/hr). "
     "Please issue a corrected GTG Parameter Schedule by amendment."),
    ("Sol. 1; P3-1, 12, 24",
     "Sol.; Part 3",
     "Desc.; 1.0; 1.3; 1.12.15",
     "The Solicitation and Part 3 Section 1.0 require a Diesel Exhaust Fluid (DEF) system, and the GTG "
     "Parameter Schedule requires the unit to be 'Tier IV compliant.' EPA Tier 4 standards apply to "
     "nonroad reciprocating diesel engines; for a combustion turbine, DEF is used only with selective "
     "catalytic reduction (SCR), and turbine emissions are addressed under 40 CFR 60 Subpart KKKK, which "
     "is also cited. Please confirm whether SCR is required. If it is, please state the required NOx "
     "emission limit (ppmvd at 15% O2 or lb/MWh). If it is not, please confirm the DEF system requirement "
     "is deleted. Please also identify who is responsible for obtaining any VDEQ air permit for the units."),
    ("13; 24; 32; 86",
     "Part 3",
     "1.3; 1.12.15; 1.18; 1.35",
     "Paragraph 1.35 requires eight 550-gallon fuel tanks per GTG (4,400 gallons total). The GTG "
     "Parameter Schedule lists full-load fuel consumption of 350 gal/hr, which gives about 12.6 hours of "
     "runtime, while a 14.6-hour full-load runtime is also required. Please confirm which governs: the "
     "tank quantity and size, or the 14.6-hour full-load runtime."),
    ("Sol. 1; P3-1, 23",
     "Sol.; Part 3; Part 6",
     "Desc.; 1.0; 1.12.15; E-107",
     "The Solicitation and Part 3 limit the trailer to 48 ft (L) x 8.5 ft (W) x 13.5 ft (H). Concept "
     "drawing E-107 shows a 32 ft deck with a 10 ft fuel tank module at each end, about 52 ft overall. "
     "Please confirm whether the 48 ft maximum length includes the fuel tank modules, or whether the "
     "E-107 configuration is acceptable."),
    ("12; 21",
     "Part 3",
     "1.3.20.3; 1.11.15",
     "Paragraph 1.3.20.3 gives a site altitude of 15 ft above MSL, and paragraph 1.11.15 gives a site "
     "elevation of 112 ft above MSL. Please confirm the design elevation to be used at each substation."),
    ("2-4; 21-22",
     "Part 3",
     "2.1 (1.0.7-1.0.14, 1.1.3); 1.12.14",
     "Paragraph 1.12.14 and the referenced standards (IEEE 308, 323, 344, 384, 387 and 603, and ASME "
     "NQA-1) contain nuclear power plant Class 1E criteria, including design-basis events, a 60-year "
     "qualified life, radiation dose and ECCS actuation. Please confirm these nuclear qualification "
     "requirements do not apply, and that the GTGs are to be designed, built and tested as commercial / "
     "industrial standby units."),
    ("75-78",
     "Part 3",
     "1.34.2.5; 1.34.2.7",
     "Paragraphs 1.34.2.5 through 1.34.2.7 require 150 consecutive start-and-load tests with no failures "
     "and a 22-hour continuous plus 2-hour 110% load capability test at the factory. Please confirm "
     "whether these tests are required on each production unit, or whether manufacturer prototype or "
     "type-test data for the same model may be submitted instead."),
    ("6; 32",
     "Part 3",
     "3.1; 1.18.1.1.1",
     "Paragraph 1.18.1.1.1 requires the fuel tank to be secured to a concrete pad large enough to support "
     "the generator. The fuel tanks are trailer-mounted, and paragraph 3.1 states Substation #1 is on "
     "asphalt and Substation #6 is on a concrete pad. Please confirm whether a new concrete pad is "
     "required at each substation. If the existing Substation #6 pad is to be used, please provide any "
     "available data on it (dimensions, thickness, reinforcement)."),
    ("P4-4; P6",
     "Part 4; Part 6",
     "D5010.1; E-101, E-102 Note 3",
     "Part 4 D5010.1 prohibits exposed exterior conduit and requires underground conduit in Schedule 40 "
     "PVC. Note 3 on concept drawings E-101 and E-102 requires metal conduit. Please confirm the required "
     "conduit route between each GTG connection point and the substation switchgear: above-grade rigid "
     "metal conduit on supports, or underground duct bank."),
    ("P6",
     "Part 6",
     "E-103 to E-106",
     "Drawings E-103 through E-106 show a new 1,200 A generator breaker and bus differential protection "
     "in the existing 4.16 kV switchgear. Substation #1 shows 'Feeder Future' cells; Substation #6 shows "
     "no spare cubicle. Please identify the existing switchgear manufacturer and model at each substation "
     "and the cubicle to be used for the generator breaker. If no spare cubicle exists at Substation #6, "
     "is a new switchgear section acceptable? Are the bus differential relays and CTs new work?"),
    ("11",
     "Part 3",
     "1.3.17",
     "Paragraph 1.3.17 requires operation synchronized to the utility grid and in parallel with the "
     "local utility. Please confirm whether utility-parallel operation is required, or only islanded "
     "operation on the substation bus. If parallel operation is required, please identify the owner of "
     "the upstream utility service and any utility interconnection requirements that apply."),
    ("6; 43",
     "Part 3",
     "2.4; 1.21",
     "Section 2.4 requires control systems to meet DoDI 8500.01 and DoDI 8510.01 (RMF), and Section 1.21 "
     "requires Modbus TCP integration with SCADA. Please identify the SCADA system the GTGs will connect "
     "to and its owner, the required authorization type (ATO, IATT or type authorization), and the "
     "Government ISSM point of contact. Is the Contractor responsible for preparing the full RMF "
     "authorization package?"),
    ("75-76",
     "Part 3",
     "1.34.2.6; Table 1.17.2.6.1",
     "The GTG Load List (Table 1.17.2.6.1) used for the margin tests in paragraph 1.34.2.6 totals more "
     "than 7 MW of starting load, against a 3.4 MW unit rating. Please provide the intended margin-test "
     "load profile."),
    ("22; 39; 89",
     "Part 3",
     "1.12.15; 1.20.3; G4010",
     "Section G4010 requires a 4-wire, multi-grounded wye primary distribution system. The GTG is "
     "specified as high-resistance grounded through a neutral grounding resistor (paragraphs 1.12.15 and "
     "1.20.3). Please confirm the existing 4.16 kV system grounding at Substations #1 and #6 and the "
     "intended GTG grounding method."),
    ("9-10",
     "Part 3",
     "1.2.4.1; 1.2.5",
     "Paragraphs 1.2.4.1 and 1.2.5 require black start from a 480 V source through an automatic transfer "
     "switch with a 50 ft cable. No 480 V source is shown at either substation. Please identify the "
     "black-start power source at each substation, or confirm that a new low-voltage source and "
     "connection point are to be provided by the Contractor."),
    ("5; P2-11, 12",
     "Part 3; Part 2",
     "2.2; 01 33 29",
     "Section 2.2 and UFGS 01 33 29 require a Sustainability eNotebook and an HPSB Checklist attached to "
     "the interim DD 1354. This project installs mobile generators and does not include a building. "
     "Please confirm whether these deliverables apply; if they do, please state which are required."),
    ("88; P6",
     "Part 3; Part 6",
     "G10; E-101",
     "Section G10 states that jurisdictional tidal and non-tidal wetlands 'have [not] been identified' on "
     "the project site; the bracketed text is unresolved. Drawing E-101 shows a hatched water / wetland "
     "area west of Substation #1. Please confirm the wetland and Chesapeake Bay Resource Protection Area "
     "status at each site, and provide any existing delineation."),
    ("70-71",
     "Part 3",
     "1.27.4",
     "Paragraph 1.27.4 requires a full CO2 discharge test at the site per NFPA 12. Please confirm a full "
     "discharge test is required for each of the two units, and whether the Government will accept a "
     "factory discharge test or a room-integrity test in lieu of a site discharge test."),
]

CHARS_PER_LINE = 52   # col F (40.7 wide) at Calibri 8
LINE_PT = 11.0


def row_height(text: str) -> float:
    lines = sum(max(1, math.ceil(len(p) / CHARS_PER_LINE)) for p in text.split("\n"))
    return max(30.0, lines * LINE_PT + 6)


def build(stage: Path) -> Path:
    wb = openpyxl.load_workbook(FORM)
    ws = wb.active
    if ws["F1"].value != "QUESTION" or ws["A1"].value != "Question No.":
        raise SystemExit("Form layout changed; header check failed.")
    if ws.tables:
        raise SystemExit("Form unexpectedly contains Excel tables.")

    # Clear stray content below the header (template has a lone ' ' in H14).
    for row in ws.iter_rows(min_row=FIRST_ROW):
        for c in row:
            if c.value is not None and str(c.value).strip() == "":
                c.value = None

    top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)
    top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for i, (page, section, para, question) in enumerate(PPIS):
        r = FIRST_ROW + i
        vals = {"A": i + 1, "C": page, "D": section, "E": para, "F": question}
        for col, v in vals.items():
            c = ws[f"{col}{r}"]
            c.value = v
            c.alignment = copy.copy(top_left if col == "F" else top_center)
        for col in "BG":
            ws[f"{col}{r}"].alignment = copy.copy(top_left)
        ws.row_dimensions[r].height = row_height(question)

    ws.print_area = f"A1:G{FIRST_ROW + len(PPIS) - 1}"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    out_dir = stage / "published"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / OUT_NAME
    wb.save(out)
    return out


def main() -> None:
    require_proposal_or_engine_root(REPO_ROOT)
    ap = argparse.ArgumentParser()
    ap.add_argument("--stage", default=str(REPO_ROOT / "_scratch" / "stage_ppi"))
    a = ap.parse_args()
    out = build(Path(a.stage))
    print(f"Wrote {out} ({len(PPIS)} PPIs)")


if __name__ == "__main__":
    main()
