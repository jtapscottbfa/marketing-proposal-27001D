#!/usr/bin/env python3
"""scripts/pursuit/build_field_packet.py — 27001D site-walk RFP print packet.

Extracts the RFP pages to carry on the site walk into one bookmarked PDF:
  * concept drawings E-101–E-108 (SOW pp 203–210, 24×36 in), scaled to
    11×17 landscape;
  * SOW text pages (8.5×11) used to back up the field questions.
Each page gets a small footer giving its source page, so marked-up
sheets can be traced back to the RFP.

    python scripts/pursuit/build_field_packet.py [--stage _scratch/stage_sitewalk]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts" / "lib"))
from guard_cwd import require_proposal_or_engine_root  # noqa: E402

import fitz  # noqa: E402

SOW = REPO_ROOT / "00_admin/rfp_intake/attachments/FINAL RFP FY27 CIP MINCON SUB STATION SOW (3).pdf"

# (bookmark, SOW page numbers (1-based), kind)
PACKET = [
    ("E-101 New site plan (Substation #1)", [203], "dwg"),
    ("E-102 New site plan (Substation #6)", [204], "dwg"),
    ("E-103 Switchgear general arrangement (Substation #1)", [205], "dwg"),
    ("E-104 Switchgear general arrangement (Substation #6)", [206], "dwg"),
    ("E-105 Electrical single-line diagram (Substation #1)", [207], "dwg"),
    ("E-106 Electrical single-line diagram (Substation #6)", [208], "dwg"),
    ("E-107 GTG with double drop-deck trailer", [209], "dwg"),
    ("E-108 GTG with double drop-deck trailer", [210], "dwg"),
    ("Part 3 §1.0–2.1 Project description, objectives, codes", [24], "txt"),
    ("Part 3 §2.4–6.0 Cybersecurity, site analysis, existing site conditions", [29], "txt"),
    ("Part 3 §1.11–1.12 Site conditions, principal design criteria", [44], "txt"),
    ("Part 5 §3.2–3.7 Utility outages, outage coordination, electrical", [195, 196, 197, 198, 199], "txt"),
]
TABLOID = (17 * 72, 11 * 72)


def stamp(page, text):
    r = page.rect
    page.insert_text((r.width - 4.4 * 72, r.height - 0.18 * 72), text, fontsize=7, fontname="helv", color=(0.35, 0.35, 0.35))


def build(out: Path) -> list[tuple[str, int, str]]:
    src = fitz.open(SOW)
    doc = fitz.open()
    toc, log = [], []
    for title, pages, kind in PACKET:
        toc.append([1, title, doc.page_count + 1])
        for p in pages:
            sp = src[p - 1]
            if kind == "dwg":
                page = doc.new_page(width=TABLOID[0], height=TABLOID[1])
                # sheet is drawn sideways on a 24×36 portrait page; rotate it to read landscape on 11×17
                page.show_pdf_page(page.rect, src, p - 1, rotate=90)
            else:
                doc.insert_pdf(src, from_page=p - 1, to_page=p - 1)
                page = doc[-1]
            note = " · HALF-SIZE: do not scale" if kind == "dwg" else ""
            stamp(page, f"27001D site-walk packet · RFP SOW p. {p} of {src.page_count}{note}")
            log.append((title, p, f"{page.rect.width / 72:.1f}x{page.rect.height / 72:.1f} in"))
    doc.set_toc(toc)
    doc.set_metadata({"title": "27001D Site Walk RFP Field Packet", "author": "Bowman, Foster & Associates"})
    doc.save(out, garbage=3, deflate=True)
    return log


def main():
    require_proposal_or_engine_root(REPO_ROOT)
    ap = argparse.ArgumentParser()
    ap.add_argument("--stage", default="_scratch/stage_sitewalk")
    a = ap.parse_args()
    out = REPO_ROOT / a.stage / "published" / "27001D_10_Site_Walk_RFP_Packet.pdf"
    out.parent.mkdir(parents=True, exist_ok=True)
    for row in build(out):
        print(*row, sep="  |  ")
    print("pages", fitz.open(out).page_count, "->", out)


if __name__ == "__main__":
    main()
